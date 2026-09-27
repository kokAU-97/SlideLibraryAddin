/* My Slide Library — task pane logic
 * Stores slides (as standalone base64 .pptx files + thumbnail images) in
 * IndexedDB, grouped into collapsible categories, and inserts them after
 * the current slide using the PowerPoint JavaScript API.
 */

(function () {
  "use strict";

  var DB_NAME = "SlideLibraryDB";
  var DB_VERSION = 1;
  var STORE = "slides";
  var DEFAULT_CATEGORY = "Uncategorized";
  var PAGE_SIZE = 30;

  var dbPromise = null;
  var allSlides = [];       // in-memory cache, newest first
  var officeReady = false;
  var insertSupported = false;   // PowerPointApi 1.2 (insertSlidesFromBase64)
  var captureSupported = false;  // PowerPointApi 1.8 (exportAsBase64 / getImageAsBase64)
  var importCancelled = false;

  // categoryKey -> { fullList, grid (DOM el), sentinel (DOM el|null), flat (bool) }
  var sectionState = {};
  var categoryCollapsed = {};
  var lazyObserver = null;

  // ---------------------------------------------------------------------
  // IndexedDB helpers
  // ---------------------------------------------------------------------

  function openDB() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise(function (resolve, reject) {
      var req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = function () {
        var db = req.result;
        if (!db.objectStoreNames.contains(STORE)) {
          db.createObjectStore(STORE, { keyPath: "id" });
        }
      };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
    });
    return dbPromise;
  }

  function tx(mode) {
    return openDB().then(function (db) {
      return db.transaction(STORE, mode).objectStore(STORE);
    });
  }

  function dbGetAll() {
    return tx("readonly").then(function (store) {
      return new Promise(function (resolve, reject) {
        var req = store.getAll();
        req.onsuccess = function () { resolve(req.result); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function dbPut(record) {
    return tx("readwrite").then(function (store) {
      return new Promise(function (resolve, reject) {
        var req = store.put(record);
        req.onsuccess = function () { resolve(record); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function dbDelete(id) {
    return tx("readwrite").then(function (store) {
      return new Promise(function (resolve, reject) {
        var req = store.delete(id);
        req.onsuccess = function () { resolve(); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function dbClear() {
    return tx("readwrite").then(function (store) {
      return new Promise(function (resolve, reject) {
        var req = store.clear();
        req.onsuccess = function () { resolve(); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function makeId() {
    return "slide-" + Date.now() + "-" + Math.random().toString(36).slice(2, 9);
  }

  function escapeHtml(str) {
    var div = document.createElement("div");
    div.textContent = str == null ? "" : String(str);
    return div.innerHTML;
  }

  // ---------------------------------------------------------------------
  // Collapsed-category state (persisted per browser profile)
  // ---------------------------------------------------------------------

  function loadCollapsedState() {
    try {
      var raw = localStorage.getItem("slideLibCollapsed");
      return raw ? JSON.parse(raw) : {};
    } catch (e) {
      return {};
    }
  }

  function saveCollapsedState() {
    try {
      localStorage.setItem("slideLibCollapsed", JSON.stringify(categoryCollapsed));
    } catch (e) { /* ignore */ }
  }

  // ---------------------------------------------------------------------
  // DOM references
  // ---------------------------------------------------------------------

  var el = {};

  function cacheDom() {
    el.app = document.getElementById("app");
    el.loading = document.getElementById("loading");
    el.addBtn = document.getElementById("add-btn");
    el.addMenuPanel = document.getElementById("add-menu-panel");
    el.saveCurrentBtn = document.getElementById("save-current-btn");
    el.importAllBtn = document.getElementById("import-all-btn");
    el.fileInput = document.getElementById("file-input");
    el.unsupportedBanner = document.getElementById("unsupported-banner");
    el.importConfirm = document.getElementById("import-confirm");
    el.importConfirmText = document.getElementById("import-confirm-text");
    el.importCategoryInput = document.getElementById("import-category-input");
    el.importConfirmYes = document.getElementById("import-confirm-yes");
    el.importConfirmNo = document.getElementById("import-confirm-no");
    el.importProgress = document.getElementById("import-progress");
    el.importProgressText = document.getElementById("import-progress-text");
    el.importProgressBar = document.getElementById("import-progress-bar");
    el.importCancelBtn = document.getElementById("import-cancel-btn");
    el.searchInput = document.getElementById("search-input");
    el.clearAllBtn = document.getElementById("clear-all-btn");
    el.keepFormattingCheckbox = document.getElementById("keep-formatting-checkbox");
    el.status = document.getElementById("status");
    el.gridContainer = document.getElementById("grid-container");
    el.emptyState = document.getElementById("empty-state");
    el.noResults = document.getElementById("no-results");
    el.noResultsQuery = document.getElementById("no-results-query");
    el.categoryOptions = document.getElementById("category-options");
  }

  function setStatus(message) {
    el.status.textContent = message || "";
  }

  // ---------------------------------------------------------------------
  // Add menu (dropdown)
  // ---------------------------------------------------------------------

  function closeMenu() {
    el.addMenuPanel.hidden = true;
    el.addBtn.setAttribute("aria-expanded", "false");
  }

  function toggleMenu() {
    var willOpen = el.addMenuPanel.hidden;
    el.addMenuPanel.hidden = !willOpen;
    el.addBtn.setAttribute("aria-expanded", String(willOpen));
  }

  document.addEventListener("click", function (e) {
    if (!el.addMenuPanel || el.addMenuPanel.hidden) return;
    if (e.target === el.addBtn || el.addMenuPanel.contains(e.target)) return;
    closeMenu();
  });

  // ---------------------------------------------------------------------
  // Category datalist (autocomplete suggestions across all inputs)
  // ---------------------------------------------------------------------

  function updateCategoryDatalist() {
    var seen = {};
    allSlides.forEach(function (s) {
      var c = (s.category || "").trim();
      if (c && c !== DEFAULT_CATEGORY) seen[c] = true;
    });
    var options = Object.keys(seen).sort();
    el.categoryOptions.innerHTML = options
      .map(function (c) { return '<option value="' + escapeHtml(c) + '">'; })
      .join("");
  }

  // ---------------------------------------------------------------------
  // Rendering: grouped-by-category view, with lazy-loaded pages per group
  // ---------------------------------------------------------------------

  function currentQuery() {
    return (el.searchInput.value || "").trim().toLowerCase();
  }

  function groupByCategory(list) {
    var groups = {};
    list.forEach(function (s) {
      var cat = (s.category && s.category.trim()) || DEFAULT_CATEGORY;
      if (!groups[cat]) groups[cat] = [];
      groups[cat].push(s);
    });
    return groups;
  }

  function categorySortFn(a, b) {
    if (a === DEFAULT_CATEGORY) return 1;
    if (b === DEFAULT_CATEGORY) return -1;
    return a.localeCompare(b);
  }

  function getObserver() {
    if (!lazyObserver) {
      lazyObserver = new IntersectionObserver(
        function (entries) {
          entries.forEach(function (entry) {
            if (!entry.isIntersecting) return;
            var key = entry.target.dataset.key;
            lazyObserver.unobserve(entry.target);
            appendMore(key);
          });
        },
        { root: el.gridContainer, rootMargin: "150px" }
      );
    }
    return lazyObserver;
  }

  // Appends the next page of cards to an already-rendered section, in
  // place — this is what keeps scrolling smooth: we never clear/rebuild
  // the container while the user is mid-scroll, only add more below.
  function appendMore(key) {
    var state = sectionState[key];
    if (!state) return;
    if (state.sentinel && state.sentinel.parentNode) state.sentinel.remove();

    var rendered = state.grid.querySelectorAll(".slide-card").length;
    var next = state.fullList.slice(rendered, rendered + PAGE_SIZE);
    next.forEach(function (slide) {
      state.grid.appendChild(buildCard(slide, { showCategoryBadge: state.flat }));
    });

    var totalRendered = rendered + next.length;
    if (totalRendered < state.fullList.length) {
      var sentinel = document.createElement("div");
      sentinel.className = "scroll-sentinel";
      sentinel.dataset.key = key;
      state.grid.appendChild(sentinel);
      state.sentinel = sentinel;
      getObserver().observe(sentinel);
    } else {
      state.sentinel = null;
    }
  }

  function renderSection(key, label, list, opts) {
    opts = opts || {};
    var section = document.createElement("section");
    section.className = "category-section";

    if (!opts.flat) {
      var collapsed = !!categoryCollapsed[key];
      var header = document.createElement("button");
      header.type = "button";
      header.className = "category-header";
      header.innerHTML =
        '<svg class="chev" viewBox="0 0 16 16" width="12" height="12" fill="currentColor"><path d="' +
        (collapsed ? "M6 4l4 4-4 4V4z" : "M4 6l4 4 4-4H4z") +
        '"/></svg><span>' + escapeHtml(label) + '</span><span class="count">' + list.length + '</span>';
      header.addEventListener("click", function () {
        categoryCollapsed[key] = !categoryCollapsed[key];
        saveCollapsedState();
        render();
      });
      section.appendChild(header);
      if (collapsed) {
        el.gridContainer.appendChild(section);
        return;
      }
    }

    var grid = document.createElement("div");
    grid.className = "category-grid";
    section.appendChild(grid);
    el.gridContainer.appendChild(section);

    sectionState[key] = { fullList: list, grid: grid, sentinel: null, flat: !!opts.flat };

    var firstBatch = list.slice(0, PAGE_SIZE);
    firstBatch.forEach(function (slide) {
      grid.appendChild(buildCard(slide, { showCategoryBadge: !!opts.flat }));
    });

    if (list.length > PAGE_SIZE) {
      var sentinel = document.createElement("div");
      sentinel.className = "scroll-sentinel";
      sentinel.dataset.key = key;
      grid.appendChild(sentinel);
      sectionState[key].sentinel = sentinel;
      getObserver().observe(sentinel);
    }
  }

  // Full rebuild — used for deliberate actions (search, add, delete,
  // clear-all, category collapse/expand) where resetting to the top of
  // that view is expected. Never called from the scroll/lazy-load path.
  function render() {
    var query = currentQuery();
    el.gridContainer.innerHTML = "";
    sectionState = {};

    var hasAny = allSlides.length > 0;
    el.emptyState.hidden = hasAny;
    updateCategoryDatalist();

    if (query) {
      var matches = allSlides.filter(function (s) {
        return (
          s.name.toLowerCase().indexOf(query) !== -1 ||
          (s.category || "").toLowerCase().indexOf(query) !== -1
        );
      });
      el.noResults.hidden = !(hasAny && matches.length === 0);
      if (!el.noResults.hidden) el.noResultsQuery.textContent = query;
      if (matches.length) {
        renderSection("__search__", "Search results", matches, { flat: true });
      }
      return;
    }

    el.noResults.hidden = true;
    if (!hasAny) return;

    var groups = groupByCategory(allSlides);
    Object.keys(groups)
      .sort(categorySortFn)
      .forEach(function (key) {
        renderSection(key, key, groups[key], { flat: false });
      });
  }

  function buildCard(slide, opts) {
    opts = opts || {};
    var card = document.createElement("div");
    card.className = "slide-card";
    card.tabIndex = 0;
    card.setAttribute("role", "button");
    card.title = "Insert \"" + slide.name + "\" after the current slide";

    var thumb = document.createElement("div");
    thumb.className = "slide-thumb";
    if (slide.thumbnail) {
      var img = document.createElement("img");
      img.src = slide.thumbnail;
      img.alt = slide.name;
      thumb.appendChild(img);
    } else {
      thumb.classList.add("placeholder");
      thumb.innerHTML =
        '<svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="currentColor" stroke-width="1.6">' +
        '<rect x="2.5" y="5" width="19" height="14" rx="1.5"/><path d="M2.5 9h19"/></svg>';
    }
    if (opts.showCategoryBadge) {
      var badge = document.createElement("span");
      badge.className = "category-badge";
      badge.textContent = slide.category || DEFAULT_CATEGORY;
      thumb.appendChild(badge);
    }
    var del = document.createElement("button");
    del.className = "delete-btn";
    del.type = "button";
    del.setAttribute("aria-label", "Delete " + slide.name);
    del.textContent = "\u00D7";
    del.addEventListener("click", function (e) {
      e.stopPropagation();
      removeSlide(slide.id);
    });
    thumb.appendChild(del);
    card.appendChild(thumb);

    var nameInput = document.createElement("input");
    nameInput.className = "slide-name";
    nameInput.value = slide.name;
    nameInput.setAttribute("aria-label", "Slide name");
    nameInput.addEventListener("keydown", function (e) {
      if (e.key === "Enter") nameInput.blur();
    });
    nameInput.addEventListener("change", function () {
      var newName = nameInput.value.trim() || slide.name;
      nameInput.value = newName;
      slide.name = newName;
      dbPut(slide).catch(function () {});
    });
    card.appendChild(nameInput);

    var catInput = document.createElement("input");
    catInput.className = "slide-category";
    catInput.setAttribute("list", "category-options");
    catInput.setAttribute("aria-label", "Category");
    catInput.placeholder = "Category";
    catInput.value = slide.category && slide.category !== DEFAULT_CATEGORY ? slide.category : "";
    catInput.addEventListener("keydown", function (e) {
      if (e.key === "Enter") catInput.blur();
    });
    catInput.addEventListener("change", function () {
      var val = catInput.value.trim();
      slide.category = val || DEFAULT_CATEGORY;
      dbPut(slide)
        .then(function () { render(); })
        .catch(function () {});
    });
    card.appendChild(catInput);

    function activate(e) {
      if (e.target.tagName === "INPUT" || e.target === del) return;
      insertSlide(card, slide);
    }
    card.addEventListener("click", activate);
    card.addEventListener("keydown", function (e) {
      if ((e.key === "Enter" || e.key === " ") && e.target === card) {
        e.preventDefault();
        insertSlide(card, slide);
      }
    });

    return card;
  }

  function refresh() {
    return dbGetAll().then(function (records) {
      records.sort(function (a, b) { return b.createdAt - a.createdAt; });
      allSlides = records;
      render();
    });
  }

  function removeSlide(id) {
    dbDelete(id)
      .then(function () {
        setStatus("Slide deleted.");
        return refresh();
      })
      .catch(function (err) {
        console.error(err);
        setStatus("Couldn't delete that slide.");
      });
  }

  // ---------------------------------------------------------------------
  // Clear all (two-click confirm — native confirm() is unreliable in
  // Office task panes)
  // ---------------------------------------------------------------------

  var clearAllPendingTimer = null;

  function resetClearAllButton() {
    if (clearAllPendingTimer) {
      clearTimeout(clearAllPendingTimer);
      clearAllPendingTimer = null;
    }
    el.clearAllBtn.textContent = "Clear all";
    el.clearAllBtn.classList.remove("btn-confirm");
  }

  function clearAll() {
    if (!allSlides.length) return;

    if (!clearAllPendingTimer) {
      el.clearAllBtn.textContent = "Click again to confirm";
      el.clearAllBtn.classList.add("btn-confirm");
      clearAllPendingTimer = setTimeout(resetClearAllButton, 4000);
      return;
    }

    resetClearAllButton();
    dbClear()
      .then(function () {
        setStatus("Library cleared.");
        return refresh();
      })
      .catch(function (err) {
        console.error(err);
        setStatus("Couldn't clear the library.");
      });
  }

  // ---------------------------------------------------------------------
  // Save current slide
  // ---------------------------------------------------------------------

  function saveCurrentSlide() {
    closeMenu();
    if (!captureSupported) {
      setStatus("Your PowerPoint version can't capture slides directly.");
      return;
    }
    setStatus("Saving current slide…");

    PowerPoint.run(function (context) {
      var slide = context.presentation.getSelectedSlides().getItemAt(0);
      var exported = slide.exportAsBase64();
      var image = slide.getImageAsBase64({ width: 320 });
      return context.sync().then(function () {
        return dbPut({
          id: makeId(),
          name: "Slide " + (allSlides.length + 1),
          category: DEFAULT_CATEGORY,
          thumbnail: "data:image/png;base64," + image.value,
          slideBase64: exported.value,
          createdAt: Date.now()
        });
      });
    })
      .then(function () {
        setStatus("Slide saved to your library.");
        return refresh();
      })
      .catch(function (err) {
        console.error(err);
        setStatus("Couldn't save that slide: " + (err.message || err));
      });
  }

  // ---------------------------------------------------------------------
  // Import all slides from the currently open presentation
  // ---------------------------------------------------------------------

  function promptImportAll() {
    closeMenu();
    if (!captureSupported) {
      setStatus("Your PowerPoint version can't capture slides directly.");
      return;
    }
    PowerPoint.run(function (context) {
      context.presentation.slides.load("items");
      return context.sync().then(function () {
        return context.presentation.slides.items.length;
      });
    })
      .then(function (count) {
        if (!count) {
          setStatus("This presentation has no slides.");
          return;
        }
        el.importConfirmText.textContent =
          "Import all " + count + " slides from this presentation as separate library items?";
        el.importCategoryInput.value = "";
        el.importConfirm.hidden = false;
        el.importConfirm.dataset.count = String(count);
      })
      .catch(function (err) {
        console.error(err);
        setStatus("Couldn't read this presentation's slides.");
      });
  }

  function runImportAll() {
    var count = parseInt(el.importConfirm.dataset.count, 10) || 0;
    var category = el.importCategoryInput.value.trim() || DEFAULT_CATEGORY;
    el.importConfirm.hidden = true;
    if (!count) return;

    importCancelled = false;
    el.importProgress.hidden = false;
    el.importProgressBar.style.width = "0%";
    el.importProgressText.textContent = "Importing slide 1 of " + count + "…";

    var imported = 0;
    var i = 0;

    function step() {
      if (importCancelled || i >= count) {
        finish();
        return;
      }
      var index = i;
      PowerPoint.run(function (context) {
        var slide = context.presentation.slides.getItemAt(index);
        var exported = slide.exportAsBase64();
        var image = slide.getImageAsBase64({ width: 320 });
        return context.sync().then(function () {
          return dbPut({
            id: makeId(),
            name: "Slide " + (index + 1),
            category: category,
            thumbnail: "data:image/png;base64," + image.value,
            slideBase64: exported.value,
            createdAt: Date.now() + index // keep stable relative order
          });
        });
      })
        .then(function () {
          imported++;
        })
        .catch(function (err) {
          console.warn("Skipped slide " + (index + 1) + ":", err);
        })
        .then(function () {
          i++;
          el.importProgressBar.style.width = Math.round((i / count) * 100) + "%";
          el.importProgressText.textContent = "Importing slide " + Math.min(i + 1, count) + " of " + count + "…";
          setTimeout(step, 0);
        });
    }

    function finish() {
      el.importProgress.hidden = true;
      setStatus(
        importCancelled
          ? "Import cancelled after " + imported + " slide" + (imported === 1 ? "" : "s") + "."
          : "Imported " + imported + " of " + count + " slides."
      );
      refresh();
    }

    step();
  }

  // ---------------------------------------------------------------------
  // Upload pre-made .pptx files (fallback path, works without capture support)
  // ---------------------------------------------------------------------

  function readFileAsDataUrl(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function () { resolve(reader.result); };
      reader.onerror = function () { reject(reader.error); };
      reader.readAsDataURL(file);
    });
  }

  function niceName(fileName) {
    var withoutExt = fileName.replace(/\.[^/.]+$/, "");
    return withoutExt.replace(/[-_]+/g, " ").trim() || fileName;
  }

  function handleFiles(fileList) {
    closeMenu();
    var files = Array.prototype.slice.call(fileList || []).filter(function (f) {
      return /\.pptx$/i.test(f.name);
    });
    if (!files.length) {
      setStatus("Please choose .pptx files.");
      return;
    }
    setStatus("Adding " + files.length + " slide file" + (files.length > 1 ? "s" : "") + "…");

    Promise.all(
      files.map(function (file) {
        return readFileAsDataUrl(file).then(function (dataUrl) {
          var base64 = dataUrl.split(",")[1] || "";
          return dbPut({
            id: makeId(),
            name: niceName(file.name),
            category: DEFAULT_CATEGORY,
            thumbnail: null,
            slideBase64: base64,
            createdAt: Date.now()
          });
        });
      })
    )
      .then(function (added) {
        setStatus(added.length + " slide" + (added.length > 1 ? "s" : "") + " added.");
        return refresh();
      })
      .catch(function (err) {
        console.error(err);
        setStatus("Something went wrong while adding those files.");
      });
  }

  // ---------------------------------------------------------------------
  // Insert into the presentation
  // ---------------------------------------------------------------------

  function insertSlide(cardEl, slide) {
    if (!officeReady || !insertSupported) {
      setStatus("Open this add-in inside PowerPoint to insert slides.");
      return;
    }

    cardEl.classList.add("inserting");
    setStatus("Inserting \"" + slide.name + "\"…");

    var keepFormatting = el.keepFormattingCheckbox.checked;

    PowerPoint.run(function (context) {
      var selected = context.presentation.getSelectedSlides().getItemAt(0);
      selected.load("id");
      return context.sync().then(function () {
        context.presentation.insertSlidesFromBase64(slide.slideBase64, {
          formatting: keepFormatting ? "KeepSourceFormatting" : "UseDestinationTheme",
          targetSlideId: selected.id
        });
        return context.sync();
      });
    })
      .then(function () {
        cardEl.classList.remove("inserting");
        setStatus("Inserted \"" + slide.name + "\".");
      })
      .catch(function (err) {
        cardEl.classList.remove("inserting");
        console.error(err);
        setStatus("Couldn't insert that slide: " + (err.message || err));
      });
  }

  // ---------------------------------------------------------------------
  // Wiring
  // ---------------------------------------------------------------------

  var domReady = false;
  var capabilitiesKnown = false;

  function applyCapabilityUI() {
    if (!domReady || !capabilitiesKnown) return;
    if (!captureSupported) {
      el.unsupportedBanner.hidden = false;
      el.saveCurrentBtn.disabled = true;
      el.importAllBtn.disabled = true;
    }
  }

  function init() {
    cacheDom();
    categoryCollapsed = loadCollapsedState();

    el.addBtn.addEventListener("click", function (e) {
      e.stopPropagation();
      toggleMenu();
    });
    el.saveCurrentBtn.addEventListener("click", saveCurrentSlide);
    el.importAllBtn.addEventListener("click", promptImportAll);
    el.fileInput.addEventListener("change", function (e) {
      handleFiles(e.target.files);
      e.target.value = "";
    });

    el.importConfirmYes.addEventListener("click", runImportAll);
    el.importConfirmNo.addEventListener("click", function () {
      el.importConfirm.hidden = true;
    });
    el.importCancelBtn.addEventListener("click", function () {
      importCancelled = true;
    });

    el.searchInput.addEventListener("input", render);
    el.clearAllBtn.addEventListener("click", clearAll);

    domReady = true;
    applyCapabilityUI();

    refresh().then(function () {
      el.loading.hidden = true;
      el.app.hidden = false;
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }

  if (window.Office && Office.onReady) {
    Office.onReady(function () {
      officeReady = true;
      try {
        insertSupported = Office.context.requirements.isSetSupported("PowerPointApi", "1.2");
        captureSupported = Office.context.requirements.isSetSupported("PowerPointApi", "1.8");
      } catch (e) {
        insertSupported = false;
        captureSupported = false;
      }
      capabilitiesKnown = true;
      applyCapabilityUI();
    });
  }
})();
