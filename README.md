# My Slide Library — a PowerPoint add-in

Works like think-cell's slide library: save slides you reuse often, then
drop any of them into the current presentation with one click — the new
slide lands right after whichever slide you currently have selected.

## Features

- **Categories**: every slide belongs to a category (type a name in the
  small field under each thumbnail — existing category names autocomplete).
  Categories render as collapsible sections with a chevron and a count, and
  your expand/collapse choices are remembered between sessions.
- **Infinite scroll**: each category only renders 30 slides at a time and
  quietly loads more as you scroll near the bottom — safe for a library of
  hundreds of slides without the pane ever feeling sluggish.
- **"Save current slide"** captures whatever slide you're currently on,
  including a real thumbnail — no manual file prep needed.
- **"Import all slides from this file"** is the answer to "I already have
  200 slides in one file": open that file as your active presentation,
  click this, optionally give the whole batch one category, and it
  captures every slide individually and automatically. You don't need to
  split the file yourself.
- **"Upload slide files (.pptx)"** is a manual fallback: drop in
  ready-made single-slide .pptx files if you'd rather build the library
  that way, or if your Office version is too old for live capture.
- A "Keep original slide design" checkbox controls whether inserted slides
  keep their own look or adopt your current presentation's theme.

## One real requirement to know about

Capturing slides ("Save current slide" / "Import all slides") needs a
fairly recent Microsoft 365 (roughly **April 2025 or newer** — PowerPoint
build 18730.20030+). If your version is older, the add-in shows a banner
and disables those two buttons, but you can still **insert** items and
build the library via "Upload slide files" instead — inserting only needs
a much older Office version, so that part works almost everywhere.

## Setup

1. Push this whole `slide-library-addin` folder into your GitHub repo
   (public), then enable **Pages** (Settings → Pages → Deploy from branch
   → main → root).
2. `manifest.xml` already has your real repo's URLs filled in — no editing
   needed unless your repo/folder name changes later.
3. Put `manifest.xml` in a local folder (e.g.
   `C:\Users\olive\Documents\SlideLibraryAddin`), share that folder
   (right-click → "Grant access to" → "Specific people"), and note the
   `\\...` path Windows gives you.
4. **`TrustSlideLibraryCatalog.reg`** is pre-filled assuming that path is
   `\\DESKTOP-J7BBO6F\Users\olive\Documents\SlideLibraryAddin` — if yours
   differs, edit the `Url` line before running it (or send me the real
   path).
5. Double-click the `.reg` file, confirm the prompt, fully restart
   PowerPoint, and add it from **Einfügen → Add-Ins hexagon → GETEILTER
   ORDNER**, same as the icon add-in.

## Updating later

Same rule as the icon add-in: bump the `?v=` numbers in both
`manifest.xml` and inside `taskpane.html` whenever you change
`taskpane.js`/`.css`, or a cached copy can keep loading indefinitely.
