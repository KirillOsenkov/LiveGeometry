# Testing and automation

Checking drawings in batch, and driving the desktop and browser builds from tools/.

## Batch-checking drawings

`LiveGeometry.Desktop.exe --check <folder> <out>` (`MainView.BatchCheck.cs`) opens every
`.lgf`/`.dgf` under the folder (or the one drawing, given a file) in the real editor, zooms to fit (a captioned drawing is laid
out by `GalleryDrawing.Fit` instead, with the ribbon folded when the window is small - so the
sheet shows the gallery layout at whatever size the window was last closed at, or at the
size `--size 390x645` gives the window's client area in layout pixels (an iPhone's
viewport; the window's saved place is then neither read nor written; `--dark` and
`--size` come in any order after `<out>`), saves `<out>/<relative path>.png`
and `.lgf` (the conversion) and appends to `<out>/report.txt`: figure counts, `NOT EXISTING`
(only the *root* failures - figures whose dependencies all exist - plus a dump of every point),
load errors. It exits when done. `dotnet tools/contactsheet.cs -- <png folder> <out.png>
[columns] [tile width]` tiles the PNGs into one image: the fastest way to eyeball a whole
folder (the whole gallery fits on one 4-column sheet). The VB6 CD library (221 files, kept
outside the repo, read-only) all loads; what the report still calls "missing" there is second intersections that fall outside
a segment or ray, and sides of a polygon that don't cross - legitimately absent.

## UI automation (tools/)

`dotnet run tools\regression.cs` runs focused construction, property/editor, dragging,
undo/redo, expression-cycle, and LGF/GGB regression cases against the desktop libraries,
including loading and reloading every gallery drawing. Pointer cases briefly open test
windows. This is not an exhaustive UI or importer compatibility suite.

Screenshots are PNGs; image pixels are the click coordinates in both tools.

A tool that starts something that outlives it (Edge, `serve.cs`) starts it through the shell
(`UseShellExecute = true`, hidden): with `UseShellExecute = false` the child inherits every
inheritable handle, the caller's output pipe included, and a run piped through `tail` waited
until Edge was closed. A tool that waits on another process or on a page gives it a timeout
and fails with words (`playerparity.cs`'s `Webauto`).

- `tools/winauto.cs` - any desktop window (the VB6 app, the Avalonia desktop app).
  `list`, `tree <t>`, `menu <t>`, `invoke <t> <menuId>`, `shot <t> out.png [--screen]` (the
  window's own rendering, or what is on the screen there - the only way to see the title bar), `click <t> x y [right|double]`,
  `move <t> x y` (hover, for click previews and `cursor`),
  `drag <t> x1 y1 x2 y2 [steps] [--shift] [--alt]`, `path <t> x1 y1 x2 y2 ... [--right]
  [--shift] [--alt] [--hold]` (a drag through the points, a lasso; `--hold` keeps the button
  and keys down for a `shot` in the middle, `release [--right] [--shift] [--alt]` lets go),
  `wheel <t> x y <notches>`, `keys <t> "^s"`, `text`, `focus`,
  `cursor`, `place <t> x y w h`, `maximize <t>`. Target = process name | `pid:N` | `hwnd:0x..` | `title:substr`.
  - The VB6 app starts maximized: `place <t> 100 100 1500 1000` first. The Avalonia desktop app
    remembers its window (`LiveGeometry.Desktop/WindowPlacementPersistence.cs`, the
    `WindowPlacement` line of `Settings.txt` in the user's local app data, next to the theme
    choice). It is left maximized for testing (`winauto maximize <t>` once, then close it),
    since a full screen and an iPhone in portrait (below) are the two layouts that matter: check
    a change at both, not in an in-between window. The screen is 3840x2160 at 200% scaling, so
    a full-screen `shot` (about 3866x2064) comes back scaled down: multiply what is read off it
    by the factor the image reports before clicking, or `shot --region` the part to look at.
    Leave it maximized after a phone check. Close test instances with
    `(Get-Process -Id N).CloseMainWindow()`: that goes through the normal close path, which is
    what saves the placement (and it is more reliable than Alt+F4).
    `LiveGeometry.Desktop.exe --arrange` opens the gallery with its tiles draggable; every drop
    rewrites the `Items` block of `GalleryCatalog.cs` in the new order (comments between the
    items go). `LiveGeometry.Desktop.exe --gallery <slug>` opens a gallery drawing the way the
    browser's `/gallery/<slug>` does - a file on the command line opens as the user's own
    drawing instead. An iPhone 13 Pro in portrait gives Safari a 390x645 viewport, of which the
    two toolbar rows take 82 and the canvas gets 390x563; landscape is about 844x310 with a
    844x250 canvas (estimated). For the browser build that is `webauto start <url> 390 645`;
    for the desktop app at 200% scaling `place <t> 100 100 780 1352` (the chrome and toolbar
    take 110 DIPs) and `1688 800` for landscape.
  - When testing pans, start the drag on an empty spot (a drag that starts on a figure moves
    the figure), use many steps (`drag <t> x1 y1 x2 y2 300`), and compare against the axis
    numbers.
  - VB6 has a native menu: `menu Geometry` lists every command with its id and `invoke` runs one
    without touching the mouse (49 = Segment, 56 = Circle, 46 = Point...). Its status bar shows
    the current tool's prompt. VB6 is DPI-unaware; `shot` compensates.
  - Avalonia has no native menus or child windows: screenshot, then click by coordinates. A
    native file dialog shows in `list` as a `#32770` window of the app: `text hwnd:0x.. <path>`
    then `keys hwnd:0x.. "{ENTER}"`. A context menu is its own top-level window: while one is
    open, `shot LiveGeometry.Desktop` captures the menu, not the main window (and `list`
    doesn't show it). It opens at the cursor, so click its items by main-window coordinates:
    right-click point + the item's place in the menu shot, with `hwnd:` of the main window
    (by `pid:` a click can land on the main window's system menu).
  - After a figure with a length is made (segment, square, circle...) the length panel opens
    at the right of the canvas and swallows clicks there; the next click on the canvas closes
    it. Take a shot after each construction, or keep test clicks away from the right edge.
  - `winauto keys` sends real virtual keys for lowercase ASCII letters/digits (needed for the
    single-letter shortcuts); other characters go as Unicode packets, which KeyDown-based
    shortcuts never see.
- `tools/macauto.cs` - winauto's counterpart on a Mac (CoreGraphics events, `screencapture`,
  System Events to place a window; the terminal that runs it needs the Accessibility and
  Screen Recording permissions). `list`, `shot <t> out.png [--screen]`, `click <t> x y
  [right|double] [--shift] [--alt] [--cmd] [--ctrl]`, `move`, `drag <t> x1 y1 x2 y2 [steps]
  [modifiers]`, `wheel <t> x y <notches>`, `key <t> <key> [cmd] [ctrl] [shift] [alt]` (`s cmd`,
  `Escape`, `Return`), `keys <t> <letters>` (real key codes: tool letters), `text <t> <text>`
  (Unicode, for text boxes), `focus`, `place <t> x y w h` (in points). Target = process name |
  `pid:N` | `window:N` | `title:substr`. Pixels are those of a Retina shot (twice the points),
  relative to the window's top left, title bar included; a y above 0 reaches the menu bar.
  Run it with `dotnet run --file tools/macauto.cs -- ...`: from inside a project folder a bare
  `dotnet run` takes the project instead. A native save panel and a context menu are windows
  of their own in `list` (a menu on layer 101), and `key`/`text` with `window:N` go to them;
  in the save panel Cmd+Shift+G types a folder.
- `tools/loadperf.cs` and `tools/scrollperf.cs` - how the browser build loads and scrolls, in
  headless Edge of their own (ports 9334 and 9335): see "Measuring speed".
- `tools/webauto.cs` - the browser build in headless Edge over CDP (port 9333, Edge stays alive
  between calls). `start <url> [w h] [--lang xx-XX] [--dark]`, `stop`, `nav`, `wait <console text>`,
  `console [--errors]`, `shot`, `click`, `drag`, `move`, `key <Key> [ctrl] [shift] [alt]`,
  `text`, `eval <js>`, and fingers: `tap x y [double]`, `touch x1 y1 x2 y2` (one-finger
  drag), `pinch x1 y1 x2 y2 x3 y3 x4 y4` (one finger from 1 to 2, the other from 3 to 4).
  - The app takes several seconds to boot after `start`; the splash `div` stays in the DOM, so
    don't test for its absence - take a screenshot. A tap or pinch before the app has
    loaded reaches the page, which zooms itself (`visualViewport.scale` 4); `stop` and
    `start` again for a clean page.
  - Edge won't make its window narrower than about 490 px: phone widths are tested on the
    desktop app (`winauto place` at 736 wide is 368 at 200% scaling).
  - `text` (CDP's `Input.insertText`) does not reach a text box of the app: type with `key`,
    one character a call (`key -`, `key x`, `key ^`, `key 2`, `key Enter`).
  - Edge runs with `--guest`; without it Edge signs the throwaway profile into the Windows
    account and opens a sync dialog as an extra page target.
  - `--dark` (on `start`, a fresh one: `stop` first) runs Edge with `--force-dark-mode`, so
    `prefers-color-scheme` answers dark and the app's `System` theme comes up dark. A CDP
    media emulation would not do: it dies with the session that set it, before the app reads
    the query. The stored choice is tested with
    `eval "localStorage.setItem('LiveGeometry.Theme','Dark')"` and a `nav`; `--guest` starts
    with an empty storage every time.
- `tools/serve.cs <publish>/wwwroot [port]` - static server for a publish output (blocks; run in
  the background). Separate file because a running server locks its own exe.

Web smoke test (run before deploying changes that touch file I/O, clipboard, fonts, or anything
reflection-based): publish Release, `serve` its wwwroot, `webauto start http://localhost:5005/`,
`console --errors` must show no `CRASH:`, draw a segment via the Lines tab, `shot`, `stop`.
A change to the theme wants the same once more with `start ... --dark`.
