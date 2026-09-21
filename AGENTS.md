# Notes for agents working in this repo

Persistent working notes. Keep them factual and short; update or delete entries that stop
being true. Things derivable from the code or git history don't belong here.

## Layout

- `Main/Avalonia/` - the live codebase. Solution `LiveGeometry.Avalonia.slnx`:
  - `DynamicGeometry/` - the geometry library (assembly `DynamicGeometry.Avalonia`), a fork of
    the WPF sources rewritten onto Avalonia, plus `WpfCompat/` shims.
  - `LiveGeometry/` - shared UI (`MainView`), used by both heads.
  - `LiveGeometry.Desktop/` - desktop head (net10.0).
  - `LiveGeometry.Browser/` - WASM head (net10.0-browser), deployed to https://livegeometry.com.
- `Main/DynamicGeometryLibrary`, `Main/WPFClient`, `Main/Silverlight*` - the older WPF/Silverlight
  generation. Reference only.
- `Reference/VB6/Source/` - source of the original VB6 app ("DG"). A built copy lives outside the
  repo at `C:\Dropbox\Projects\DG 1\Source\Geometry.exe` (process name `Geometry`).
- `tools/` - file-based C# tools (`dotnet run tools/x.cs -- args`), see below.

## Working rules

- Kirill reviews and commits everything himself: never commit, push, or otherwise change git
  state. Leave changes in the working tree, no need to report them.
- Scripting is C# file-based apps (`dotnet run file.cs`). No Python on this machine.

## Code style

Same conventions as the Helix repo (`C:\Ide\AGENTS.md`), minus what is specific to that codebase.

- **Apply these rules to new code only.** Keep diffs minimal - don't reformat or rename existing
  code as a drive-by. Most of `DynamicGeometry/` is decades-old WPF/Silverlight-era code with
  block namespaces and LF endings; it stays that way unless a cleanup is explicitly asked for.
- **New files: CRLF line endings, file-scoped namespaces.** Existing files keep what they have.
  The Write tool emits LF: after creating a file, fix it with the Helix MCP
  (`get_file_info` / `set_file_format lineEnding=CRLF`; needs `start_ide` first). Never grep
  for `\r`. `stop_mcp` before building.
- **Never edit workspace files from the shell** (sed/perl/heredocs). Read/Edit/Write or Helix tools.
- **Backwards compatibility is not a concern** inside the repo: every caller is in-tree, so
  change all call sites rather than keep a worse shape. The exception here is the `.lgf` file
  format - users have saved drawings, so old files must keep loading.
- **Always brace** single-statement `if`s, even one-liners.
- **No two consecutive blank lines.** One blank between members, none at the top of a block,
  none before `}`.
- **Blank line after a block** - control-flow statements and member/type declarations alike -
  unless it is the last thing in its enclosing block.
- **Before finishing a file**, sort usings and remove unused ones.
- **Avoid `sealed`. Avoid `internal`. Prefer `public`.** Only with a concrete reason.
- **No abbreviations in identifiers** (`operation`, not `op`). Exceptions: `sb`, `ex`, `sw`.
- **No `Async` suffix on async methods.**
- **More than 4 parameters: one per line**, in declarations and at call sites.
- **Name literal arguments at call sites**: `overwrite: true`, `parent: null`, `radius: 0.32` -
  a bare `true`/`null`/number in an argument list is opaque. Self-describing expressions don't
  need it. (Runs of x/y coordinates, as in `IconBuilder` calls, are fine bare.)
- **MSBuild conditions: no quotes** around property names or `true`/`false`:
  `Condition="$(Configuration) == Debug"`. Quote only values that can be empty or contain spaces.
- **No global.json.** The repo must build with whatever new-enough .NET SDK is installed.
- Build with `dotnet build` / `dotnet publish` here (unlike Helix, there is no WPF markup
  compilation in the Avalonia solution). Pass `-bl` and read the binlog with the binlog MCP
  tools when a build misbehaves, instead of parsing console output.

## Build and run

- Desktop: `dotnet build Main/Avalonia/LiveGeometry.Desktop/LiveGeometry.Desktop.csproj`, then
  run `bin/Debug/net10.0/LiveGeometry.Desktop.exe [drawing.lgf|.dgf]`. Fastest loop; use it for
  feature work. A drawing can be hand-written as XML (`docs/*.lgf` are examples: styles, then
  figures by type name with `<Dependency Name=...>` children; a point's name shows through a
  separate `PointLabel` figure that depends on it) and opened from the command line to check it.
- Browser: needs the `wasm-tools` workload (`WasmBuildNative=true` links Skia/HarfBuzz; without
  it Skia throws DllNotFoundException). Do a full bin/obj clean when native assets change.
- Browser publish: `dotnet publish Main/Avalonia/LiveGeometry.Browser/LiveGeometry.Browser.csproj -c Release -o <dir>`.
  Output is `<dir>/web.config` + `<dir>/wwwroot/`. CI (`.github/workflows/main_livegeometry.yml`)
  does exactly this and deploys to Azure App Service (IIS).

## Gotchas

- **Trimming only happens on Release publish**, never in `dotnet run`. The geometry library is
  discovered via reflection (`GetTypes()` for behaviors, serializers, figures), hence
  `TrimmerRootAssembly DynamicGeometry.Avalonia` in the Browser csproj. A trimming break shows
  up as a white screen and a `CRASH: ...` console line (`Program.cs` prints those). Anything new
  that is reached only by reflection outside that assembly needs its own root.
- **StyleKeyOverride**: a subclass of a templated Avalonia control (TabControl, TabItem, ListBox,
  Button, UserControl, ColorPicker...) renders invisible unless it overrides `StyleKeyOverride`
  to return the base type.
- **Input**: all WPF-style mouse events funnel through `DynamicGeometry/Behaviors/Behavior.cs`
  (pointer adapters live there). Avalonia has no static Keyboard; modifiers come from event args.
- **Mutating a collection in place does not redraw in Avalonia** the way a WPF Freezable did.
  `Polygon/Polyline.Points` only rebuild geometry when the property gets a different list.
  Figures keep their allocation-free point cache and call `Shape.PointsChanged()`
  (`PolygonShape`/`PolylineShape` in `WpfCompat.cs`, created by `Factory`). Suspect the same
  thing whenever something "renders once and then never updates".
- **Right button / hover** go through `Behavior.MouseRightClick` and `Behavior.GetCursor`
  (virtual, per tool). Tool letter shortcuts live in `UI/BehaviorShortcuts.cs` (also feeds the
  toolbar tooltips); plain-key handling (letters, arrows, +/-, H) is in `MainView.HandlePlainKey`.
- **The view** (`Figures/Coordinates/CoordinateSystem.cs`) is an origin in pixels plus `UnitLength`
  (pixels per unit). Everything goes through `Zoom(factor, focus)` (the point under `focus` stays
  put: the cursor for the wheel, the canvas middle for +/- and the menu), `Fit`/`SetView`
  (`ZoomExtend` = zoom to fit with a pixel margin, `CenterContent` = Home, `SetViewport` for files)
  and the resize handler, which keeps the middle of the canvas in the middle. "Content" for fit is
  `TryGetContentBounds`: points, whole ellipses, labels - not lines. Labels have a fixed *pixel*
  size, so fit re-measures and refits a few times. `winauto wheel <t> x y <notches>` tests the wheel.
- **No pixel snapping of figure geometry.** The WPF code rounded line endpoints to pixel centers
  (`Round() + 0.5` in `Utilities.Set(Line, PointPair)`); points weren't snapped, so lines missed
  their own points by up to a pixel, always down-right. Lines are exact now and point shapes have
  `UseLayoutRounding = false`. Anything new that must line up with figures: same two rules. To
  check alignment, `winauto shot ... --region x,y,60,60 --zoom 16` around a point.
- **Default point styles are by kind** (`StyleManager.AddDefaultStyles` / `AssignDefaultStyle`):
  `FreePoint` (yellow), `PointOnFigure` (green) - the draggable ones, size 10 - and
  `IntersectionPoint` (blue), `Midpoint` (orange), `DependentPointStyle` (gray, every other
  constructed point) at size 8. Looked up by *name*; a drawing from a file brings its own styles,
  and when the named one is missing (older files) the first point style is used. The click
  preview ghost uses the same lookup. The Point tool's panel style overrides the kind only once
  the user picks something other than the first style.
- **Cursor philosophy** (`Behavior.GetCursor`): cross = a new *free* point appears here; hand =
  the click picks something already there - a figure the tool needs, an existing point, or a
  place defined by figures (intersection, midpoint); arrow = everything else, including a new
  point sliding along a figure and clicks that do nothing. For points it is derived from the
  `PointPlacement` (`Behavior.GetCursor(PointPlacement)`), so cursor, preview and click agree.
  `winauto cursor` prints the cursor showing now (screenshots don't include it).
- **What a click makes of a point** is decided in one place, `Behaviors/PointPlacement.Find`
  (existing point / free / on a figure / intersection of the nearest crossing pair / midpoint
  when the cursor is within `MidpointReach` of a segment's middle, or anywhere on it with "Snap
  to center" on; a midpoint that already exists for the two points is reused, never duplicated -
  `FindExistingMidpoint`, which the Midpoint tool checks too). The Point tool and every `FigureCreator`
  (`FindPointPlacement`, `CreatePointForClick`) use it for the click, and `Behavior.GetClickPreview`
  feeds the same answer to `Behaviors/ClickPreview` on hover: a 0.4-opacity ghost point, a halo
  on the source figures, equal-halves ticks for a midpoint. A tool that needs a figure rather than a point gets a halo
  on the figure a click would pick (`Behavior.GetFigureToPick`, `FigureCreator.FindFigureToPick`;
  an existing point the click would take counts too and gets a disc behind it;
  halos exist for points, lines, circles/ellipses, arcs and polygons - add a case to
  `ClickPreview.CreateHalo` for anything else). The preview is plain canvas visuals,
  never figures. Hit testing uses the *snapped* coordinates, so with snap to grid on the grid
  wins. Typed coordinates always give a free point. A point must never be placed on the figure
  being constructed (`FigureCreator.CanPlacePointOn`), that would be a dependency cycle.
- **Side panel (property grid) look**: the surface is a Border in `DrawingHost.CreatePropertyGrid`
  using `RibbonTheme` colors; the editors are restyled by scoped Avalonia styles in
  `PropertyGrid/PropertyGridTheme.cs` (no per-editor styling code). The Fluent theme paints
  hover/selected states on the template's ContentPresenter, so overrides must target that part.
  Rows share the label column width via `SharedSizeGroup`. Layout is declarative:
  `[PropertyGridGroup("Name")]` on properties/methods boxes them together (editors, then their
  buttons in a row); `[PropertyGridDestructive]` on a method puts its button last, under a
  divider, with a trash can (`PropertyGrid.Arrange`, `MethodCallerButton`).
- **Color/brush picking** lives in `DynamicGeometry/Controls/ColorPicker/` and is layered so the
  parts can be swapped: `ColorPalette` (which colors, in what order and how many columns -
  `WebColors` is the hand-arranged 14x10 map from the Helix picker, `ArrangeByHue` computes one) ->
  `ColorPage` (one way to pick: `SwatchPage`, `SpectrumPage`; `SwatchPage` has
  `SwatchSize`/`Spacing`/`SwatchCornerRadius` and virtual `CreateSwatch`/`ShowSelected`/`CreateLayout`)
  -> `ColorPickerView` (page switcher + sample + name/hex box) -> `BrushPickerView` (Solid |
  Gradient; a `GradientStopBar` with any number of draggable stops, click to add, drag away to
  remove; the one color picker edits the selected stop). In the property grid they appear through
  `ExpandingPickerEditor` (a chip that unfolds the picker under its row, one at a time, inside a
  frame that groups row + picker). The Solid|Gradient and Swatches|Spectrum strips are
  `SegmentSwitcher`s, drawn with the ribbon's `TabOutline`; whoever hosts a picker sets its
  `Surface` to the background it sits on so the selected tab blends into it.
- **The Write tool turns a CRLF file into LF.** After rewriting an existing CRLF file with Write
  (rather than Edit), set CRLF again with Helix `set_file_format`.
- **Numbers with a range** (`[Domain(min, max)]` doubles: stroke width, point size, font size) are
  edited by `SliderEditor`: text box + `Controls/UpDownControl` (repeat buttons, also Up/Down keys
  and the wheel over the box; `UpDownControl.Step` lands on whole steps and clamps) + slider.
- **Colors in files are always `#AARRGGBB`** (`ColorText.ToArgbHex`). Never `Color.ToString()`:
  Avalonia writes the *name* of a known color, and builds before 2026-09 did exactly that, so
  `ToColor()` accepts names too - 6-9 letter names used to be parsed as hex and threw.
  A gradient fill is a child element of the style (`<Fill><LinearGradientBrush>`), not an attribute.
- **The library's own types shadow framework ones**: `Math`, `Ellipse`, `Polygon`, `Path`...
  In a file-scoped-namespace file a `using X = ...;` alias does NOT win over a type of the
  enclosing namespace - write `System.Math.Max`, `Avalonia.Controls.Shapes.Ellipse` in full.
- **`Style` and `Setter` are ambiguous in the library**: `DynamicGeometry.Style`/`Setter` are the
  WPF shims and shadow Avalonia's. For real Avalonia styles alias them
  (`using AvaloniaStyle = Avalonia.Styling.Style;`).
- **Keyboard focus drifts into tool panels.** A tool's PropertyBag panel (e.g. "Point by
  coordinates") takes focus into its TextBox after every construction step, so neither the canvas
  KeyDown nor `MainView_KeyUp` (which skips TextBox focus) sees keys then. Anything that must
  always work (Escape) belongs in the `MainView_KeyDown` tunnel handler.
- **Toolbar look** is centralized in `UI/Ribbon/RibbonTheme.cs`; `ButtonGrid` draws the
  hover/pressed/checked plate. `Ribbon` and `TabPanel` replace the Fluent TabControl/TabItem
  templates with their own (in code): the header row has a bottom line *behind* the headers and
  the selected group header paints a tab shape over it (`TabOutline`, laid out wider than the
  header by its flare via negative margin). Group headers show the icon of the group's active tool. An on/off `Command` exposes `IsChecked` (a `Func<bool>`), which
  its button re-reads after any toggle is clicked - don't go back to `CheckBox` icons.
- **Browser has no system fonts.** Text renders only because `Avalonia.Fonts.Inter` is embedded
  (`.WithInterFont()`); font names stored in drawings (Arial etc.) fall back to it.
- **web.config**: `LiveGeometry.Browser/web.config` is hand-written (serves the precompressed
  `.br` files, sets immutable caching on fingerprinted assets, `no-cache` on entry files). The
  wasm SDK drops a project web.config from publish, so the csproj copies it with an explicit
  `AfterTargets="Publish"` target. It sits at the publish root and rewrites into `wwwroot\`.
  There is no IIS locally: verify changes after deploy with
  `curl -s -o /dev/null -D - -H "Accept-Encoding: gzip, br" https://livegeometry.com/_framework/<file>`
  and expect `Content-Encoding: br` and `Cache-Control: public, max-age=31536000, immutable`.
  A malformed web.config takes the whole site down (HTTP 500).

## Deployment and caching

- A push to main takes about 8 minutes to go live (GitHub Actions: build ~5, deploy ~3); the
  files flip at the very end. Until then every reload, cached or not, gets the previous version.
  The commit shown at the right end of the menu bar (`BuildVersion`, also logged to the console at
  startup) tells which build is on screen.
- Caching was verified end to end (2026-09-20): entry files are `no-cache` with an ETag that
  changes on every deploy, a request with the old ETag gets 200, and a browser holding the
  previous version picks up the new one on a plain reload. If something looks stale, check the
  commit stamp and the Actions run before suspecting the cache.

## UI automation (tools/)

Screenshots are PNGs; image pixels are the click coordinates in both tools.

- `tools/winauto.cs` - any desktop window (the VB6 app, the Avalonia desktop app).
  `list`, `tree <t>`, `menu <t>`, `invoke <t> <menuId>`, `shot <t> out.png`, `click <t> x y [right|double]`,
  `drag <t> x1 y1 x2 y2`, `keys <t> "^s"`, `text`, `focus`, `place <t> x y w h`.
  Target = process name | `pid:N` | `hwnd:0x..` | `title:substr`.
  - Both apps start maximized on a 4K/200% monitor: `place <t> 100 100 1500 1000` first.
  - VB6 has a native menu: `menu Geometry` lists every command with its id and `invoke` runs one
    without touching the mouse (49 = Segment, 56 = Circle, 46 = Point...). Its status bar shows
    the current tool's prompt. VB6 is DPI-unaware; `shot` compensates.
  - Avalonia has no native menus or child windows: screenshot, then click by coordinates.
- `tools/webauto.cs` - the browser build in headless Edge over CDP (port 9333, Edge stays alive
  between calls). `start <url> [w h]`, `stop`, `nav`, `wait <console text>`, `console [--errors]`,
  `shot`, `click`, `drag`, `move`, `key <Key> [ctrl] [shift] [alt]`, `text`, `eval <js>`.
  - The app takes several seconds to boot after `start`; the splash `div` stays in the DOM, so
    don't test for its absence - take a screenshot.
  - Edge runs with `--guest`; without it Edge signs the throwaway profile into the Windows
    account and opens a sync dialog as an extra page target.
- `tools/serve.cs <publish>/wwwroot [port]` - static server for a publish output (blocks; run in
  the background). Separate file because a running server locks its own exe.

Web smoke test (run before deploying changes that touch file I/O, clipboard, fonts, or anything
reflection-based): publish Release, `serve` its wwwroot, `webauto start http://localhost:5005/`,
`console --errors` must show no `CRASH:`, draw a segment via the Lines tab, `shot`, `stop`.

  - `winauto keys` sends real virtual keys for lowercase ASCII letters/digits (needed for the
    single-letter shortcuts); other characters go as Unicode packets, which KeyDown-based
    shortcuts never see. A context menu is its own top-level window: while one is open,
    `shot LiveGeometry.Desktop` captures the menu, not the main window.

## VB6 parity backlog

Done (2026-09): point labels draggable in an orbit around the point, right-click (cancel /
back to Drag / close polygon / context menu), hover cursors, Shift = snap to grid, tool letters,
arrows/+/-/H, opening `.dgf` files (cp1251 fallback). Deliberately out of scope for now:
Calculator, step-by-step construction playback.

Still missing compared to VB6, roughly by value: symmetric point (about a point) and inverted
point (in a circle) tools; tracing locus of a point ("Create locus" on a point); "snap free point
to figure" / "release point" from the context menu; "Choose point/figure" disambiguation for
overlapping figures; double-click opens properties (here: double-click = zoom to fit); measurement
label dragging constraints; point shape/size per point and name color; line dash styles per
figure; Show/Hide, message, sound and launch buttons; live cursor coordinates in the status bar;
rulers; undo/redo captions naming the action; unsaved-changes prompt; recent files; export
(BMP/WMF/EMF -> here PNG would do) and print; "tool select once" option; settings persistence;
languages (en/ru/uk/de).

## Not yet verified in the browser

.lgf open/save through the browser storage provider, PNG export, printing, demo download, and
the `Hyperlink` figure's use of `WebClient`.
