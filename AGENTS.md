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
  `TryGetContentBounds`: points, whole ellipses, labels - not lines. Panning (`MoveTo`) must not
  round the origin: a drag is many sub-pixel steps (0.5 px at 200% scaling) and rounding each one
  made the plane run up to twice as fast as the cursor. When testing pans with winauto, start the
  drag on an empty spot (a drag that starts on a figure moves the figure, and the grid "stays
  behind"), use many steps (`drag <t> x1 y1 x2 y2 300`), and compare against the axis numbers. Labels have a fixed *pixel*
  size, so fit re-measures and refits a few times. `winauto wheel <t> x y <notches>` tests the wheel.
- **Vectors**: `Vector` = an invisible `Segment` + an `Arrow` (a 7-point polygon: shaft + head).
  The arrow is computed in pixels (`Arrow.HeadLength/HeadHalfWidth/HeadGrowth`, shaft = the
  style's stroke width), is filled with the *line* color, has no outline, and stops at the rim
  of its end point. `Vector.OnAddingToCanvas` gives it the default `LineStyle` before the base
  call, otherwise the polygon default (pale translucent fill) wins; vectors in older files keep
  their polygon style and get its fill.
- **`PathFigure.IsClosed` defaults to true in Avalonia** (false in WPF), and `IsFilled` to true.
  Every hand-built `PathFigure` must set both; a forgotten one draws a line from the end of the
  path back to its start (it did, for function graphs and loci, in `Curve`).
- **No pixel snapping of figure geometry.** The WPF code rounded line endpoints to pixel centers
  (`Round() + 0.5` in `Utilities.Set(Line, PointPair)`); points weren't snapped, so lines missed
  their own points by up to a pixel, always down-right. Lines are exact now and point shapes have
  `UseLayoutRounding = false`. Anything new that must line up with figures: same two rules. To
  check alignment, `winauto shot ... --region x,y,60,60 --zoom 16` around a point.
- **Right angles**: `Figures/Lines/RightAngleMark.cs` draws the two-sides-of-a-square sign. It is
  a passive visual owned by figures that are perpendicular by construction (`PerpendicularLine`
  at the foot, when the foot is on the base figure; `SegmentBisector` at the midpoint), gray and
  translucent, switchable per figure ("Right angle mark", saved as `RightAngleMark="false"`).
  Both figures derive from `PerpendicularLineBase`, which owns the mark. Which of the four
  corners it sits in is *stored* (`RightAngleMark.Corner`, 0-3 counterclockwise from the base
  line's P1->P2 direction, saved as `RightAngleCorner`), chosen once when the line first shows up
  and never derived from the geometry again - deriving it made the mark flip-flop on rounding.
  A click on the mark with the Drag tool moves it to the next corner (undoable; the mark has a
  transparent square as its click area and marks the press handled so the Dragger never sees it). It
  hides when an `AngleArc` sits at the same vertex, because a measured angle of 90° (within
  0.005°, i.e. exactly when the label reads "90°") draws the same sign itself in its own style
  instead of an arc. Deliberately not an auto-created angle figure.
- **Angle marks**: an angle is two figures - `AngleMeasurement` (the number) and `AngleArc` (the
  mark), same three dependencies. `AngleArc.ArcCount` 0-3 draws nothing, ), )) or ))) (saved as
  `Arcs`, default 1; spacing = stroke width + 2 px like VB6) and `Size` is the radius (`Radius`,
  default 16). The label's panel has the same "Arcs" box and forwards to its arc
  (`AngleMeasurement.FindArc`), because with 0 arcs there is little left to click. The pair is
  found by `AngleArc.FindCompanion`; "Convert to opposite angle" on either goes through
  `AngleArc.ConvertToOpposite(figure)`, which swaps the sides of both. `AngleArc.HitTest` covers
  the whole band from the first arc to the last (the base class only knows the first radius). The path keeps
  its first `PathFigure` even for "nothing" (no segments): Avalonia doesn't repaint a path whose
  figure list became empty. `DGFReader.ReadMeasureAngle` creates the arc from VB6's DrawStyle /
  AuxInfo(2) - not tested, there is no sample .dgf with an angle in the repo.
- **Measurement labels are draggable** (`Measurement.AllowMove`): a drag only changes the label's
  `Offset` from its anchor, although the figure has dependencies.
- **Dashes**: `LineStyle.Dash` is a `LineDash` enum (`Styles/LineDash.cs`: Solid, Dash, Dot,
  DashDot, DashDotDot - the VB6 DrawStyle 0-4, which `DGFReader` maps onto it). Inherited by shape
  and point styles, so circles, arcs and polygons dash too. Saved as `Dash="DashDot"` by the
  generic `EnumSerializer` (any enum property of a style or figure now round-trips by name; an
  unknown name keeps the default). The pattern is put on in `LineStyle.OnApplied`, not through a
  setter, because `StrokeDashArray` counts in stroke widths and a selected figure is thicker.
  Anything that applies a style to a shape by hand (sample glyphs) must call `OnApplied` too.
  To drive a native file dialog from winauto: `list` shows it as a `#32770` window of the app;
  `text hwnd:0x.. <path>` then `keys hwnd:0x.. "{ENTER}"`.
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
- **Gallery and routes** (`LiveGeometry/Gallery/`, `MainView` "Pages" region). `MainView` shows
  either `GalleryView` (the start page: "New Drawing" + a tile per drawing) or the editor. Three
  states: gallery, a gallery drawing (`CurrentSample`; the toolbar shows ◀ n/33 ▶ + title, Page
  Up/Down, edits are dropped silently, Save = save as and the drawing becomes the user's own), or
  the user's own drawing (only the Gallery button). The user's drawing object is parked in
  `OwnDrawing` (with its undo history) while they look around, and re-attached by "Back to My
  Drawing". Paths: `/` gallery, `/gallery/<slug>`, `/drawing`; `AddressBar` is the abstraction,
  `BrowserAddressBar` + `main.js` do pushState/popstate, so Back works and links can be shared.
  `index.html` needs `<base href="/">` for that (all routes serve it; `web.config` and
  `tools/serve.cs` fall back to it). All page changes go through `MainView.Show*`.
- **Gallery drawings** are embedded `.lgf` (`Gallery/Drawings`, order and titles in
  `GalleryCatalog`), generated by `tools/gallerize.cs` from two sources: the Windows Phone
  originals (outside the repo; never edit those) and `Gallery/Sources/*.lgf`, which are DG 1.0
  CD library `.dgf` files converted by `--check` (below) - the CD is outside the repo too.
  Change texts in the tool and rerun it rather than editing the output; per-drawing XML
  surgery (a point's `Parameter`, moving a hidden slider, adding a `Locus`, the plane of a grid
  drawing) is a `tweak:` lambda on the sample. Phone samples get `IntersectionOrder="Legacy"`,
  CD samples must not (their intersections were picked from saved coordinates and are right). Each has two unclickable labels, `Title` and `Description` (may contain live `[AB^2]`
  expressions - then list the points as dependencies). Their *position is computed at open time*
  by `GalleryDrawing.Fit`: right of the figure in a wide canvas, below it in a tall one, iterated
  with zoom-to-fit because labels are sized in pixels; refitted on resize until the first edit.
  Drawings with `Grid="true"` (graphs) keep their file viewport in view (`GalleryItem.Plane`),
  because graphs and lines have no bounds.
- **Tiles are live drawings**, not bitmaps (`DrawingThumbnail`): a `Drawing` on its own 560x380
  canvas inside a Viewbox, no Behavior attached, text hidden, loaded one per idle tick. Hovering
  makes the draggable points drift (`IsAnimated`). A load error of a tile goes to the console as
  `Gallery: <file>: ...` - that is the quickest way to check all drawings at once.
- **Old files and circle-line intersections**: `Math.GetIntersectionOfCircleAndLine` at some point
  swapped P1/P2 for a line through the center ("New code - preserves order"). Drawings from before
  (all the phone ones) pick the other intersection, so squares built with perpendicular + circle
  flip inward. Files carry no usable version, so the fix is opt-in: `<Drawing
  IntersectionOrder="Legacy">` makes `DrawingDeserializer` swap the algorithm of every
  intersection whose line passes through the center (`IntersectionPoint.
  UpgradeLegacyCircleAndLineOrder`, numeric test, in construction order). Also: a label without
  `DecimalsToShow` now gets the default 2, not 0.
- **Saved `.lgf` declare `encoding="utf-8"`** now (`DrawingSerializer.Utf8StringWriter`);
  builds before 2026-09-21 wrote `utf-16` into a UTF-8 file, which our own loader tolerates but
  `XDocument.Load` does not. `LineByEquation` in general form gets its two points from the
  foot of the perpendicular from the origin plus the direction, not from the intercepts (a
  line through the origin had both intercepts at the origin and was degenerate).
- **Zoom to fit** counts the vertices of visible segments/polygons/Béziers even when the points
  are hidden, measures labels itself (`Measure`, their Bounds are stale right after a load), and
  takes an optional rect to keep in view.
- **No menu.** New, Open, Save | Undo, Redo are one row (`LiveGeometry/MainToolbar.cs`) above
  the ribbon, in the *same gray as the ribbon's header row* and with no line between the two, so
  they read as one band (two grays in the window, not three); the build stamp is at its right.
  Tried and rejected: its own lighter strip (three grays), buttons right-aligned inside the
  header row (collide with the last tabs at ~850 px), a two-row block at the left of the header
  row (small targets). `Ribbon.HeaderStart`/`HeaderEnd` slots exist from those experiments and
  are unused. The buttons:
  (icons are drawn in code, `MainToolbarIcons`, 20x20 grid; Undo/Redo follow `DrawingControl.
  CommandUndo/CommandRedo` as command observers) plus the build stamp. Everything else is keys:
  Ctrl+N/O/S/Z/Y/A/C/V are handled on key *down* (`MainView.HandleControlShortcut`; on key up
  Ctrl may already be released and a bare S is the Segment tool), plain keys in `HandlePlainKey`.
  Lost their menu entry and are unreachable for now: Lock, Figure List, the settings page.
- **Keyboard focus drifts into tool panels.** A tool's PropertyBag panel (e.g. "Point by
  coordinates") takes focus into its TextBox after every construction step, so neither the canvas
  KeyDown nor `MainView_KeyUp` (which skips TextBox focus) sees keys then. Anything that must
  always work (Escape) belongs in the `MainView_KeyDown` tunnel handler.
- **Toolbar look** is centralized in `UI/Ribbon/RibbonTheme.cs`; `ButtonGrid` draws the
  hover/pressed/checked plate. `Ribbon` and `TabPanel` replace the Fluent TabControl/TabItem
  templates with their own (in code): the header row has a bottom line *behind* the headers and
  the selected group header paints a tab shape over it (`TabOutline`, laid out wider than the
  header by its flare via negative margin). To bring the group headers closer together, change
  `ButtonGrid.HeaderOverlap` (neighboring headers share their flare zones - only the selected
  one draws feet), not the padding inside the tab. Group headers show the icon of the group's active tool. An on/off `Command` exposes `IsChecked` (a `Func<bool>`), which
  its button re-reads after any toggle is clicked - don't go back to `CheckBox` icons.
- **Browser has no system fonts.** Text renders only because `Avalonia.Fonts.Inter` is embedded
  (`.WithInterFont()`). Inter reaches text by *inheritance* from the window (the theme sets it
  there); an explicit family that doesn't exist - and `FontFamily.Default` and
  `FontManager.DefaultFontFamily` too - renders as Noto Mono in the browser. So `TextStyle` sets
  the font a drawing names (Arial, Segoe UI) only when it really resolves, and naming "Inter"
  in a drawing doesn't work (embedded, not installed).
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

## Running instances

Always build into the normal `bin/` - no scratch output folders. Kirill rarely has the app running
and doesn't keep state in it that matters, so a `LiveGeometry.Desktop` process that locks `bin/` is
almost certainly a leftover test instance, and killing it is fine (his words: in 90% of cases).
Only if a build actually fails on a lock, enumerate and clean up:

`Get-Process LiveGeometry.Desktop | Select Id, StartTime, MainWindowTitle, Path` (pwsh), then
`(Get-Process -Id N).CloseMainWindow()` and, if it is still there a few seconds later,
`Stop-Process -Id N`. Alt+F4 through winauto can miss (it goes to whatever window of the app has
focus, e.g. a popup), so check that the process is really gone. Target test windows by `pid:`/`hwnd:`
rather than by process name when more than one could exist.

## Batch-checking drawings

`LiveGeometry.Desktop.exe --check <folder> <out>` (`MainView.BatchCheck.cs`) opens every
`.lgf`/`.dgf` under the folder in the real editor, zooms to fit, saves `<out>/<relative path>.png`
and `.lgf` (the conversion) and appends to `<out>/report.txt`: figure counts, `NOT EXISTING`
(only the *root* failures - figures whose dependencies all exist - plus a dump of every point),
load errors. It exits when done. Contact sheets of the PNGs (a System.Drawing script) are the
fastest way to eyeball hundreds of files. The VB6 CD library
(`C:\Dropbox\Projects\DG 1\DG CD Version 1.0\Library\English`, 221 files, read-only) all loads
as of 2026-09-21; what is still "missing" there is second intersections that fall outside a
segment or ray, and sides of a polygon that don't cross - legitimately absent.

## .dgf (DG 1.0) reader facts

Learned from `Reference/VB6/Source` while making the CD library load (`DGFReader.cs`):
- `modFileIO.bas` writes an `AuxInfo(n)` / `AuxPoints(n).X/Y` only when it is not 0: absent
  means 0 (`GetAuxInfo`/`GetAuxPoint`). An analytic line `x - y = 0` has no `AuxInfo(3)`.
- A point's `Type` is the DrawState that made it; 0 = free point, and old files write
  `ParentFigure=0` for those (which is *not* Figure0). Every dependent point has `Type != 0`.
- `ParentFigure` of a point and figure indices are 0-based; Points, Labels, Buttons 1-based.
- Buttons: type 0 show/hide, 1 message box, 2 sound, 3 open another drawing. Only 0 becomes a
  figure; a show/hide button's list may reference the others.
- DG's angle bisector (`Math.bas GetBisector`) is the *interior* bisector and a whole line;
  ours is oriented (counterclockwise from side 1 to side 2) and a ray. The reader orders the
  sides for the interior one and sets `AngleBisector.IsLine` ("Whole line" in the property
  grid, saved as `Line="true"`).
- A point on a figure is placed by moving it to its saved X,Y in one `MoveTo` (setting X then Y
  projects twice from off the figure and lands elsewhere); `AuxInfo(1)` (t on a line, clockwise
  angle on a circle) is ignored.
- `AuxPoints(6)` of a measurement is the label's shift from its default place, in pixels.
- Intersection points are picked by the saved coordinates of both solution points, so the
  Legacy circle/line order problem does not apply to `.dgf`.
- The VB6 expression language has more than ours: `[A,B]` distance with a comma, comparison and
  logical operators, `IF`/`MAX`/`MIN`, `°` inside brackets. Those labels show the error text;
  the compiler never throws out of a label any more (`Compiler.CompileExpression` catches).
- `IniFile` skips blank lines (every CD file has them between sections).

## UI automation (tools/)

Screenshots are PNGs; image pixels are the click coordinates in both tools.

- `tools/winauto.cs` - any desktop window (the VB6 app, the Avalonia desktop app).
  `list`, `tree <t>`, `menu <t>`, `invoke <t> <menuId>`, `shot <t> out.png`, `click <t> x y [right|double]`,
  `drag <t> x1 y1 x2 y2`, `keys <t> "^s"`, `text`, `focus`, `place <t> x y w h`.
  Target = process name | `pid:N` | `hwnd:0x..` | `title:substr`.
  - The VB6 app starts maximized on a 4K/200% monitor: `place <t> 100 100 1500 1000` first. The
    Avalonia desktop app remembers its window (`LiveGeometry.Desktop/WindowPlacementPersistence.cs`,
    Get/SetWindowPlacement as in Helix; `%LOCALAPPDATA%\LiveGeometry\MainWindowPosition.txt` =
    flags,showCmd,min x,y,max x,y,left,top,width,height). It is left at 100,100 1700x1100 for
    testing - `place` is only needed again if someone resized it. Close test instances with
    `(Get-Process -Id N).CloseMainWindow()`: that goes through the normal close path, which is
    what saves the placement (and it is more reliable than Alt+F4).
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
