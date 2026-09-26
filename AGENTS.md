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
  The Write tool emits LF: after creating a file, or rewriting an existing CRLF file with Write
  rather than Edit, fix it with the Helix MCP (`get_file_info` / `set_file_format
  lineEnding=CRLF`; needs `start_ide` first). Never grep for `\r`. `stop_mcp` before building.
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

## Avalonia and framework traps

- **Trimming only happens on Release publish**, never in `dotnet run`. The geometry library is
  discovered via reflection (`GetTypes()` for behaviors, serializers, figures), hence
  `TrimmerRootAssembly DynamicGeometry.Avalonia` in the Browser csproj; `LiveGeometry` is a root
  too (the property grid reads `ExceptionReport` by reflection, and `[DynamicallyAccessedMembers]`
  on the class did not keep its rows). A trimming break shows up as a white screen and a
  `CRASH: ...` console line (`Program.cs` prints those). Anything new that is reached only by
  reflection outside those assemblies needs its own root.
- **StyleKeyOverride**: a subclass of a templated Avalonia control (TabControl, TabItem, ListBox,
  Button, UserControl, ColorPicker...) renders invisible unless it overrides `StyleKeyOverride`
  to return the base type.
- **Input**: all WPF-style mouse events funnel through `DynamicGeometry/Behaviors/Behavior.cs`
  (pointer adapters live there). Avalonia has no static Keyboard; modifiers come from event args.
- **Mutating a collection in place does not redraw in Avalonia** the way a WPF Freezable did.
  `Polygon/Polyline.Points` only rebuild geometry when the property gets a different list;
  figures call `Shape.PointsChanged()` (`WpfCompat.cs`) instead. Suspect the same thing whenever
  something "renders once and then never updates". Likewise a `Path` whose figure list became
  empty is not repainted: keep the first `PathFigure` and empty its segments (`AngleArc` does).
- **`PathFigure.IsClosed` defaults to true in Avalonia** (false in WPF), and `IsFilled` to true.
  Every hand-built `PathFigure` must set both; a forgotten one draws a line from the end of the
  path back to its start.
- **A point with an infinite coordinate does not exist** (`Math.Exists(Point)` rejects NaN
  and infinity; `IsValidValue` for a double). Avalonia's layout throws "Invalid Arrange
  rectangle" for a canvas child placed at infinity, unhandled: the desktop app dies and the
  browser app freezes on the spot. `CenterAt`/`MoveTo` in `Utilities` also refuse a non-finite
  place. Anything new that positions a control from figure coordinates must not hand it NaN
  or infinity.
- **`RotateTransform`**: no CenterX/CenterY - Avalonia rotates about `RenderTransformOrigin`
  (the middle by default), and WPF-style centering shifts a tilted ellipse off its center.
- **TextChanged arrives late**: Avalonia raises a TextBox's TextChanged through the dispatcher,
  after a programmatic-set guard is gone. Text editors of the property grid remember the text
  they put in the box (`StringEditor.ShownText`) and ignore a TextChanged carrying it;
  otherwise showing a figure records a property set per text row.
- **Stale Bounds after a load**: zoom to fit measures labels itself (`Measure`) because their
  Bounds are stale right after a load. `Label.MeasureSize` invalidates the Border before
  measuring: its measure stays valid when only the TextBlock inside changed, and answers with
  the old size.
- **The Fluent theme paints hover/selected states on the template's ContentPresenter**, so
  style overrides must target that part (`PropertyGrid/PropertyGridTheme.cs` does).
- **The library's own types shadow framework ones**: `Math`, `Ellipse`, `Polygon`, `Path`...
  In a file-scoped-namespace file a `using X = ...;` alias does NOT win over a type of the
  enclosing namespace - write `System.Math.Max`, `Avalonia.Controls.Shapes.Ellipse` in full.
- **`Style` and `Setter` are ambiguous in the library**: `DynamicGeometry.Style`/`Setter` are the
  WPF shims and shadow Avalonia's. For real Avalonia styles alias them
  (`using AvaloniaStyle = Avalonia.Styling.Style;`).
- **Browser has no system fonts.** Text renders only because `Avalonia.Fonts.Inter` is embedded
  (`.WithInterFont()`). Inter reaches text by *inheritance* from the window (the theme sets it
  there); an explicit family that doesn't exist - and `FontFamily.Default` and
  `FontManager.DefaultFontFamily` too - renders as Noto Mono in the browser. So `TextStyle` sets
  the font a drawing names (Arial, Segoe UI) only when it really resolves, and naming "Inter"
  in a drawing doesn't work (embedded, not installed).
- **The app runs under the invariant culture** (`App.UseInvariantCulture`, first thing in
  both `Main`s; the Browser project also builds with `InvariantGlobalization`, so no ICU
  data is downloaded). Numbers use a point everywhere: the expression language, `.lgf` and
  `.dgf` files, Avalonia path markup, labels. Anything that formats or parses numbers for a
  parser still says `InvariantCulture` explicitly; a `$"M{x},13.5"` under a decimal-comma
  culture makes `Geometry.Parse` throw. To test a culture: `webauto stop`, then
  `webauto start <url> 1280 800 --lang de-DE`.
- **Files in the browser** (`MainView.SaveDrawingToFile`/`OpenDrawingFromFile`): the File
  System Access API takes a file type only as a MIME type with its extensions, and Avalonia
  drops a `FilePickerFileType` without `MimeTypes`. The stream from `OpenWriteAsync` has only
  `WriteAsync`; a `StreamWriter` flushes synchronously on dispose and throws, so the text is
  written as bytes with `WriteAsync`/`FlushAsync`. To test saving headless (no native dialog):
  `webauto eval` a fake `globalThis.showSaveFilePicker` returning a handle with `kind`, `name`,
  `getFile`, `queryPermission`, `requestPermission`, `createWritable`
  (`write`/`close`/`seek`/`truncate`) *before the first picker use* - Avalonia's picker
  polyfill captures the global when its storage module is first imported, later replacements
  are ignored - then click Save and read back the bytes; throw a
  `DOMException(..., "AbortError")` for cancel.
- **Keyboard focus drifts into tool panels.** A tool's PropertyBag panel (e.g. "Point by
  coordinates") takes focus into its TextBox after every construction step, so neither the canvas
  KeyDown nor `MainView_KeyUp` (which skips TextBox focus) sees keys then. Anything that must
  always work (Escape) belongs in the `MainView_KeyDown` tunnel handler. Ctrl shortcuts are
  handled on key *down* (`MainView.HandleControlShortcut`): on key up Ctrl may already be
  released and a bare S is the Segment tool. Plain keys: `MainView.HandlePlainKey`; tool letters:
  `UI/BehaviorShortcuts.cs` (also feeds the tooltips).
- **Every exception is shown** (`MainView.CurrentDomain_FirstChanceException`): status bar plus
  an `ExceptionReport` page in the side panel, also printed to the console. Add to `IsBenign`
  when a framework exception turns out to be noise (the browser's file picker throws
  `JSException` "AbortError..." on cancel). A `JSException` has no .NET stack, so the handler
  captures `Environment.StackTrace` at the throw. Text boxes of the property grid wrap at
  480 px (`PropertyGridTheme`), so a stack trace doesn't stretch the panel across the window.

## Design decisions

- **The view** (`Figures/Coordinates/CoordinateSystem.cs`) is an origin in pixels plus `UnitLength`
  (pixels per unit); every zoom goes through `Zoom(factor, focus)` (the point under `focus`
  stays put), every fit through `Fit`/`SetView`. "Content" for fit (`TryGetContentBounds`) is
  points, whole ellipses (a circular arc only its own extent), labels and the vertices of visible
  segments/polygons/Béziers even when their points are hidden - not lines. Panning (`MoveTo`)
  must not round the origin: a drag is many sub-pixel steps (0.5 px at 200% scaling) and
  rounding each one makes the plane run faster than the cursor. Labels have a fixed *pixel*
  size, so fit re-measures and refits a few times.
- **The grid step adapts to the zoom** (`CoordinateSystem`, "Grid step" region): 1, 2 or 5 times
  a power of ten, with a fainter minor tier between. Shift-snapping lands on the labeled step
  (`MajorGridStep`), not on a fixed 1; `<Viewport GridStep="1">` floors the step for a drawing
  that must keep its unit squares (Pick's Theorem). Whether the grid shows is the drawing's own
  (`CoordinateGrid.Visible`): a new drawing starts without one and a file says. There is no
  global setting: one would leak from every loaded file (each gallery tile included) into the
  next new drawing.
- **No pixel snapping of figure geometry.** Lines are exact and point shapes have
  `UseLayoutRounding = false`, so lines hit their own points. Anything new that must line up
  with figures: same two rules. To check alignment, `winauto shot ... --region x,y,60,60 --zoom 16`
  around a point.
- **What a click makes of a point** is decided in one place, `Behaviors/PointPlacement.Find`
  (existing point / free / on a figure / intersection / midpoint; an existing midpoint is reused,
  never duplicated). Every tool uses it for the click and `Behavior.GetClickPreview` feeds the
  same answer to `Behaviors/ClickPreview` on hover (ghost point, halo on the sources), so cursor,
  preview and click agree. Halos need a case per figure kind in `ClickPreview.CreateHalo`. A
  tool that takes a figure where it would otherwise expect a point (Distance, Circle by Radius)
  overrides `FigureCreator.FindFigureInsteadOfPoint`. The preview is plain canvas visuals, never
  figures. Hit testing uses the *snapped* coordinates. Typed coordinates always give a free
  point. A point must never be placed on the figure being constructed
  (`FigureCreator.CanPlacePointOn`), that would be a dependency cycle.
- **Cursor philosophy** (`Behavior.GetCursor`): cross = a new *free* point appears here; hand =
  the click picks something already there - a figure the tool needs, an existing point, or a
  place defined by figures (intersection, midpoint); arrow = everything else, including a new
  point sliding along a figure and clicks that do nothing. `winauto cursor` prints the cursor
  showing now (screenshots don't include it).
- **Default point styles are by kind** (`StyleManager.AssignDefaultStyle`, looked up by name):
  the draggable kinds (`FreePoint` yellow, `PointOnFigure` green) at size 10, constructed ones
  at 8. A drawing from a file brings its own styles; when the named one is missing (older
  files) the first point style is used. Gallery point sizes follow the same standard:
  `dotnet tools/pointsizes.cs -- <folder> [--apply]` lists and raises undersized styles. The
  Rose's 90 control points stay at 5 px on purpose (at 10 they swallow the flower).
- **Any figure is draggable**: dragging a dependent figure moves its root free points, so gallery
  text can say "drag the circle" even when the points it is built on are hidden. Labels are the
  exception (`AllowMove`): a drag moves the label, not what it measures or names.
- **Numbers are shown with `Settings.DisplayDecimals`** (2) by every editor of the property grid
  and by labels by default; values keep their digits and a typed number is taken as typed.
- **Property grid layout is declarative** (`PropertyGrid/`): `[PropertyGridGroup]` boxes rows
  and their buttons together, `[PropertyGridDestructive]` puts a button last under a divider,
  `[PropertyGridIcon]` puts a drawn icon (`PropertyGridIcons`) in front of a caption (every verb
  button has one - give a new verb one too), `[PropertyGridPreferredEditor("UpDown")]` picks the
  editor, `[Domain(min, max)]` on a double gives a `SliderEditor`. Rows that are editable only
  sometimes: `[PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]` on the
  property and `IConditionalProperties` on the figure (`CanEdit`, `Caption`); the same interface
  vetoes buttons by method name. A property setter that changes the figure list (the length
  panel's Show, the point's Free toggles) does so directly: the grid records the property set as
  the undo step and the undo library refuses an action recorded from inside another.
- **Setting a segment's length stretches it once** (not a constraint); **Fix length** makes the
  constraint: the end becomes a `TranslatedPoint` from the pivot with an auxiliary `Number` at
  the current distance and a free direction (`IFixableLength`, `Figures/Lines/LengthConstraint.cs`;
  segments, vectors, regular polygons, circles). Right after such a figure is made the side
  panel shows a `LengthPanel` (`FigureCreator.ShowCreatedFigure`); the tools themselves have no
  length box. The swap is `Actions.ReplacePoint`, which hands over the point's name label
  (`ReplaceFigureAction` skips it on purpose). A creator's undo transaction spans one
  construction, opened at the first click (`FigureCreator.EnsureTransaction`), so an edit made
  in that panel between constructions is an undo step of its own.
- **Numbers are figures** (`Figures/Values/Number.cs`): roots with no shape, named n1, n2...,
  usable in expressions by name, a length and an angle at once. There is no tool that makes one
  yet; a typed value in the Translation tool becomes one. `Auxiliary` (any figure) marks one
  created on demand for another: it is removed with its last dependent and comes back on undo.
  Every deletion goes through `RemoveFigureAction`, one figure per action inside a transaction,
  so undo restores labels and a polygon loses a vertex rather than dying.
- **TranslatedPoint** (`Figures/Points/TranslatedPoint.cs`): distance and direction are each
  *tied* to a figure (vector, length/angle provider, `ILine` as it points, a Number) or *free*
  (the parameter dragging changes). Roles are stored by index into the dependency list, never
  inferred from types (a `Label` is both a length and an angle provider). A point with a free
  quantity takes the green `PointOnFigure` style. The Translation tool is stepwise (source,
  distance, direction, placement); its panel is cleared in `Stopping` because that event is
  raised before `Started` resets the state, and `DrawingControl` re-shows a behavior's
  `PropertyBag` on "construction complete" as well.
- **Ellipses**: center, end of the long axis, end of the short axis, the last by its distance
  from the long axis, so a point on the short axis is on the ellipse. The tool puts a free third
  click on a hidden perpendicular through the center (`EllipseCreator.FindPointPlacement`), so
  scaling the long axis scales the short one with it.
- **Vectors** are an invisible `Segment` plus an `Arrow` polygon sized in pixels, filled with the
  line color. `Vector.OnAddingToCanvas` sets the default `LineStyle` before the base call,
  otherwise the polygon default (pale fill) wins.
- **Right angle marks** (`Figures/Lines/RightAngleMark.cs`) are a passive visual owned by
  `PerpendicularLineBase` (perpendicular line, segment bisector), deliberately not an angle
  figure. Which corner the mark sits in is *stored* (`Corner`), chosen once and never derived
  from the geometry again - deriving it makes the mark flip-flop on rounding; a click with the
  Drag tool moves it to the next corner. It hides when an `AngleArc` sits at the same vertex,
  because a measured angle of exactly 90° draws the same sign itself.
- **An angle is two figures**, `AngleMeasurement` (the number) and `AngleArc` (the mark, 0-3
  arcs), paired by `AngleArc.FindCompanion`. `DGFReader.ReadMeasureAngle` creates the arc from
  VB6's DrawStyle / AuxInfo(2) - not tested, there is no sample .dgf with an angle in the repo.
- **Dashes**: `LineStyle.Dash` is put on in `LineStyle.OnApplied`, not through a setter, because
  `StrokeDashArray` counts in stroke widths and a selected figure is thicker. Anything that
  applies a style to a shape by hand (sample glyphs) must call `OnApplied` too. Any enum
  property of a style or figure round-trips by name through the generic `EnumSerializer`.
- **Color/brush picking** (`DynamicGeometry/Controls/ColorPicker/`) is layered so parts can be
  swapped: `ColorPalette` -> `ColorPage` (swatches, spectrum) -> `ColorPickerView` ->
  `BrushPickerView` (solid | gradient); in the property grid through `ExpandingPickerEditor`.
  Whoever hosts a `SegmentSwitcher` sets its `Surface` to the background it sits on so the
  selected tab blends into it.
- **The paper** is `Drawing.Background`, edited through the "Background" button on the
  Coordinates tab (`DrawingHost.ToggleDrawingProperties` puts the drawing itself in the property
  grid). A gallery tile takes a drawing's paper as its plate and turns its caption white on a
  dark one.
- **No menu.** New, Open, Save | Undo, Redo are one row (`LiveGeometry/MainToolbar.cs`) above the
  ribbon, with the tour group (◀ n/N ▶ + title) between them and the Octocat at the right,
  whose tooltip is the build. The first button is the app's mark and folds the ribbon (Ctrl+F1,
  `MainView.UpdateRibbon`): folded by default when a gallery drawing opens on a small screen
  (under 700x500), open otherwise; once pressed, the user's choice holds for the session.
  Everything else is keys. Lost their menu entry and are unreachable for now: Lock, Figure List,
  the settings page. Not on the Selection tab (obscure for the audience; the commands and
  settings are still there in `DrawingHost`): Ortho, Polar, Snap to grid, Snap to point, Snap to
  center. Shift while dragging or clicking still snaps to the grid, and a click near the middle
  of a segment still makes a midpoint.
- **Ribbon look** is centralized in `UI/Ribbon/RibbonTheme.cs`; `ButtonGrid` draws the
  hover/pressed/checked plate; `Ribbon`/`TabPanel` replace the Fluent templates in code. To
  bring the group headers closer together, change `ButtonGrid.HeaderOverlap`, not the padding
  inside the tab. An on/off `Command` exposes `IsChecked` (a `Func<bool>`), which its button
  re-reads after any toggle (`CommandToolButton.UpdateToggles`, called by anything that toggles
  from a key) - don't go back to `CheckBox` icons.
- **The splash** (`wwwroot/index.html` + `app.css`, shown until Avalonia adds `splash-close`)
  animates Euclid's first construction, with a progress bar that `main.js` feeds from a
  `withResourceLoader` wrapper counting fetches (the loader's `onDownloadResourceProgress` is on
  the module config, which the .NET 10 host builder doesn't expose). The wrapper hands back a
  Response only for assemblies, the wasm and ICU data; the runtime's own JavaScript modules and
  its config must be left to the default loading (a Response for `dotnetjs` leaves the site
  spinning on the splash forever). A change to `main.js` needs the publish smoke test below
  before deploying. `?splash` on the url shows the splash without starting the app: serve the
  source `wwwroot` with `tools/serve.cs` and open `http://localhost:<port>/index.html?splash`.
  It animates regardless of `prefers-reduced-motion` on purpose: the query follows the Windows
  "Animation effects" setting, off on this machine, and a still splash looks stuck.
- **web.config**: `LiveGeometry.Browser/web.config` is hand-written (serves the precompressed
  `.br` files, sets immutable caching on fingerprinted assets, `no-cache` on entry files). The
  wasm SDK drops a project web.config from publish, so the csproj copies it with an explicit
  `AfterTargets="Publish"` target. It sits at the publish root and rewrites into `wwwroot\`.
  There is no IIS locally: verify changes after deploy with
  `curl -s -o /dev/null -D - -H "Accept-Encoding: gzip, br" https://livegeometry.com/_framework/<file>`
  and expect `Content-Encoding: br` and `Cache-Control: public, max-age=31536000, immutable`.
  A malformed web.config takes the whole site down (HTTP 500).

## File format (.lgf)

- **Colors are always `#AARRGGBB`** (`ColorText.ToArgbHex`). Never `Color.ToString()`: Avalonia
  writes the *name* of a known color. `ToColor()` accepts names too, for old files. A gradient
  is a child element (`<Fill><LinearGradientBrush>`, `<Background>...` on `<Viewport>` for the
  paper), not an attribute.
- **`<Drawing Version="1">`** (`Settings.CurrentDrawingVersion`) says label offsets are pixels:
  a `LabelWithOffset` (point labels, measurements) sits at `Offset` pixels from its anchor in
  the plane, so the gap keeps its size at every zoom. A file without the version is upgraded on
  load, after the viewport, at the zoom it opens at (`UpgradeOffsetFromUnits`) - the best guess,
  since files don't say. Anything new that places text by something in the plane should keep
  its distance in pixels the same way. A point label is also kept in an orbit around its point
  (`PointLabel.ClampPosition`).
- **`<Drawing IntersectionOrder="Legacy">`**: `Math.GetIntersectionOfCircleAndLine` once swapped
  P1/P2 for a line through the center, and files carried no version then, so old drawings (the
  phone ones) that pick the other intersection opt in to a swap on load
  (`IntersectionPoint.UpgradeLegacyCircleAndLineOrder`). No file in the repo carries the mark
  (`LiveGeometry.Desktop.exe --modernize <folder>` writes into a file what loading it upgrades);
  the code stays for old files from elsewhere.
- **Saved files declare `encoding="utf-8"`**; files from older builds say `utf-16`, which our
  own loader tolerates but `XDocument.Load` does not.
- **`TranslatedPoint`** without `DistanceSource`/`DirectionSource`/`FreeDistance`/`FreeDirection`
  is the old format (typed values as attributes, roles by position, direction in radians);
  `UpgradeLegacyValues` gives the typed values auxiliary Numbers.
- **Pinned labels** (`Label.Pin`, `Figures/Controls/LabelPin.cs`): `Pin="TopRight" OffsetX
  OffsetY` is the pixel distance from a canvas corner to the *same* corner of the label, so it
  keeps that edge while the plane zooms and pans under it. `Coordinates` always tell where it
  is in the plane right now, so hit testing, dragging and the property grid need nothing
  special. `WrapWidth` (px) and `Backdrop` (a plate of the paper's color) exist for captions.
  Pinned labels draw above figures, zoom to fit ignores them, thumbnails hide them.
- **Scenes** (`<Scene Left Top Right Bottom />` under `<Drawing>`, 1 or 2) are opt-in suggested
  views for drawings whose content has no useful bounds (endless ground: Castle, The Falling
  Ladder; a locus: Spiral). Where they exist, fit and tile show the scene nearest in shape to
  the room instead of the content bounds, and a gradient paper spans the scene, not the canvas.
- **`.dgf` paper**: `PaperColor1`/`PaperColor2`/`GradientPaper` of `[General]`, top to bottom.

## Gallery

- **Gallery and routes** (`LiveGeometry/Gallery/`, `MainView` "Pages" region): three states -
  gallery, a gallery drawing (`CurrentSample`: tour group in the toolbar, Page Up/Down, edits
  dropped silently, Save = save as and the drawing becomes the user's own), the user's own
  drawing (parked in `OwnDrawing` with its undo history while they look around). Paths `/`,
  `/gallery/<slug>`, `/drawing`; `AddressBar` is the abstraction, `BrowserAddressBar` +
  `main.js` do pushState/popstate. `index.html` needs `<base href="/">` for that (all routes
  serve it; `web.config` and `tools/serve.cs` fall back to it). All page changes go through
  `MainView.Show*`. Started at a drawing's address the gallery isn't built at all (the address
  is readable synchronously since `main.js` registers its imports before the runtime starts).
- **Gallery drawings** are embedded `.lgf` (`Gallery/Drawings`, order and titles in
  `GalleryCatalog`) and are *the* source: edit them directly. They were forked once from the
  Windows Phone samples and the DG 1.0 CD library; `tools/gallerize.cs` made that fork and must
  not be rerun (it would overwrite the hand-edited drawings). To add another CD drawing:
  `--check` it (below), take the `.lgf`, drop the old labels, add `Title`/`Description` labels
  in the `GalleryTitle`/`GalleryText` styles (copy from any drawing here), set `Grid`, list it
  in the catalog. Drawings with `Grid="true"` (graphs) keep their file viewport in view
  (`GalleryItem.Plane`), because graphs and lines have no bounds.
- **The caption** (`Title` + `Description`, which may contain live expressions - then list the
  points as dependencies) is pinned to the screen; the files carry no X/Y and
  `GalleryDrawing.Fit` recomputes pin, offsets and wrap width at open time: a column at the
  right in a wide canvas, a strip at the bottom in a tall one. Labels are sized in pixels, so
  the figure gets the canvas minus the text and the zoom is computed from that in one go -
  never iterate "place text, zoom to fit": it runs away once the text needs more than its
  share. If the text doesn't fit it runs off the bottom and the reader drags it up. Refitted on
  resize until the first edit, through `Drawing.SizeChanged` (`MainView.KeepFitted`) and not the
  canvas's event: the coordinate system's own resize handler shifts the origin and has to run
  first. A file with a caption opened from disk is fitted the same way. Explanations have no
  hand line breaks (a blank line separates paragraphs). On an iPhone in portrait (canvas about
  390x565) the long explanations run off the bottom.
- **Tiles are live drawings**, not bitmaps (`DrawingThumbnail`): no Behavior, text hidden, loaded
  one per idle tick; hovering makes the draggable points drift. A load error of a tile goes to
  the console as `Gallery: <file>: ...` - the quickest way to check all drawings at once.
- **Generated drawings** - regenerate rather than edit the file: Line of Best Fit
  (`dotnet tools/bestfit.cs -- <the .lgf>`), Fibonacci Spiral (`tools/fibonacci.cs`),
  5 Platonic Solids (`tools/platonic.cs`, `--alpha` for a translucent variant). Hidden
  `PointByCoordinates` whose coordinates are expressions are the library's variables.
- **Drawings that must not fall apart** when a kid drags the wrong thing (Castle, The Falling
  Ladder): fixed points are `PointByCoordinates` with constant coordinates (a polygon of those
  has no free point, so dragging it does nothing), and the only things that move are
  `PointOnFigure` sliders on hidden rays or segments plus a free point or two.
- **The Spiral's rings** ("Drag to here") are placed for `Locus.StepCount` = 60 samples;
  changing `StepCount` moves them.

## Deployment and caching

- A push to main takes about 8 minutes to go live (GitHub Actions: build ~5, deploy ~3); the
  files flip at the very end. Until then every reload, cached or not, gets the previous version.
  The commit in the tooltip of the Octocat at the right end of the toolbar (`BuildVersion`, also
  logged to the console at startup) tells which build is on screen.
- Entry files are `no-cache` with an ETag that changes on every deploy, so a browser holding the
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
`.lgf`/`.dgf` under the folder in the real editor, zooms to fit (a captioned drawing is laid
out by `GalleryDrawing.Fit` instead, with the ribbon folded when the window is small - so the
sheet shows the gallery layout at whatever size the window was last closed at: `place` it,
close it, then run the check), saves `<out>/<relative path>.png`
and `.lgf` (the conversion) and appends to `<out>/report.txt`: figure counts, `NOT EXISTING`
(only the *root* failures - figures whose dependencies all exist - plus a dump of every point),
load errors. It exits when done. `dotnet tools/contactsheet.cs -- <png folder> <out.png>
[columns] [tile width]` tiles the PNGs into one image: the fastest way to eyeball a whole
folder (the whole gallery fits on one 4-column sheet). The VB6 CD library
(`C:\Dropbox\Projects\DG 1\DG CD Version 1.0\Library\English`, 221 files, read-only) all
loads; what the report still calls "missing" there is second intersections that fall outside
a segment or ray, and sides of a polygon that don't cross - legitimately absent.

## .dgf (DG 1.0) reader facts

Learned from `Reference/VB6/Source` while making the CD library load (`DGFReader.cs`):
- `modFileIO.bas` writes an `AuxInfo(n)` / `AuxPoints(n).X/Y` only when it is not 0: absent
  means 0 (`GetAuxInfo`/`GetAuxPoint`). An analytic line `x - y = 0` has no `AuxInfo(3)`.
- A point's `Type` is the DrawState that made it; 0 = free point, and old files write
  `ParentFigure=0` for those (which is *not* Figure0). Every dependent point has `Type != 0`.
- `ParentFigure` of a point and figure indices are 0-based; Points, Labels, Buttons 1-based.
- Buttons: type 0 show/hide, 1 message box, 2 sound, 3 open another drawing. Only 0 becomes a
  figure; a show/hide button's list may reference the others. The list goes into the
  `ShowHideControl`'s own `Dependencies` (what it shows and hides, and what the `.lgf` saves);
  `AddDependencies` alone only registers the dependents side and the boxes do nothing.
- DG's angle bisector (`Math.bas GetBisector`) is the *interior* bisector and a whole line;
  the reader sets `AngleBisector.Interior` and `IsLine` ("Whole line", saved as `Line="true"`).
  `Interior` ("Inside the angle", saved as `Interior="true"`) halves the angle under 180°
  whichever way round the sides are; off, the bisector halves the angle counterclockwise from
  side 1 to side 2, which swings outside a triangle dragged the other way round. New
  bisectors are interior; a file without the attribute gets the oriented one it was saved
  with. Morley's trisectors are expressions and get the same effect from
  `SGN(pi - OANG(...)) * ANG(...)` - `OANG` is the counterclockwise angle in [0, 2pi).
- A point on a figure is placed by moving it to its saved X,Y in one `MoveTo` (setting X then Y
  projects twice from off the figure and lands elsewhere); `AuxInfo(1)` (t on a line, clockwise
  angle on a circle) is ignored.
- `AuxPoints(6)` of a measurement is the label's shift from its default place, in pixels, y
  down - taken as our `Offset` as it is.
- Intersection points are picked by the saved coordinates of both solution points, so the
  Legacy circle/line order problem does not apply to `.dgf`.
- The VB6 expression language has more than ours: `[A,B]` distance with a comma, comparison and
  logical operators, `IF`/`MAX`/`MIN`, `°` inside brackets. Those labels show the error text;
  the compiler never throws out of a label (`Compiler.CompileExpression` catches).
- `IniFile` skips blank lines (every CD file has them between sections).

## UI automation (tools/)

Screenshots are PNGs; image pixels are the click coordinates in both tools.

- `tools/winauto.cs` - any desktop window (the VB6 app, the Avalonia desktop app).
  `list`, `tree <t>`, `menu <t>`, `invoke <t> <menuId>`, `shot <t> out.png`, `click <t> x y [right|double]`,
  `drag <t> x1 y1 x2 y2 [steps]`, `wheel <t> x y <notches>`, `keys <t> "^s"`, `text`, `focus`,
  `cursor`, `place <t> x y w h`. Target = process name | `pid:N` | `hwnd:0x..` | `title:substr`.
  - The VB6 app starts maximized on a 4K/200% monitor: `place <t> 100 100 1500 1000` first. The
    Avalonia desktop app remembers its window (`LiveGeometry.Desktop/WindowPlacementPersistence.cs`,
    `%LOCALAPPDATA%\LiveGeometry\MainWindowPosition.txt`). It is left at 100,100 1700x1100 for
    testing - `place` is only needed again if someone resized it. Close test instances with
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
    open, `shot LiveGeometry.Desktop` captures the menu, not the main window.
  - `winauto keys` sends real virtual keys for lowercase ASCII letters/digits (needed for the
    single-letter shortcuts); other characters go as Unicode packets, which KeyDown-based
    shortcuts never see.
- `tools/webauto.cs` - the browser build in headless Edge over CDP (port 9333, Edge stays alive
  between calls). `start <url> [w h] [--lang xx-XX]`, `stop`, `nav`, `wait <console text>`,
  `console [--errors]`, `shot`, `click`, `drag`, `move`, `key <Key> [ctrl] [shift] [alt]`,
  `text`, `eval <js>`.
  - The app takes several seconds to boot after `start`; the splash `div` stays in the DOM, so
    don't test for its absence - take a screenshot.
  - Edge runs with `--guest`; without it Edge signs the throwaway profile into the Windows
    account and opens a sync dialog as an extra page target.
- `tools/serve.cs <publish>/wwwroot [port]` - static server for a publish output (blocks; run in
  the background). Separate file because a running server locks its own exe.

Web smoke test (run before deploying changes that touch file I/O, clipboard, fonts, or anything
reflection-based): publish Release, `serve` its wwwroot, `webauto start http://localhost:5005/`,
`console --errors` must show no `CRASH:`, draw a segment via the Lines tab, `shot`, `stop`.

## VB6 parity backlog

Deliberately out of scope for now: Calculator, step-by-step construction playback.

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

PNG export, printing, demo download, and the `Hyperlink` figure's use of `WebClient`.
Saving through the browser storage provider works; opening has only been checked as far as
the picker (see "Files in the browser").
