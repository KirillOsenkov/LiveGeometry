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
- `Reference/VB6/Source/` - source of the original VB6 app ("DG"). A built copy may exist outside
  the repo (process name `Geometry`).
- `tools/` - file-based C# tools (`dotnet run tools/x.cs -- args`), see below.

## Working rules

- The maintainer reviews and commits everything: never commit, push, or otherwise change git
  state. Leave changes in the working tree, no need to report them.
- Scripting is C# file-based apps (`dotnet run file.cs`). Don't assume Python is installed.

## Code style

- **Apply these rules to new code only.** Keep diffs minimal - don't reformat or rename existing
  code as a drive-by. Most of `DynamicGeometry/` is decades-old WPF/Silverlight-era code with
  block namespaces and LF endings; it stays that way unless a cleanup is explicitly asked for.
- **New files: CRLF line endings, file-scoped namespaces.** Existing files keep what they have.
  The Write tool emits LF: after creating a file, or rewriting an existing CRLF file with Write
  rather than Edit, convert its line endings with an editor tool. Never grep for `\r`.
- **Never edit workspace files from the shell** (sed/perl/heredocs). Use the editing tools.
- **Backwards compatibility is not a concern**: every caller is in-tree, so change all call
  sites rather than keep a worse shape. That goes for the `.lgf` file format too: as of
  2026-09-29 there are no drawings to stay compatible with other than the gallery's, so the
  format may change freely as long as the gallery drawings are updated to it (edit them, or
  `--rewrite` the folder) and still load. Don't add upgrade paths for old files; the ones the
  loader has are from before this rule.
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
- Build with `dotnet build` / `dotnet publish` (there is no WPF markup compilation in the
  Avalonia solution). Pass `-bl` and read the binlog with the binlog MCP
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

## The ribbon

Tools are `Behavior` subclasses found by reflection (`Behavior.LoadBehaviors`): `[Category]`
names the tab, `[Order]` the place in it, `[Ignore]` keeps one off the ribbon; the button shows
`Name`, the tooltip adds the letter from `UI/BehaviorShortcuts.cs`, the status bar shows
`HintText`. Tab order is the order of the constants in `BehaviorCategories`; tools the user
defines go onto Misc. Non-tool commands are added in `MainView.InitializeCommands`. Tab by tab
(letter in parentheses):

- **Selection**: Drag (Q) - drags points and figures (with Alt a point snaps onto figures and
  lets go of them, see "Snapping and releasing points"); also the tool every construction
  returns to. Figure List (toggle): see "The Figure List".
- **Points**: Point (P) - free, on a figure, or at an intersection; Midpoint (M) - two points
  or a segment; Intersection (I) - two figures that cross, the click on the second picks the
  nearer crossing (`PointPlacement.Intersection`, shared with the Point tool); Coordinates (X) -
  the Point tool with its X/Y panel always on (`PointByCoordinatesCreator`), so a point by
  coordinates needs no trip to the Coordinates tab's toggle (that toggle stays for typing
  the points of other figures); Label new points (toggle).
- **Lines**: Segment (S), Ray (Y), Line (L), Vector - two points each; Parallel (N) and
  Perpendicular (E) - a line then a point; Perpendicular Bisector - two points or a segment; Angle Bisector
  (B) - vertex then two side points, or an angle measurement; Line at Angle - a point, at the
  angle in the tool's panel (0 until changed, so a horizontal line is one click), or click an
  angle measurement, its arc or a slider first to tie the angle to it; right after the click
  the panel shows the new line's angle instead (see "Tied values"). Join segments (a point between
  two segments joins their other ends) and Polyline (points, double-click or click an
  existing point to finish) exist but are `[Ignore]`d as rarely used.
- **Circles**: Circle (C) - center then a point on it; By Radius (R) - two points, a segment, a
  distance or a slider, then the center; a first click on empty paper makes a slider for the
  radius (see "Sliders"); Ellipse - center, end of the long axis, end of the short axis;
  Circular Arc (A) - center, start, end (counterclockwise); Elliptical Arc - center, semi-major,
  semi-minor, begin angle, end angle.
- **Shapes**: Triangle - 3 points; Square - two adjacent vertices; Polygon (W) - points, then
  Enter, a right-click or a click on a vertex closes it; Regular polygon - center then a vertex.
  Triangle and Polygon show no length panel (a side's length means nothing for them). They
  (and Square, for its first side) draw their sides as segments (the polygon's own outline is transparent in the default
  style), except where a visible segment, ray, line or vector on the two points is there
  already (`FindLine`; not any line that depends on both - a perpendicular bisector of the
  two took the side's place and left it undrawn). (Polygon intersection
  exists but is `[Ignore]`d.)
- **Transform** (the tools are verbs, as the tab is; the classes stay `ReflectionCreator`...):
  Reflect (T) - source figure, then a mirror (point, line, segment, ray, or a
  circle for a point source); Rotate - source, center, angle (a figure with an angle or a
  typed value); Translate - source, distance, direction (see "TranslatedPoint"); Dilate -
  source, center, factor (a figure with a length or a typed value); after each of the three
  the panel shows the new figure's values (see "Tied values"). Last of the tabs that
  draw with figures alone; the two after it work with numbers. A figure is transformed by
  transforming what it is built on, down to points (`Transformer`), so a source is taken
  only when all of that can be (`CanBeTransformSource`, recursive) - except a radius given
  by a number (a slider, a Number, a measurement: By Radius makes a slider whenever its
  first click is on paper), which the image of a reflection, rotation or translation
  shares and a dilation can't scale, so Dilate leaves such a circle alone. Asked only
  about the figure itself, the tool threw at its last click.
- **Coordinates**: Background and Grid (G) (commands); Function - an expression in x; Line - by
  slope and intercept expressions; Circle - by center and radius expressions; Point by
  coordinates (toggle: gives the point tools an X/Y panel).
- **Measure**: Distance - two points or a segment; Angle (J) - vertex then two side points,
  the angle under 180° whichever side comes first (the tool orders the sides; an angle goes
  counterclockwise from its first side, and clicked the other way round a triangle's angle
  said 270°; "Convert to opposite angle" gives the other); Area (K) - a polygon, ellipse,
  circle or list of points, Enter or a right click when the points are done; Slider - where it sits, then where
  its knob starts (or press, drag, release): a number with a handle, taken wherever a tool
  asks for a length or an angle, named in expressions (a, b, c).
- **Misc**: Bezier - four points; Locus (D) - a point that depends on a point on a figure;
  Text - a label at the click; Define figure - records a construction as a new tool: click
  the figures it starts from, OK, click the figures it makes, Create tool (neither step
  goes on with nothing picked). The new tool lands on Misc; it is not kept between runs
  (`ToolStorage` is a stub) nor saved with a drawing, and a result whose text names
  figures (a label's [AB]) still means the figures it was recorded on.

## Avalonia and framework traps

- **Trimming only happens on Release publish**, never in `dotnet run`. The geometry library is
  discovered via reflection (`GetTypes()` for behaviors, serializers, figures), hence
  `TrimmerRootAssembly DynamicGeometry.Avalonia` in the Browser csproj; `LiveGeometry` is a root
  too (the property grid reads `ExceptionReport` by reflection, and `[DynamicallyAccessedMembers]`
  on the class did not keep its rows). A trimming break shows up as a white screen and a
  `CRASH: ...` console line (`Program.cs` prints those). Anything new that is reached only by
  reflection outside those assemblies needs its own root.
- **Inherited attributes come twice in the browser**: Mono's `inherit: true` lookup on a
  property declared in a base class (`Circle.Length` from `CircleBase`) adds the base
  property's attributes again. With `AllowMultiple = true` they aren't deduplicated and
  `Attribute.GetCustomAttribute` throws AmbiguousMatchException; CoreCLR (desktop) returns
  one. An attribute read singly must say `AllowMultiple = false`.
- **StyleKeyOverride**: a subclass of a templated Avalonia control (TabControl, TabItem, ListBox,
  Button, UserControl, ColorPicker...) renders invisible unless it overrides `StyleKeyOverride`
  to return the base type.
- **Input**: all WPF-style mouse events funnel through `DynamicGeometry/Behaviors/Behavior.cs`
  (pointer adapters live there). Avalonia has no static Keyboard; modifiers come from event args.
  Fingers (`PointerType.Touch`, "Touch" region there) don't go to the tool as they come: a
  finger's press is held back until it is a tap (lifted where it came down), a drag (gone
  `TouchSlop` pixels: the tool gets the press where it came down, then the moves) or one of
  two fingers, which zoom and pan the view (`CoordinateSystem.PanAndZoom`, the paper under
  the fingers stays under them) and of which the tool hears nothing. Passed on as they came,
  a second finger was a second press: the Drag tool's view jumped between the fingers, any
  other tool left points behind. While a tool handles a finger, `Math.CursorTolerance` is
  `TouchTolerance` (10 px, a finger reaches further than a mouse); whatever sizes things in
  pixels must not take the cursor's tolerance for it (`PointLabel.Margin` does now: a label
  placed by touch sat further out). No long press yet: what the right button does (the
  context menu) has no touch equivalent. `webauto tap/touch/pinch` drive it headless.
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
  or infinity. A size neither: `Width` throws on infinity right in the setter, and a
  figure's `UpdateVisual` runs also while it doesn't exist (a circle around a center that
  is nowhere has an infinite radius: the exception ended a drag, which then could not be
  undone; `CircleBase`/`EllipseBase.UpdateVisual` return early).
- **`Shape.Render` is sealed** (Avalonia 12): a Shape can't draw anything but its geometry.
  `PointMarker` draws a character through an `EmojiGlyph` visual child it measures and arranges
  itself; with no geometry, `Shape.ArrangeOverride` returns size 0 and the shape collapses to
  the middle of its place, so the override returns the final size then.
- **A `FontFamily` in a collection that isn't registered yet** ("fonts:Emoji#...") makes text
  layout throw *while rendering*, which ends the desktop app. `EmojiFont.Family` is the default
  family until the font has loaded.
- **`RotateTransform`**: no CenterX/CenterY - Avalonia rotates about `RenderTransformOrigin`
  (the middle by default), and WPF-style centering shifts a tilted ellipse off its center.
- **TextChanged arrives late**: Avalonia raises a TextBox's TextChanged through the dispatcher,
  after a programmatic-set guard is gone. Text editors of the property grid remember the text
  they put in the box (`StringEditor.ShownText`) and ignore a TextChanged carrying it;
  otherwise showing a figure records a property set per text row. Only until the user's first
  edit: after that the same text is theirs (typed and deleted back to empty must reach the
  property). A commit (Enter, leaving the box) applies the text itself rather than trust the
  pending TextChanged, since a panel's Enter handler (Add point) reads the property next.
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
- **Files in the browser** (`MainView.SaveDrawingAs`/`OpenDrawingFromFile`): the File
  System Access API takes a file type only as a MIME type with its extensions, and Avalonia
  drops a `FilePickerFileType` without `MimeTypes`. The stream from `OpenWriteAsync` has only
  `WriteAsync`; a `StreamWriter` flushes synchronously on dispose and throws, so the text is
  written as bytes with `WriteAsync`/`FlushAsync`. To test saving headless (no native dialog):
  `webauto eval` a fake `globalThis.showSaveFilePicker` returning a handle with `kind`, `name`,
  `getFile`, `queryPermission`, `requestPermission`, `createWritable`
  (`write`/`close`/`seek`/`truncate`) *before the first picker use* - Avalonia's picker
  polyfill captures the global when its storage module is first imported, later replacements
  are ignored - then click Save and read back the bytes; throw a
  `DOMException(..., "AbortError")` for cancel. Opening headless works the same way: copy the
  file into the served `wwwroot`, `eval` a `globalThis.showOpenFilePicker` returning one
  handle whose `getFile` gives a `File` built from `fetch('/name.ggb')`, then `key O ctrl`.
- **Keyboard focus drifts into tool panels.** A tool's PropertyBag panel (e.g. "Point by
  coordinates") takes focus into its TextBox after every construction step, so neither the canvas
  KeyDown nor `MainView_KeyUp` (which skips TextBox focus) sees keys then. Anything that must
  always work (Escape) belongs in the `MainView_KeyDown` tunnel handler. When the focused
  control leaves the tree (a panel rebuilt around its button), Avalonia focuses nothing and
  keys go to the window alone, past MainView; `MainView.TopLevel_KeyDown` catches that,
  refocuses the canvas and forwards the key. Ctrl shortcuts are
  handled on key *down* (`MainView.HandleControlShortcut`): on key up Ctrl may already be
  released and a bare S is the Segment tool. In a text box Ctrl+C, V, X, Z, Y, A are the
  box's own; S, O, N and F1 are not (they did nothing there, and the browser took them),
  and the canvas takes the keyboard first so that the typed text is committed. The letter
  of a Ctrl shortcut is skipped on its way up (`shortcutKeyDown`) - until it is pressed
  again on its own, since after Ctrl+S or Ctrl+O its release goes to the file dialog.
  Cmd does what Ctrl does (`MainView.IsCommandModifier`, `Behavior.IsCtrlPressed`): in the
  browser on a Mac every shortcut was dead, Cmd+S saved the web page, and Ctrl+click there
  is a right click. Ctrl+Shift+Z redoes as Ctrl+Y does, and Backspace (a Mac's "delete"
  key) deletes the selection as Delete does.
  Plain keys: `MainView.HandlePlainKey`; tool letters: `UI/BehaviorShortcuts.cs` (also feeds
  the tooltips). The keys that move the view (arrows, Home, Page Up/Down, +/-) do nothing
  while the side panel has the keyboard: its lists, sliders and combos use them, and an
  arrow that picked the next style also panned the canvas.
- **Every exception is shown** (`MainView.CurrentDomain_FirstChanceException`): status bar plus
  an `ExceptionReport` page in the side panel, also printed to the console. Add to `IsBenign`
  when a framework exception turns out to be noise (the browser's file picker throws
  `JSException` "AbortError..." on cancel; `GeoGebraReader.LeftOutException` is how that
  reader leaves out an element it can't read, and says so in the status). So never use a
  caught exception as a test (decode
  as UTF-8 and catch, as `DecodeLegacyText` once did): check instead. `--check` runs print
  every one, and should print none. One that escapes an event handler or a posted job is
  then swallowed (`Dispatcher.UIThread.UnhandledException` in `App.Initialize`), so a bug
  shows as an error report and the app keeps going. Not tried on the layout exceptions of
  "A point with an infinite coordinate", which may come back on every layout pass: still
  never hand layout NaN or infinity. A `JSException` has no .NET stack, so the handler
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
  size, so fit re-measures and refits a few times. What the user asks for as "zoom to fit"
  (double click, H, Home, the context menu) is `Drawing.ZoomToFit`: the drawing's
  `FitToWindow` when whoever showed it set one (a drawing of the gallery:
  `GalleryDrawing.Fit`, with its plane - only there, since it lays out the caption, which
  in the user's own drawing is theirs and saved), else a scene if it has scenes, else
  `ZoomExtend`. Plain `ZoomExtend` put
  a gallery figure under its caption, zoomed a graph onto its few points and showed the
  Castle's points instead of its scene.
- **The grid step adapts to the zoom** (`CoordinateSystem`, "Grid step" region): 1, 2 or 5 times
  a power of ten, with a fainter minor tier between. Below a step of 1 the whole-number axis
  labels are bold (`AxisLabel.SetEmphasis`). Shift-snapping lands on the labeled step
  (`MajorGridStep`), not on a fixed 1; `<Viewport GridStep="1">` floors the step for a drawing
  that must keep its unit squares (Pick's Theorem). Whether the grid shows is the drawing's own
  (`CoordinateGrid.Visible`): a new drawing starts without one and a file says. The axes
  are a second switch for a grid that shows (`ShowAxes`, `Axes="false"` in the file also
  while the grid is hidden: unticked then, it was an undo step that changed nothing saved,
  and ticked it brought the axes up without the grid until the file was opened). An axis is
  drawn as a vector's `Arrow` over an invisible line (`Axis`), the head where it leaves the
  window to the right or the top. There is no
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
  point, or the point that is exactly there (`FigureCreator.AddTypedPoint`, for both
  coordinate panels: the one of the Shapes tools looked under the last mouse click instead,
  so a vertex typed after a clicked one found the clicked one again and was dropped; and a
  point within a cursor's reach of the typed place was taken for it). A
  point must never be placed on the figure being constructed
  (`FigureCreator.CanPlacePointOn`), that would be a dependency cycle. A figure that
  doesn't exist right now is not hit (`FigureList.HitTest`, `HitTestMany`): its own
  `HitTest` still answers from where it was last, and a tool took the invisible
  intersection or the circle built on it - what was made on it never appeared, and a point
  put "on" it was saved at (0, 0). And a figure doesn't exist while what it is built on
  doesn't: a figure whose `Recalculate` decides `Exists` itself (by its own coordinates
  being numbers) must ask `Dependencies.Exists()` too, as the transformed points, the point
  by coordinates and the angle bisector do now - the image of an intersection that had
  gone stayed on screen, frozen at its last place, with everything built on it. A point by
  coordinates whose expression doesn't compile (it names a figure that isn't there) doesn't
  exist either: it stood at (0, 0).
- **The reach of a click is in pixels**, the cursor's tolerance plus half the stroke, for
  every figure. A circle, ellipse or arc is hit by its distance from the curve along the
  ray from the center (`Math.RadialDistanceToEllipse`); it was the left side of the
  ellipse's equation minus 1, a number without units, against a length - a circle of
  radius 100 took every click on the screen, one of radius 0.2 almost none. The `math`
  mode of the harness checks it (2 px beside every figure a point can sit on hits, 16 px
  doesn't).
- **A first click on what a tool doesn't take is no step**: a tool that wants a figure
  (Parallel, Rotate, Locus...) and is clicked on empty paper starts no construction
  (`FigureCreator.DropEmptyTransaction`). It did: the Undo button lit up on an empty
  drawing, the first Ctrl+Z only put the tool back, Redo, Delete and Paste did nothing. A
  click that stands for several dependencies (a segment for the Perpendicular Bisector's
  two points, an angle for the Angle Bisector's three) does so only as the first click.
  A tool that may take a point again (`CanReuseDependency`) says where a repeat makes
  nothing (`FigureCreator.IsDegenerateRepeat`): a double click made circles and arcs of
  radius 0. Not a blanket "not twice in a row" - By Radius centers its second circle on
  the second end of its radius (the equilateral triangle), and an elliptical arc may
  begin at the end of the axis just clicked.
  An angle's mark is no length (`IFigureExtensions.GivesLength`: its size is in pixels and
  changes with the zoom); tools and tied values ask `GivesLength`/`GivesAngle`.
- **Figures are named after their points** (`FigureBase.NameFromDependencies`): segment, ray,
  line through two points and vector AB, polygon ABC (up to 10 vertices), polyline, Bezier
  ABCD; a clash gets a number (line AB next to segment AB is AB2), and the numbers go by
  the order of the drawing's list, worked out from the drawing as it is
  (`FigureBase.SettleDefaultNames`: after a figure enters, leaves or moves in the root
  list, and at the end of a rename wave, for the group named after the same points; while
  a file is read it waits, `Drawing.KeepsNamesAsRead`, since expressions are compiled by
  the file's names). Handed out to whoever asked first, AB2 stayed AB2 after AB was
  deleted until the file was opened again, and undo of a join gave segment AB and line
  AB2 their names back the other way round. The points are read in
  the order that comes first alphabetically (A-Z, then A1-Z1) among the readings that name
  the same figure: a polygon from any vertex either way round (ECBA is ABCE), a segment,
  line, polyline or Bezier either way, a ray or vector only as it goes (`PointOrder`); any
  of those readings counts as a default name. Everything else is numbered
  by type (Circle1); hidden points too, so a helper doesn't take a letter from the points on
  screen. `HasDefaultName` (nobody typed a name) is not stored: a name that reads like the
  default is the default, old `Segment1` included, and loading renames those (so does a
  paste: a copy is numbered by type while it is read, and takes its points' name once it
  is in the drawing). A default name
  follows its points - renamed, replaced, joined, a vertex deleted - and undo needs nothing
  special, since it restores the cause. The grid's Name box (`NameEditor`) refuses an empty
  name, one another figure has (case matters: slider `a` next to point `A`), and one that
  reads like a default name the figure would not keep (AB3 or BA for segment AB: it would
  be AB again at once, `FigureBase.KeepsTypedName`); the setter
  itself would take the name from the other figure, and turns an empty one into the default.
  The grid's title is `Title`: the figure's `Kind` in front of the name ("Triangle ABC",
  "Point A"), left off when the name says it already (Circle1, Bezier3) or there is no kind.
  A polygon's kind is by vertex count (Triangle, Pentagon, Hexagon, else Polygon); four
  vertices go through `Quadrilaterals.Classify` (Square, Rectangle, Rhombus, Parallelogram,
  Kite, Trapezoid) with a rounding-only tolerance: a shape dragged to look square by eye
  stays a Quadrilateral. The title is refreshed on every move (`PropertyGrid.RefreshNumbers`).
  `ToString()` stays the bare name: messages and dumps use it.
- **A name's index is a subscript on screen** (`Figures/NameDisplay.cs`): `A_1` draws as A₁
  on the canvas (point labels, slider captions) and in `Title` (the grid's header, the
  Figure List), as GeoGebra and TeX read an underscore; the trailing digits of a name
  without one (`A1`, our own default names after Z, `n1`, `Circle1`) draw as a subscript too.
  The name itself stays as typed everywhere else - files, expressions (`A_1.X` parses, `_` is
  a letter to the scanner), the Name box. What follows an underscore is the run of letters
  and digits, or anything in braces; Unicode has the ten subscript digits and a few letters
  (aeoxhklmnpst), and a part it can't write stays as typed. The browser's embedded Inter
  font has the subscript glyphs.
- **Expressions follow renames**. Compiled expressions hold the figures themselves, but their
  *text* (saved, edited) holds names. `FigureBase.Name` collects a rename wave (the figure,
  the default names that follow it, one that had to give up the name) and at the end of the
  outermost set rewrites every `IRenamableExpressions` in the dependents closure: label
  `[...]` parts, function graphs, `DrawingExpression`s (point by coordinates, line and circle
  by equation). `ExpressionRenamer` reads names as `ExpressionTreeBuilder` binds them, under
  the old names (A.X, AB as two points, the points of ang/dist/area, a Number) and swaps them
  all at once; two points whose new names run together would read differently (PB next to a
  point named PB) become `dist(P, B)`. Only the text changes, and not through undo: undoing the rename renames back.
  A name may end in primes (A', A'': `Scanner.IsName`, the one rule of what an expression
  can say), and the Name box refuses a name no expression can say only for a figure an
  expression names, directly or through a figure named after it - renamed "my point", the
  label `[A.X]` became an error; a caption for a name ("Drag me!", as the Ladder has) is
  fine elsewhere.
  `pi` and `e` in lowercase are the constants whatever the drawing has; in any other case
  (`PI`) the two-points reading comes first, and with points P and I that is their
  distance. (Lowercase too, until 2026-10-01: `sin(pi * x)` came out wrong without a word
  in a drawing with points P and I. Generated expressions still say `rad(45)` or the digits
  of π.) A number (slider, Number) called exactly what the text
  says, capitals and all, comes before the two-points reading (`Binder.ResolveExactNumber`,
  mirrored in the renamer): two points match in any case, so with points A and B a slider
  named `ab` was the distance AB, and undo of renaming it could not find it in the text.
  A figure given by expressions (`IExpressionOwner`: point by coordinates, line and circle
  by equation) depends on what all of its expressions name, listed in the expressions'
  order - not on what the last one compiled named, appended (X = A.X, Y = A.Y, X edited:
  the point no longer followed A).
- **Expression cycle checks include every binding form**: bare distances, point-function
  arguments (`dist`, `ang`, `area`), numbers and property access. Labels and function graphs
  reject themselves and their descendants, just as coordinate expressions do. A function
  referencing a point on itself otherwise creates a dependency cycle.
- **The expression language calculates as it is written in class** (`Expressions/Parser`):
  a minus in front takes everything up to the next + - * /, powers included, so `-x^2` is
  -(x²) (it was (-x)²: the parabola y = -x^2 opened upward); `^` is right-associative. A
  function of numbers is one of ours (`Expressions/Functions.cs`) or of `System.Math` with
  as many `double` parameters as it is given arguments (`sin`, `max(a, b)`, `atan2(y, x)`),
  each argument an expression (`Binder.ResolveMethod` by name and count, ours first: they
  stand in where `System.Math` does what a drawing doesn't want - `round(2.5)` is 3, not
  the banker's 2; `sign` and `clamp` give "undefined" where `System.Math` throws, for the
  sign of what has no value or bounds the wrong way round); ours that take the names of
  points (`dist`, `ang`, `area`). Everything is a double: a whole number is converted (`sign(x)`, a
  polygon's `NumberOfSides` - unconverted, the first operator applied to it threw), and a
  property that is no number (`A.Name`) is an error said in words
  (`ExpressionTreeBuilder.AsNumber`). Nothing a user can type may throw, caught or not: a
  first-chance exception is an error report on screen. Checked with random texts built
  from the language's tokens and the drawing's names (20000 of them: compile as an
  expression and as a function, evaluate, run the renamer), next to a list of texts with
  known values and of wrong texts that must each give an error with words in it. Left
  as they are: `sqr` is the square root (VB's `Sqr`, for `.dgf` files; of a negative number
  undefined - it took the root of the absolute value), `log` the natural logarithm (`lg`
  is base 10), and `2x` is an error, not a product.
- **Snapping and releasing points** (`Figures/Points/PointSnapping.cs`) swap a point for another
  kind where it is through `Actions.ReplacePoint` (name, label, dependents, lock, a chosen style
  go along). Snap: a free point onto a figure through it - "Snap to line AB" in the grid when
  there is exactly one (`FreePoint` is `IConditionalProperties`; the grid captions buttons
  through `CaptionedMethod`), a submenu in the context menu when more. Release ("Free point"):
  a point on a figure, an intersection point or a midpoint. "Convert to point by coordinates"
  (grid and context menu) makes a free point a `PointByCoordinates` whose X and Y start as
  the numbers it is at, rounded as the grid shows them and without an exponent (the scanner
  reads none), with the keyboard in X; "Free point" is the way back, but not Alt-drag
  (`CanFree` against `CanRelease`: typed coordinates are there on purpose). Alt while
  dragging (read live, on every move) releases a tied point and makes the free one snap to what
  `PointPlacement.FindSnap` finds - what the Point tool would make there (no second midpoints);
  a snap lets go at `Dragger.StickyReach` times the reach. Over another point it sits on top
  of it and the drop *joins* it (`PointSnapping.Join`: its dependents rewired to the target,
  itself removed). No un-join (which dependents would go back?): undo. `CanJoin` refuses a
  target built on the point and one that shares a dependent with it (segment EF: F onto E
  would give a segment EE, and undo's ReplaceDependency would swap its ends). In a join the
  target, with what it is built on, moves before the first figure that comes to depend on
  it (the list is in dependency order, see "The figure list is in dependency order"), a
  part of a composite that went over with its composite is not rewired a second time (a
  regular polygon's sides: it threw), and expressions that named the point name the target
  ([A.X] becomes [C.X]: `ExpressionRenamer` with the target under the point's name, once
  the point has left; undo puts the texts back as they were,
  `IRenamableExpressions.ExpressionTexts`) - so a join into a point without a name (a
  vertex a regular polygon works out) is refused when expressions name the point. A point
  a locus is drawn from - the one that slides, the one traced - is not released, snapped
  or joined at all (`PointSnapping.IsHeldByLocus`; Alt does nothing to it and it has no
  "Free point"; Fix length leaves it alone as well, `LengthConstraint.CanStretch`, and
  Free length only lets its distance go): the locus would trace nothing. Released, the
  sliding point of a gallery locus left exceptions on every move and an undo that did not
  restore the drawing; the `Locus` itself now draws nothing when its second point is not
  on a figure. Swaps happen
  only when Alt first applies and at the drop; the whole drag is one undo transaction.
  Inside `ReplacePoint` the replacement has a temporary name until it takes the point's, on
  undo too, so expressions sit the swap out (`FigureBase.SuppressRenameInExpressions`) and
  the ones that name the point are compiled again at the end
  (`IRenamableExpressions.RebindExpressions`): compiled, they hold the figure that left.
  The point's label is handed over before its other dependents, or it is registered twice.
- **The Figure List** (`UI/FigureExplorer.cs`, left column of `DrawingHost` with a splitter):
  top-level figures but the grid and point labels, by `Title`, with the ribbon icon of the
  tool that makes each (`UI/FigureIcons.cs`: figure type -> tool, walking base types; a new
  figure kind wants an entry). Hidden and `Auxiliary` figures are faded. Its selection *is*
  the drawing's (`Selected` + one `RaiseSelectionChanged`); arrows in the left margin go from
  the keyboard's row to its dependencies. It rebuilds on `ActionManager.CollectionChanged`
  (posted, coalesced), never between `ConstructionStepStarted` and a complete step: temporary
  figures of a tool are never recorded and its real steps sit in its transaction. Undoing a
  deletion puts each figure back at its old index (`RemoveFigureAction.Indices`), so it
  doesn't jump to the end of the list. A hidden figure that is selected is shown ghosted
  (`ShapeBase.IsGhost`, half opacity, shape not hit-testable) until unselected; `Visible`
  stays false, so no hit test, snap or drag sees it (they all ask `Visible`;
  `HitTestShape` refuses hidden figures too). Only figures directly in the drawing: a
  composite's parts (a vector's hidden segment) never ghost, so a hidden composite shows
  nothing. `UpdateVisual` overrides that skip hidden figures test `IsShown` instead, or
  the ghost sits at a stale place. `Actions.ReplacePoint` puts the replacement in the
  old point's place (`MoveBefore`: with what it is built on that came later, such as its
  Number, to keep dependency order; `Figures.Move` doesn't touch the canvas). Several figures selected show as a `FigureSelection`
  in the property grid (common properties + Delete).
- **Cursor philosophy** (`Behavior.GetCursor`): cross = a new *free* point appears here; hand =
  the click picks something already there - a figure the tool needs, an existing point, or a
  place defined by figures (intersection, midpoint); arrow = everything else, including a new
  point sliding along a figure and clicks that do nothing. `winauto cursor` prints the cursor
  showing now (screenshots don't include it).
- **Default point styles are by kind** (`StyleManager.AssignDefaultStyle`, looked up by name):
  the draggable kinds (`FreePoint` yellow, `PointOnFigure` green) at size 10, constructed ones
  at 8. A drawing from a file brings its own styles, and loading adds the defaults it lacks
  by name (see "Styles in files"). Gallery point sizes follow the same standard:
  `dotnet tools/pointsizes.cs -- <folder> [--apply]` lists and raises undersized styles. The
  Rose's 90 control points stay at 5 px on purpose (at 10 they swallow the flower).
- **Point shapes and emoji** (`PointStyle.Shape` / `Character`; `Size` is the shape's or the
  character's, whichever shows: the style keeps both in memory and the file only the one in
  use, so `Character` must be read before `Size`, which declaration order does): every point
  is a `PointMarker` (circle, triangle, square, diamond, pentagon, hexagon; the polygons reach
  past the circle by eye so they look as big), or one character instead, in the embedded
  Twemoji font (`Main/Avalonia/Fonts`, CC-BY, credited in the Emoji tab) with Inter named as
  the fallback: the browser has no system fonts, and a character neither has is not offered
  (`EmojiFont.CanDraw`). The font is 1.5 MB and loaded on first use (`EmojiFont.Open`: a file
  beside the desktop exe, a fetch of `fonts/` in the browser, brotli via web.config and cached
  as immutable: a different font must get a different file name). The
  style's editor has Shape | Emoji tabs (`IPropertyGridTabs`: the tab shown is what the style
  is; picking Shape drops the character, undoably; a row on two tabs, Size, gets an editor on
  each). "What the style is" is what it is under the theme on screen (`ShownCharacter`,
  and `ThemedValue` reads a property without an override from the resolved style): by
  the base values an emoji picked under Dark opened on the Shape tab, could not be taken
  off, and showed the shape's size. The Emoji tab searches `Emoji/Emoji.txt`
  (CLDR names and subgroups of the single-character emoji the font has): regenerate it with
  `dotnet tools/emoji.cs -- <emoji-test.txt> <font> <Emoji.txt>` when the font changes.
- **Touching is decided with a relative tolerance** (`Math.TangencyTolerance`, 1e-9 of the size
  of the numbers): a line and a circle, or two circles, that touch by construction come out a
  hair apart or overlapping at random, and the point there would blink as the figures move. A
  touch gives the exact foot point (no square root of a rounding error to shake what is built
  on it). Ends of segments and the start of a ray get the same allowance (`Math.EndTolerance`),
  ends of arcs an angle one across 0 = 2pi (`IsAngleBetweenAngles`); lines are parallel by the
  sine of their angle, not by a determinant that shrinks with their lengths.
  Never round coordinates or lengths to decide existence - the old 4-digit rounding in
  `GetIntersectionOfCircleAndLine` made every such intersection 1e-5 off and flip at rounding
  boundaries. The P1/P2 order of both intersections is part of the file format (`Algorithm`
  ...1/...2): circle and line, P1 first along the line; two circles, P1 to the right of
  center1 -> center2.
- **Any figure is draggable**: dragging a dependent figure moves its root free points, so gallery
  text can say "drag the circle" even when the points it is built on are hidden. Labels are the
  exception (`AllowMove`): a drag moves the label, not what it measures or names. A point by
  coordinates is no root to move (`PointByCoordinates.AllowMove` is false, and `Dragger`
  leaves such roots out): the drag goes to the other roots, and a figure built on such
  points alone doesn't move, nor does the view. (They used to be "moved", to no effect but
  an undo step that undid nothing - every drag of the Ladder's wall.) A drag whose drop
  throws still ends its transaction (`Dragger.MouseUp`): left open, it took in everything
  done afterwards.
- **The figure list is in dependency order**: a figure comes after what it is built on. The
  file is written in the list's order and read back dependencies first, so a list out of
  order comes back in another order, and `Drawing.Recalculate` goes down the list. What
  breaks it is rewiring: a join (the target may be later than what gets built on it) and a
  tool that records a figure while its last point is still the one following the cursor
  (the Angle tool added its arc with the preview; both figures are made at the end now).
  `Actions.MoveBefore` is the repair, and knows parts (`RootFigureList.FindTopLevel`).
- **Tool letters are the only plain keys**: `Behavior.KeyDown` (what a tool without a key
  handler of its own gets: Point, Coordinates, Slider, Text) used to toggle "Label new
  points" on A and "Snap to grid" on G, from before those were the letters of the Arc tool
  and the grid: G with the Point tool on showed the grid and, silently, made every new
  point snap to it.
- **Numbers are shown with `Settings.DisplayDecimals`** (2) by every editor of the property grid
  and by labels by default; values keep their digits and a typed number is taken as typed.
  Rounding for display goes through `Math.Round(value, digits)` of the library, which adds
  0.0: -0.001 rounds to a negative zero, and that is written "-0" (a point at (3, -0)).
  A half rounds up, as at school (3.125 is 3.13): through `decimal`, which takes 2.675 as
  written where the double is a hair under, and away from zero; `System.Math.Round` goes
  to the even digit, and an area of 3.125 said 3.12. A point's coordinates are written
  (3, 4), with a comma.
- **Several figures in the grid** (`FigureSelection`): a row whose values differ has no
  value (`CompositeValueProvider` answers null), and every editor has to show that as
  nothing - an empty box, no item selected (`SelectorValueEditor.ShowSelected`), a check
  box in its third state (a click then ticks them all) - and not as a value none of them
  has (0, the first item, unticked). The enum editor threw on the null: select all in a
  drawing with two kinds of measurement units. Up/down buttons do nothing without a value
  to step from.
- **Property grid layout is declarative** (`PropertyGrid/`): `[PropertyGridGroup]` boxes rows
  and their buttons together, `[PropertyGridDestructive]` puts a button last under a divider,
  `[PropertyGridIcon]` puts a drawn icon (`PropertyGridIcons`) in front of a caption (every verb
  button has one - give a new verb one too), `[PropertyGridPreferredEditor("UpDown")]` picks the
  editor, `[Domain(min, max)]` on a double gives a `SliderEditor`. String boxes are one line
  (Enter is the panel's: Plot, Add point) unless the property says `[PropertyGridMultiline]`
  (label text). Rows that are editable only
  sometimes: `[PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]` on the
  property and `IConditionalProperties` on the figure (`CanEdit`, `Caption`); the same interface
  vetoes buttons by method name. The grid refreshes a row's value on `PropertyChanged` with
  its name, and rebuilds itself (posted) on a `PropertyChanged` with a null name: for a set
  after which the rows change shape - what is editable, a caption, which buttons (a
  translated point's Free direction). A property setter that changes the figure list (the length
  panel's Show, the point's Free toggles) does so directly: the grid records the property set as
  the undo step and the undo library refuses an action recorded from inside another. A tool's
  panel holds settings, not drawing state, so it says `[PropertyGridNoUndo]` (`ToolPanel` does,
  for all that derive from it; a panel that doesn't derive says it itself): otherwise typing in
  it is an undo step that undoes nothing visible. `[PropertyGridFocus]`
  (the editor takes the keyboard whenever it appears) is for tool panels only; a figure's
  row that should take it just once, when the figure is created (a new label's text), is
  named by `RaiseDisplayProperties(figure, focusProperty:)`, otherwise selecting the figure
  anywhere (canvas, Figure List) steals the keys. A tool panel with a focused row appears
  only at the step that asks for it (Rotate, Dilate, Translate: `PropertyBag` is null
  otherwise - test `Transaction != null` too, since "construction complete" re-reads it
  before the found figures are cleared); one shown all along holding a setting (Line at
  Angle) doesn't take focus.
- **The side panel** (`DrawingHost.CreatePropertyGrid`) is one rounded surface: a fixed
  header row (the grid's title, put there through `PropertyGrid.HeaderHost`, and the ×)
  over a scroll viewer with the rows. Nested grids (complex types, method parameters) have
  no `HeaderHost` and keep their title as the first row. A page starts at its top
  (the scroll offset is reset whenever the grid is given something to show), and a list in it that brings
  its selected item into view (the styles, the emoji) scrolls only itself: the request
  is stopped on its way up to the panel's scroll viewer (`CreatePropertyGrid`), or in a
  small window the panel opened scrolled down to the figure's style, its first rows cut
  off.
- **A figure inherits the rows and buttons of its base class**, and has to take away the
  ones that are not its own: a button through `IConditionalProperties.CanEdit` (by method
  name), a row by overriding the property with `[PropertyGridVisible(false)]`, or by making
  it read-only (`ConditionalPropertyValue`). Each of these was a bug: Convert to ray /
  segment on a parallel, a perpendicular or a bisector (built on a line and a point: it
  threw; `LineTwoPoints.IsThroughTwoPoints`), Convert to segment / sector and Clockwise on
  an angle's arc (`AngleArc`), the Text box of a measurement (`Measurement`, whose text is
  worked out), the style buttons of a Number (no style), a vector's Direction and an
  angle's Arcs when nothing can change them, "Delete this style" on a default style. A
  row or button that does nothing is an undo step that undoes nothing. A row whose setter
  only stores (a point on a figure's `Parameter`, which a locus samples through) gets a
  property of its own for the grid that also moves things (`ParameterDisplay`).
- **A Convert button is not offered when the new kind lacks what something takes from the
  figure**: a segment's length (`IFigureExtensions.IsUsedForLength`: a distance measurement
  of it, a circle with it for a radius, a translation or dilation by it) for Convert to
  line / ray, a sector's or circular segment's area (`IsUsedForArea`) for Convert to arc.
  `ReplaceWithNew` hands every dependent over to the new figure whatever it is: a distance
  measurement of a segment that became a ray threw on every redraw (it now doesn't exist
  when it has nothing to measure). Convert to polyline deletes the polygon, and an area
  measurement of it with it.
- **Closing the side panel** (its ×, or a press on empty chrome: the ribbon's empty
  strip, the toolbar beside its buttons, the Figure List below its rows) goes through
  `DrawingHost.CloseSidePanel`. A tool's own panel (same type as the tool's `PropertyBag`)
  can be the only way on with that tool, so closing it puts the tool down (Drag, as
  Escape); anything else is only hidden. "Empty" is decided in `MainView.IsEmptyChrome`:
  from the hit visual up to the ribbon/toolbar/list through plain layout types only (exact
  types: a toolbar button is a `Border` subclass).
- **The button that accepts a panel says OK**, unless a verb says more (Plot, Add point,
  Close figure, Create tool), and **no panel has a Cancel**: the × of the side panel is
  that (`DrawingHost.CloseSidePanel` puts the tool down when the panel is the tool's own,
  which goes by `PropertyBag` - so a tool with a panel per step, like Define figure, must
  answer with the panel of the step at hand).
- **Enter in a tool panel's box is the panel's button** (OK, Plot, Add point). Each panel
  wires that itself: `[PropertyGridEvent("KeyDown", ...)]` on the property and a handler
  that calls the method. Nothing does it for a new panel, which is how Rotate and Dilate
  went without. The editor commits the text first (its own handler is on the text box, the
  panel's on the editor around it). A number box that says "Type a number." has kept the old
  value, so the handler looks at `StringEditor.ErrorText` before it goes on. Left alone on
  purpose: the length panel, where Enter applies the length and OK only closes. A press
  on the button of the tool that is on already (opening a tab picks its tool) takes the
  keyboard from the panel's box; `BehaviorToolButton.Click` shows the panel again, which
  gives it back. `winauto keys` swallows parentheses: test with `x*x`, not `sin(x)`.
- **Typed text that is wrong is said under its box**, never in the status bar
  (`StringEditor.ErrorText`: a pink plate attached to the box, widening the row up to the
  480 px a box may take). Only on commit - Enter or leaving the box - never while typing, and
  it goes when the text is edited. Two sources: an editor's own `Validate` (`ExpressionEditor`,
  `FunctionEditor` on `[PropertyGridPreferredEditor("Function")]`, `DoubleEditor`), and a tool
  panel's command: panels derive from `ToolPanel` and check each row with `Compile`/`Evaluate`
  (`ReportError` for anything else), which the grid routes to that row's editor. An empty box
  compiles to nothing without an error: `CompileResult.GetErrorText` supplies one. A whole
  number (`IntEditor`) is checked against the property's `[Domain]` ("Type a whole number
  from 3 to 500."). The two number boxes without an error plate put the box right on commit
  instead: `SliderEditor` takes a number beyond its range to the nearest end, `UpDownEditor`
  shows the value again when the text is no number. No editor takes "Infinity" or 1e999.
- **A commit sets the typed text once.** Enter commits, and leaving the box commits again:
  the editors remember the text that has gone into the property (`appliedText` in
  `StringEditor` and `UpDownEditor`) rather than compare it with the value read back, which
  need not equal it (a vector's direction and a segment's length are worked out from
  points: 27.57 comes back as 27.570000000000004) - each further set was an undo step that
  undid nothing. And the text an editor put in the box itself (`shownText`, kept so that
  its own late TextChanged is not taken for typing) is forgotten at the user's first edit
  in the number boxes too (`UpDownEditor`, `SliderEditor`, as in `StringEditor`): a box
  that said 12, Backspace (the length is 1), the 2 typed back - the 12 was taken for the
  editor's own text and dropped, and the segment stayed 1 long under a box saying 12.
- **Setting a segment's length stretches it once** (not a constraint); **Fix length** makes the
  constraint: the end becomes a `TranslatedPoint` from the pivot with an auxiliary `Number` at
  the current distance and a free direction (`IFixableLength`, `Figures/Lines/LengthConstraint.cs`;
  segments, vectors, regular polygons, circles). An end on a line that runs through the pivot
  takes the line as its direction instead (the Number signed along it, so the end is fully
  determined) and Free makes it a point on the line again; an end on any other figure can't be
  stretched or fixed at all. Right after such a figure is made the side
  panel shows a `LengthPanel` (`FigureCreator.ShowCreatedFigure`); the tools themselves have no
  length box. A regular polygon's panel starts with the polygon's own `NumberOfSides` row (the
  panel is an `ICustomPropertyProvider` and forwards the figure's `PropertyChanged` while the
  grid shows it, so the side and the title follow a change of the count). The swap is `Actions.ReplacePoint`, which hands over the point's name label
  (`ReplaceFigureAction` skips it on purpose). A creator's undo transaction spans one
  construction, opened at the first click (`FigureCreator.EnsureTransaction`), so an edit made
  in that panel between constructions is an undo step of its own.
- **Numbers are figures** (`Figures/Values/Number.cs`): roots with no shape, named n1, n2...,
  usable in expressions by name, a length and an angle at once. No tool makes a bare one (the
  Slider is a Number with a handle); a typed value in the Translate tool becomes one. `Auxiliary` (any figure) marks one
  created on demand for another: it is removed with its last dependent and comes back on undo.
  Every deletion goes through `RemoveFigureAction`, one figure per action inside a transaction,
  so undo restores labels and a polygon loses a vertex rather than dying.
- **A label is a length and an angle only when it says a number** (`Label.IsNumber`: its
  text is one, usually the value of an expression, "[AB * 2]" - what the GeoGebra reader
  makes for a computed value). Every `Label` has `ILengthProvider` and `IAngleProvider`,
  so wherever a tool or a tied value takes "a figure with a length" or "with an angle" it
  asks `Label.GivesNumber` too (By Radius, Distance, Rotate, Dilate, Translate, the
  `Accepts` of the tied values, `FigureCreator.LookForExpectedDependencyUnderCursor`): a
  caption, taken for 0, tied a circle's radius or a rotation to a piece of text with one
  stray click. An expression without a value shows "undefined" (`LabelBase.UndefinedText`),
  not "NaN" (nor "Infinity"), and such a label's `Value` is still not a number. A label
  that is one expression and nothing else gives that expression's value, not the text it
  shows (`LabelBase.ExactValue`): `[AB / 3]` as a radius was 0.33, and changed with the
  label's Decimals. A circle whose radius comes out negative or undefined (a label, a
  Number) doesn't exist (`CircleByRadius.UpdateExistence`).
- **Undo** (GuiLabs.Undo; `Actions/`). What must hold: every change to what a file saves is
  one undo step, undo puts back exactly what was there, redo what was made. The traps, each
  of which was a bug (2026-09-30):
  - *A grid button calls its method directly* (`MethodCallerButton`): the method records
    itself (`Actions.*`, a `CallMethodAction`, a transaction when it is several actions -
    Create new style), or it is a hole.
  - *`SetPropertyAction`* keeps an old value per object of a multiple selection (they may
    have differed). A value whose getter resolves - the theme's paper for a drawing without
    one, the base value of a style without an override - is an `IRestorableValue`
    (`ThemedValue`, `Drawing.PaperValue`): undo of the first set takes the override or the
    paper away again instead of writing the resolved value.
  - *Merging* (a run of sets as one step) is asked for by the caller (`coalesce:`): editors
    of typed or dragged values do (`LabeledValueEditor.CoalescesEdits`), a check box or a
    choice doesn't (merged, hide-then-show was a step that undid nothing). Same object,
    same property, same theme only; and never while there is something to redo: the
    library merges without ending the redo chain, so `SetPropertyAction` and `MoveAction`
    refuse then. And only within one run of edits (`SetPropertyAction.Run`, a token of the
    editor): Enter, leaving the box and the grid showing the object anew each start
    another (`EndEditRun`), or a length typed now joined the one typed five minutes ago.
    Undo of a set on several objects restores them last to first.
  - *`MoveAction`* moves by an offset once; undo and redo restore places
    (`IRestorablePlace`: coordinates, a label's pixel offset or pin offset, a translated
    point's distance and direction). Moving back by the offset left a point on a circle, a
    slider knob behind its stop and a label after a zoom somewhere else. The coordinate
    system has no place, only offsets. Keyboard pans use one list of movables so that a
    run of arrow keys merges like a drag.
  - *A setter that adds or removes a figure* (a point's label, a line's name, the length
    panel's measurement) brings the same object back, never a new one: the history may
    hold a drag of it. `PointBase` keeps its label (`keptLabel`), also across undo and redo
    of the point itself; a point is auto-labeled only the first time it enters a drawing.
    And in the same place of the figure list (`RootFigureList.Retire`, then `Return` or
    `ReturnBefore`): the list's order is the file's.
  - *A check box that stands for more than yes and no* (a translated point's Free
    distance: free, a Number, or tied to a vector) is an `IRestorableValue` too
    (`TranslatedPoint.FreedomValue`): undo puts the source back, where unticking the box
    would give it a Number. A label's pin too (`Label.PinValue`): pinned, the label goes
    where the screen carries it, so undo of the pin puts it back where it was in the
    plane, not just the pin (pin, zoom, undo left it somewhere else).
  - *Replacing a figure by another kind* (`Actions.ReplaceWithNew`: Convert to line, ray,
    segment, arc, sector; Reverse) is a transaction that is not delayed, so that each
    step sees the drawing as the one before left it: the new figure goes to the old one's
    place in the list (`MoveBefore`), takes over its name label, stays hidden or locked
    if the old one was, and is named AB by its place in the list (`SettleDefaultNames`).
    While both are in the drawing they trade the names AB and AB2, so expressions that
    name the figure (`[AB.Length]`) sit the replacement out and are compiled again at the
    end, undo included, as in `ReplacePoint` (`SuppressRenameInExpressions`, then
    `RebindExpressions`): following the renames, the text ended as `AB2.Length`, the name
    of the figure that had just left.
    Convert to polyline keeps the polygon's vertices (ABCA).
  - *Rows of a nested grid* (`ComplexTypeEditor`: a line's Equation) get the action
    manager from their parent editor, which passes it on when it is set: an editor that
    starts expanded makes its rows before that, and m and b were edited past undo.
  - *An undo or redo that throws* leaves the library's "currently running" mark on, and
    it then refuses every action: `DrawingControl.ReleaseStuckAction` takes it off.
  - *A click is not a drag* until the cursor has gone `Dragger.DragThreshold` pixels:
    a press with a wobble moved the point or the view by a pixel, selected nothing and
    left an undo step that seemed to undo nothing.
  - *While a transaction is open* (`Drawing.IsRecordingTransaction`: a construction or a
    point drag under way) anything recorded joins it and is rolled back with it. So Redo,
    Delete and Paste do nothing then, a keyboard pan and the grid toggle are not recorded,
    and a click in the Figure List abandons the construction first.
  - *Loading is not undoable*: `DrawingControl.ForgetLoading` clears the history in a
    `finally`, also when the file threw half way. A construction left half way goes with
    its drawing (`DrawingControl.Drawing` resets `ConstructionInProgress`): still "in
    progress" for the next drawing, Undo only restarted the tool and Redo did nothing.
  - *Paste* (`Actions.Paste`) is one step: the copies, then a move of their roots
    `Drawing.PasteStep` pixels down and right per paste of the same copy, and they are
    selected. On top of the originals and not selected, a paste looked like nothing, and
    each try left another copy under them. Copy with nothing selected leaves the clipboard
    as it was. The copies are read under new names where the old ones are taken, so
    whatever a figure's `ReadXml` finds by name must look among its own `Dependency`
    elements, not the names its dependencies have now (a translated point's sources: the
    copy of a fixed segment's end came back free, on its pivot). Expressions of the copies
    are compiled as they are read, when the names they say are the originals'; once the
    copies are in, their texts are rewritten to the copies' names and compiled again
    (`PasteAction.RebindExpressions`, an `ExpressionRenamer` that looks among the copies
    first): the copy of a label `[AB]` measured the original segment. The clipboard
    carries the styles of their own the copies name (`DrawingSerializer.WriteFiguresWithStyles`):
    pasted into another drawing, a figure looked its style up there by name, and took the
    default look, or another drawing's style "1" (a text style on a point). The paste brings
    a style the drawing lacks, under a free name where the name is taken by another look
    (`StyleManager.FreeName`), and takes it away again on undo. Text that is no figures
    (plain text, a page) is said so in the status, not read as XML (that threw).
  - *Parts of a composite* that are among a figure's dependents (a regular polygon's
    vertices) are not removed or put back by `RemoveFigureAction`: they go with their
    composite. Parts removed by reducing the side count are retained for reuse: redo may
    restore a later construction that still holds those exact vertex or side objects.
  - *Length edits* restore the endpoint's place and any Number holding its distance
    (`LengthPropertyValue`, also used by the length panel), not just the previous length:
    zero loses the direction and a negative length reverses it. A point on a figure restores
    its exact parameter, not a new projection of its coordinates.
  To check a change: a temporary `DispatcherTimer` in `MainView` that appends the undo and
  redo counts and a hash of `Drawing.SaveAsText()` to a file whenever they change, driven
  with `winauto`. A hash that changes without a new step is a hole; one that doesn't come
  back on Ctrl+Z is a wrong undo. What found the most (2026-09-30, second pass) was a
  sweep in the same temporary code: for each figure of a drawing, show it in the real
  grid, and for each row set another value through the row's own value provider and
  action manager, for each button invoke it; after each, compare the saved text before,
  after, after undo, after redo, load the saved text into a second drawing and save that
  (a round trip), run `Figures.CheckConsistency`, and log every first-chance exception
  with the row at hand. Run over the gallery and over one drawing made with every tool.
  The third pass added simulated input to the same checks (pointer and key events raised
  on the canvas, `PointerPressedEventArgs` and friends built by hand, so the tools' real
  handlers run): every tool walked at random - clicks on paper, on figures of the kind it
  expects, drags, double clicks, Enter, its panels filled in and their buttons pressed,
  then given up by Escape, a right click, another tool; every figure dragged, with Shift
  and with Alt onto every kind of figure; deleted; copied and pasted; the grid's real
  editors driven (text set and Enter raised on the box, sliders, check boxes); and random
  sequences of all of these with undo and redo in between, then everything undone and
  redone, each state compared with what it was when first reached. Seeded, so a failure
  replays. The temporary code need not touch the repo: a partial `MainView` outside it,
  compiled in with `-p:CustomBeforeMicrosoftCommonTargets=<a .targets file adding the
  Compile items for the LiveGeometry project>` and started by a `[ModuleInitializer]`. What
  the checks can't see: anything that depends on the view (a point on an angle's mark sat
  elsewhere in the round trip's canvas) - which is how that one was found.
- **Sliders** (`Figures/Values/Slider.cs`) are one `CompositeFigure` whose parts are library
  figures: a `FreePoint` anchor, a `TranslatedPoint` knob kept on the horizontal through it
  (free distance, direction a Number saying 0, clamped at the anchor), the `Segment` track, a
  caption label "a = 2.00" a fixed few pixels above the anchor. The drawing and the file see
  one figure (`<Slider X Y Value>`), and no part is ever handed out: `HitTest` answers with the
  slider, so a tool can't come to depend on the knob. The value is the track's length:
  `INumber` (expressions say `a`), `ILengthProvider` (By Radius, Dilate, Translate),
  `IAngleProvider` in degrees (Rotate). Dragging goes by parts (`IMovableParts`, which the
  Dragger asks): the knob changes the value, the anchor and anything else move the whole. The
  parts carry names no expression can say ("slider knob"): `Figures[name]` looks *inside*
  composites and not at them, which is also why `Binder.ResolveFigure` matches top-level names
  exactly before its case-insensitive pass (a slider `a` next to a point `A`). Default names
  are lowercase letters, skipping e, x, y and the ones that read as digits. The slider's
  style is its track's, by default `SliderTrack`: a bar 6 wide in the theme's `SliderTrack`
  gray, narrower than the points at its ends so that they cover them. `Slider.OnAddingToCanvas`
  assigns it before the parts are added, or the track, a segment, takes the default line first.
  Placing one (anchor click, knob follows the cursor, second click or the release of a drag)
  is `PendingSlider`, which the Slider tool and By Radius share: By Radius starts one when its
  first click would make a *free* point, adds it inside the construction's transaction (undo
  takes the circle and its slider together) and goes on to the center. That is decided in
  `Click`, so typed coordinates, which come in through `AddDependency`, still make a point;
  two new free points for a radius are made with the Point tool first. A `FigureCreator`
  that finishes something on mouse up must not call the base `MouseUp` afterwards: with a
  point following the cursor the base takes the release for the next click.
- **Tied values** (`Figures/Values/TiedValues.cs`): a figure's number that is either *typed* -
  an auxiliary `Number` it depends on, shared by every point of a transformed figure - or
  *tied* to a figure it depends on and follows: the angle of a `LineAtAngle` or a
  `RotatedPoint`, the factor of a `DilatedPoint`, the distance and direction of a
  `TranslatedPoint` (which may also be free, see below). `ITiedValues` names them; they are
  rows of the figure's grid, editable while typed, "Angle = ang1" read-only while tied, with
  a "Type the angle" button (`Untie` + name) for the way back, all decided by the figure's
  `IConditionalProperties`. Tying is `TiedValues.Tie`: one undo step that adds a new
  source, moves it before its owners in the list with what it is built on (dependency
  order), swaps the dependency on every sibling (the vertices rotated about the same center
  by the same angle, or translated by the same sources) and drops a Number nothing uses any
  more; a source built on an owner is refused. Right after such a figure is made,
  `FigureCreator.ShowCreatedFigure` shows a `TiedValuesPanel` (the figure's own rows, the
  Untie buttons, OK; titled after the last figure made) and remembers the figure
  (`CreatedFigure`): a click on a figure a value accepts ties it there (a vector gives a
  translation both, or only the values typed at that moment - "Type the direction", then a
  click on another vector, leaves the distance with the first; anything else goes to the
  first value that takes it) instead of starting
  the next construction, until anything else takes the side panel, the next construction
  starts or the tool stops; then the typed values become the tool's defaults
  (`TakeDefaultsFrom`). The pre-create panels stay (Line at Angle's angle, Rotate's and
  Dilate's dialogs, Translate's steps): OK in the post panel goes back to them.
- **TranslatedPoint** (`Figures/Points/TranslatedPoint.cs`): distance and direction are each
  *tied* to a figure (vector, length/angle provider, `ILine` as it points, a Number) or *free*
  (the parameter dragging changes); both are tied values (above), `SetSource` does the swap
  since a vector serves both at one index. Roles are stored by index into the dependency list, never
  inferred from types (a `Label` is both a length and an angle provider). A point with a free
  quantity takes the green `PointOnFigure` style. The Translate tool is stepwise (source,
  distance, direction, placement); its panel is cleared in `Stopping` because that event is
  raised before `Started` resets the state, and `DrawingControl` re-shows a behavior's
  `PropertyBag` on "construction complete" as well.
- **Ellipses**: center, end of the long axis, end of the short axis, the last by its distance
  from the long axis, so a point on the short axis is on the ellipse. The tool puts a free third
  click on a hidden perpendicular through the center (`EllipseCreator.FindPointPlacement`), so
  scaling the long axis scales the short one with it.
- **Arcs, sectors and segments of circles and ellipses** (`Figures/Circles/ArcBase.cs`) get
  their areas from the sweep t in the angle that parametrizes the ellipse (x = a cos t,
  y = b sin t; the central angle on a circle): sector a·b·t/2, segment a·b·(t - sin t)/2;
  the length of an elliptical arc is Simpson's rule over t. They said r²·angle/π (a quarter
  disc of radius 2 had the area 4), and "not a number" for every ellipse - so did an Area
  or Distance measurement on one, and a circle given such an arc for its radius. A point
  put on an arc from beyond its ends goes to the nearer end, the short way round
  (`GetNearestParameterFromPoint`). A point on a function graph exists wherever the
  function has a value: its place is not hit-tested, since a graph is hit by its samples
  and those end at the edges of the window. A point on a locus keeps the parameter of
  the locus's sliding point (on that point's own figure) and is worked out exactly, by
  putting the sliding point there and back as for a sample (`Locus.GetPointFromParameter`);
  it kept a fraction of the drawn curve's length, and how much of a locus is drawn depends
  on the view when the sliding point is on a line - the point slid along the curve with
  every pan and zoom. Not hit-tested either. Numbers like these were checked in the same
  temporary code as undo (see "Undo"): figures made by the `Factory` and their values
  against known ones, and for every figure a point can sit on, a grid of points projected
  onto it - each must land on the figure, stay put when projected again and be its
  nearest place (by angle on an ellipse, by x on a graph). A circular arc doesn't exist
  while its radius is 0 or its end is on its center (its length said "NaN"). A
  self-crossing polygon's area is what its even-odd fill covers (`Math.FilledArea`); the
  shoelace sum took the lobes that go round the other way away from the rest, and a bow
  tie measured 0. An open polyline's length has no closing side. An angle of a full turn
  but for rounding is 0 (`Math.OAngle`: two sides along one ray blinked between 0° and
  360°), and a mark reflected in a line goes clockwise (`AngleArc.UpdateVisual` takes the
  long way only when that way is long).
- **Curves have gaps** (`Curve.Gap`, a point that is not one, in what `GetPoints` gives):
  each stretch between gaps is a figure of its own in the geometry. A function graph has
  one where the function has no value - the graph goes on to the very edge of where it
  has one (`FindEdge`), sqrt(x) starts at 0 - and where it jumps between two samples
  (`IsJump`: halving the step towards the larger change, a climb gets smaller and a jump
  stays), so 1/x, tan x and floor(x) have no vertical lines; values far beyond the window
  are clamped (a coordinate of 1e300 pixels is not drawn, and exp(x²) vanished whole). A
  function that throws has no value there (it was 0). A locus has one wherever the traced
  point doesn't exist. Before, a curve was one line through all its points: straight
  pieces across where there is nothing.
- **Vectors** are an invisible `Segment` plus an `Arrow` polygon sized in pixels, filled with the
  line color. `Vector.OnAddingToCanvas` sets the default `LineStyle` before the base call,
  otherwise the polygon default (pale fill) wins. A vector is an `ILine` (parallel,
  perpendicular, intersection, a point on it all take it) and its hit test asks the segment
  inside after the arrow, since the arrow is a filled polygon with no room around it.
  `Vector.HitTest(Point)` answers whether hidden or not, as a segment's does: a point on a
  vector and an intersection with one exist where that says, and through the composite's
  own test (shown parts only) they all went when the vector was hidden.
- **Names of lines and circles** (`Figures/Controls/FigureLabel.cs`): "Show name" on a line,
  ray, segment or circle (`LineBase`/`CircleBase.ShowName`, over `FigureBase.HasNameLabel`)
  adds a `FigureLabel` the way a point's name is a `PointLabel`: a label depending on the
  figure, `Offset` pixels from an anchor on it - a segment's middle, the upper left of a
  circle, and for a line or ray a point of its visible part `EdgeInset` pixels in from the
  window's edge (the end nearer the top for a line, the far end for a ray), so the name stays
  on screen and moves along the line as the view pans. The setter adds and removes the label
  directly (the property set is the undo step), the same label every time, so it comes back
  where it was dragged to; the label links itself back on undo
  (`OnAddingToDrawing`), hides with its figure, follows renames (`FigureBase.Name`), and stays
  out of the Figure List and of Delete like a point label. A name label has no Visible of
  its own in the grid, and Hide in the context menu turns its figure's Show name off:
  hidden by itself, it stayed hidden whatever Show name said, and nothing listed it to find
  it again. (A label from a file hidden that way shows when the name is shown again.)
  Hide and Lock in the context menu take the whole selection, as Delete does. GeoGebra's `<show label>` on those
  figures turns it on, in the figure's color.
- **Segment marks** (`Segment.Decoration`, `Figures/Lines/SegmentDecoration.cs`): one to
  three ticks across the middle, one to three chevrons along it (pointing from the first point
  to the second), or a wave - the school notation for equal and parallel sides. A passive
  visual like the right angle mark (`SegmentDecorationMark`: a Path the segment adds to the
  canvas, in the segment's own stroke as drawn, sized in pixels, not hit-testable), chosen in
  the grid through a row of swatches (`SegmentDecorationEditor`, drawn by the same geometry),
  saved as `Decoration="TwoTicks"`. Only segments: a polygon's sides aren't figures here. The
  wave is ours; GeoGebra's file has ticks and arrows (`decoration type` 1-6, mapped on import),
  its line start/end caps are not read.
- **Right angle marks** (`Figures/Lines/RightAngleMark.cs`) are a passive visual owned by
  `PerpendicularLineBase` (perpendicular line, segment bisector), deliberately not an angle
  figure. Which corner the mark sits in is *stored* (`Corner`), chosen once and never derived
  from the geometry again - deriving it makes the mark flip-flop on rounding; a click with the
  Drag tool moves it to the next corner. Chosen when the line is first worked out, shown or
  not, with its foot on the base line or beyond: chosen when the mark first showed, it
  changed what the file saves when a hidden line (a square's helper) was shown or a point
  dragged, and undo of that did not put it back. It hides when an `AngleArc` sits at the same vertex,
  because a measured angle of exactly 90° draws the same sign itself.
- **An angle is two figures**, `AngleMeasurement` (the number) and `AngleArc` (the mark, 0-3
  arcs), paired by `AngleArc.FindCompanion`. Neither exists while a side has no length (a
  point dragged onto the vertex, `AngleArc.HasSides`): not existing hides a shape and shows it
  again later (`ShapeBase.Exists`), where setting `Shape.Visibility` by hand is forever. The
  mark is an arc in code only: a sign of a fixed size in pixels, so no point goes on it and
  nothing is intersected with it (`PointOnFigure.CanBeOnFigure`,
  `IntersectionPoint.GetAlgorithms`) - a click near the vertex glued the new point to the
  mark, and it moved with every zoom. Its grid says the angle in degrees, like the number. `DGFReader.ReadMeasureAngle` creates the arc from
  VB6's DrawStyle / AuxInfo(2) - not tested, there is no sample .dgf with an angle in the repo.
- **Dashes**: `LineStyle.Dash` is put on in `LineStyle.OnApplied`, not through a setter, because
  `StrokeDashArray` counts in stroke widths and a selected figure is thicker. Anything that
  applies a style to a shape by hand (sample glyphs) must call `OnApplied` too. Any enum
  property of a style or figure round-trips by name through the generic `EnumSerializer`.
- **Color/brush picking** (`DynamicGeometry/Controls/ColorPicker/`) is layered so parts can be
  swapped: `ColorPalette` -> `ColorPage` (swatches, spectrum) -> `ColorPickerView` ->
  `BrushPickerView` (solid | gradient); in the property grid through `ExpandingPickerEditor`.
  A gradient is its stops and an angle; the brush's two points are the line through the
  middle of the box at that angle, long enough for the end colors to be reached in the
  corners (the CSS rule: (0,0)-(1,1) at 45°, edge to edge at 0° and 90°, outside the box in
  between), so that the stops reach every part of the shape. Gradients made before
  2026-09-29 have a line of length 1, which leaves the corners of a diagonal one flat; they
  load as they are and take the new line when edited.
  Whoever hosts a `SegmentSwitcher` sets its `Surface` to the background it sits on so the
  selected tab blends into it.
- **The paper** is `Drawing.Background`, edited through the "Background" button on the
  Coordinates tab (`DrawingHost.ToggleDrawingProperties` puts the drawing itself in the property
  grid). A gallery tile takes a drawing's paper as its plate and turns its caption white on a
  dark one.
- **Save writes the drawing's own file again, without a dialog** (`MainView.SaveDrawing`),
  when it has one: `OwnFile`, the .lgf it was read from (the Open dialog, the command line)
  or last saved to. A new drawing, a drawing of the gallery and one read from a GeoGebra or
  DG file have none and get Save as (`SaveDrawingAs`, also the first item of the Export
  menu), whose file is the drawing's from then on. A file that failed to load, or loaded
  only in part, is not kept (`Drawing.Name` is set only by a load that went through, see
  "A file is read as far as it goes"), or Save would write the ruins over it. A
  construction under way is put away before saving (its point following the cursor and
  its preview are figures while it lasts, and went into the file). In the browser the first Save of an opened file makes the browser ask for
  permission to write; a browser without the File System Access API gives files to read
  only, and Save falls back to Save as there (`IsReadOnlyFile`; not tried in a real
  Firefox). A file the desktop can't write (read-only, open in another program) is said so
  in the hint and Save goes on to Save as (`MainView.TryWriteFile`, Export's too): it was
  an error report, as for a bug. There is no unsaved-changes prompt and no mark of a
  changed drawing.
- **Export** (`MainView.Export.cs`, the button after Save; its menu is a `MenuFlyout`): Save
  as .lgf (see above), then Save
  as .png, Save as .svg, Copy image - the canvas as it is on screen at that moment, same view,
  same size, without the side panel and the status bar (siblings over the canvas, not in it;
  a selection's highlight is in it). `ViewportImage` makes both: the PNG has the pixels of the
  screen (layout size × the screen's scaling; the clipboard gets the same picture), the SVG is
  in layout units and is what Skia writes (`SKSvgCanvas`) when Avalonia draws the canvas onto
  it (`DrawingContextHelper.RenderAsync`), so anything that draws on screen is in it with no
  code per figure. Skia writes text as `<text>` in a font, and a viewer without that font -
  every viewer, for Inter and the emoji - draws its own: `SvgTextOutlines` replaces each
  `<text>` by the outlines of its glyphs and a color emoji by its COLR layers in their CPAL
  colors (version 0 tables, which Twemoji has), so the file needs no font. The glyphs are
  looked up by character (cmap), not taken from the shaped run: a ligature or contextual
  alternate would come out as its plain letters. A text whose font isn't found stays
  `<text>`. Skia's `font-weight` is one step too light from 500 on ("600" is a bold 700):
  `ReadWeight` undoes that, look at it again when SkiaSharp changes. Fonts are found by the
  family name Skia wrote: the emoji font, Avalonia's `fonts:Inter` collection, then the
  installed ones; a new embedded font wants an entry in `FontFaces.OpenEmbedded`. To check
  an SVG, open it in headless Edge (`webauto start file:///...svg`) next to the PNG.
- **No menu.** New, Open, Save, Export | Undo, Redo are one row (`LiveGeometry/MainToolbar.cs`) above the
  ribbon, with the tour group (◀ n/N ▶ + title) between them and the Octocat at the right,
  whose tooltip is the build. The first button is the app's mark and folds the ribbon (Ctrl+F1,
  `MainView.UpdateRibbon`): folded by default when a gallery drawing opens on a small screen
  (under 700x500), open otherwise; once pressed, the user's choice holds for the session.
  Everything else is keys. Lost its menu entry and is unreachable for now: Lock. Not on the Selection tab (obscure for the audience; the commands and
  settings are still there in `DrawingHost`): Ortho, Polar, Snap to grid, Snap to point, Snap to
  center. Shift while dragging or clicking still snaps to the grid, and a click near the middle
  of a segment still makes a midpoint. In a narrow window (a phone, under about 410 px)
  the buttons close up (`MainToolbar.SetCompact`: square, no spacing, tighter
  separators, about 330 px), or the last of them, the settings gear, was cut off; what is
  at the right (theme, Octocat) is dropped first when there is no room.
- **The chrome's colors are a theme** (`UI/AppTheme.cs`, not `Theme`: every control has a
  `Theme` property, its ControlTheme, which would shadow the class). One `Color` property per
  role (`Strip`, `HeaderRow`, `Background`, `Text`, `Ink`, `Accent`...), and `AppTheme.Light`
  and `AppTheme.Dark` are the values; each theme is a `ThemeVariant` whose resource dictionary
  of brushes (under the property names) is registered on the application, and the chrome binds
  to those resources - the code side of DynamicResource - through `ThemeBinding`:
  `border.BindTheme(Border.BackgroundProperty, nameof(AppTheme.Background))`. A control whose
  color follows its state binds again to another name (a null name unbinds, `whenNone:` gives
  the plain value); pens made in `Render` come from a styled property bound the same way with
  `AffectsRender`; a gradient of theme colors is rebuilt from `ObserveTheme`. Never copy a
  color out of a theme brush into a plain brush: it stops following. Fluent's own controls
  (text boxes, combos, scroll bars, menus, tooltips) follow the variant on their own; the
  property grid's styles use `DynamicResourceExtension` setters (`PropertyGridTheme`).
  Switching is `AppTheme.Apply(choice)`: a theme's name, or `System` (`RequestedThemeVariant
  = Default`, the OS's or the browser's `prefers-color-scheme`); `AppTheme.CurrentChanged`
  is for what can't bind. A theme is a `[PropertyGridNoUndo]` object with a color picker per
  row: the settings page (the gear after Redo on the toolbar) opens `AppTheme.Current` through "Theme
  colors", every pick repaints the app at once (the setter writes the dictionary, which
  notifies its owner), and "Copy as code" puts the initializer on the clipboard to paste
  back into `AppTheme.cs`. A write to the dictionary makes every resource binding in the
  app (Fluent's templates included) look its resource up again, whatever the key, and
  a drag in the picker is a set per pointer move: the setter changes the property at
  once but posts the dictionary writes (`AppTheme.SetResource`, one flush per idle
  tick), and `ColorsChanged` - the drawings' refresh - fires after the flush, only for a
  color drawings take (`IsDrawingColor`: the Paper group and `Text`). `Drawing.RefreshTheme`
  applies each figure's style once (`IsRefreshingTheme` holds the styles' own notifications
  back while they re-read the theme), and only while the drawing is on screen: a hidden
  gallery tile or the parked drawing catches up through `RefreshThemeIfStale` against
  `AppTheme.Version` when it shows again (the gallery's `IsVisible`, `MainView.ShowEditor`,
  the attach to a canvas; Avalonia's `IsEffectivelyVisibleChanged` is internal). Without
  this, one pick in the browser (interpreted, no AOT) took over a second. A new theme is another instance with its own variant, inheriting
  Light or Dark for Fluent's sake, added to `All`. Tool icons draw their lines in `Ink`
  (`IconBuilder` binds them, and takes a theme color's name where a `Color` was passed);
  the drawn chrome icons (`MainToolbarIcons`, `PropertyGridIcons`) use a sentinel brush that
  `Shape` rebinds to `IconOutline`. A tool icon's fills are theme colors too: the Paper group
  for what a figure would look like (points), and the
  Icons group for what is only a picture (`ShapeIconFill`/`ShapeOutline` of every shape in
  the Shapes icons, `ImageFill` of a transformation's image - the figure transformed is
  filled as a shape is, though outlined in ink like its image -
  `RulerFill`/`RulerOutline`, `AngleFill`/`AngleOutline` with `ScaleMarks` on both,
  `AreaFill`/`AreaHatch`, `PaperIconFill` (a brush), `LineAccent` for the line or curve a
  tool makes out of the figures its icon also shows: `IconBuilder.AccentLine`, 1.5 thick)
  - no literal brush in a `CreateIcon`,
  or the theme can't reach it. A shape's icon is not filled with the Paper group's
  `ShapeFill`, the fill of a new polygon: that one is translucent, made for the paper, and
  a gradient there would fill every new polygon with it. `ShapeIconFill` and
  `ImageFill` are a `Brush`, not a `Color`, so
  the grid gives them the brush editor and they can be gradients; to make another fill one,
  change its type and its two initializers (a property bound to it must take a brush, and
  `ObserveTheme` hands out solid colors only). The sun/moon
  beside the Octocat (both pages) flips between light and dark, and landing on what the system
  says stores `System` again. On Windows the title bar goes dark too
  (`LiveGeometry.Desktop/WindowFrameTheme.cs`: the DWM attribute, and then a non-client
  activation cycle plus a frame-changed `SetWindowPos`, since Windows 10 keeps painting the
  old shade until the next activation otherwise). `winauto shot` doesn't show the frame at
  all (PrintWindow): to check it, `shot ... --screen`.
- **Drawings follow the theme through their styles.** A `FigureStyle` has its values (how it
  looks under Light, the base theme) and may hold *overrides* for another theme: the
  properties that differ, with their values (`Overrides`, `SetOverride`). Whoever draws with a
  style resolves it first (`IFigureStyle.Resolve`: the style itself without an override for
  the theme on screen, else a copy with the override applied); `GetWpfStyle` and the
  `Apply` extension do that, so figures and sample glyphs need nothing. The default styles
  (`StyleManager.AddDefaultStyles`, the grid's own in `CartesianGrid`) take their colors from
  the theme's Paper group (`AppTheme.Paper`, `Ink`, `Line` - the default line, translucent
  black in Light and an opaque light gray in Dark, where a translucent one came out too dim -
  `SliderTrack`, the point fills, `ShapeFill`, `Axis`, `GridMajor`/`GridMinor`) through `BindToTheme`: base from Light, an override from every
  other theme, read again when a theme color is tweaked (`AppTheme.ColorsChanged` ->
  `Drawing.RefreshTheme`) - except where the user has given the property another value
  since (a red fill for the default point style): `FigureStyle.ReadFromTheme` remembers
  what the binding last gave each theme and leaves a value that differs alone. It was
  overwritten whenever the themes were read again (also when the theme switched while the
  drawing was behind the gallery), and then not saved. Chosen colors that read on both papers (a red line, a blue
  outline) stay literal; a gray helper line gets a literal Dark override. A drawing's paper is
  the theme's unless it has one of its own (`Drawing.OwnBackground`, null for the theme's;
  files leave it out; the readers of foreign formats take white as none). A theme switch
  (`AppTheme.CurrentChanged`) re-applies every figure's style and the paper
  (`Drawing.RefreshTheme`, hooked by `DrawingControl` and the gallery's `DrawingThumbnail`,
  which doesn't paint the paper: `Drawing.PaintsPaper`). The canvas's own binding to the
  theme's paper is at `BindingPriority.Style`, below the value the drawing sets: at the same
  priority the binding pushes the theme's paper over a drawing's own on every switch. The
  grid's colors go by the paper it is on (`CartesianGrid.GetColor`): on a solid paper of the
  drawing's own, those of the theme whose paper is nearest in lightness, shifted by as much
  as the paper differs from that theme's. The tool icons draw their points,
  fills and lines from the same Paper group, so they show the figures as the theme would.
  **The property grid edits what is on screen**: under the base theme (Light) a style's
  property itself, under any other theme its override for that theme
  (`ThemedValue.ForCurrentTheme` wraps the value the grid edits; `IThemeOverridable` is
  what a style and a drawing's paper implement), and undo puts the override back, or takes
  it away if there was none. A "Same as in Light" button drops the theme's overrides (shown
  once there are some, on the next opening of the style). The paper works the same:
  `Drawing.Overrides` holds the paper chosen for another theme, "Theme's paper" under Dark
  stores a null override (that theme's paper). Both buttons are undo steps.
  `GalleryTitle` (a Dark override in the splash's blue), `GalleryText` (the chrome's text
  color) and `GalleryLocus` are defaults too. The gallery drawings' own styles got their
  Dark overrides from `dotnet tools/darken.cs -- <folder> [--apply]` (2026-09-29): a stroke,
  text or fill darker than lightness 0.35 is lightened to the same hue at 0.82 - 0.4 ×
  lightness (black lands on the ink), a drawing with a paper of its own is left alone, and a
  style that has a `<Dark>` already is never touched, so it is safe to rerun after adding a
  drawing. The three drawings with a light paper of their own (Castle, Pascal, Rose) carry
  `GalleryTitle`/`GalleryText` copies whose Dark override is the light color, so the caption
  stays dark on their paper. `--check <folder> <out> --dark` renders the pictures under the
  dark theme; a contact sheet of the gallery in each theme is the way to review. Anything
  that reads a style's color itself rather than through `Apply` must resolve it first
  (`Arrow.ApplyStyle`). A show/hide check box takes a text style like a label
  (`ShowHideControl.ApplyStyle`, `Text` by default, `GalleryText` in the gallery) and pins
  Fluent's per-state caption resources to that color, or the hover would repaint the
  caption in Fluent's own. The gallery's tiles keep their one list of pastels: under the dark
  theme a tile deepens its pastel (`GalleryTile.Deepen`: same hue, value 0.24) and derives
  a lighter border from a dark plate; a drawing's own paper on a tile is the paper as the
  theme resolves it. A tile follows `AppTheme.CurrentChanged` on its own.
- **Settings between runs** (`LiveGeometry/SettingsStore.cs`): `Get`/`Set` by key, the desktop
  head keeping them as `key=value` lines in `%LocalAppData%\LiveGeometry\Settings.txt`
  (`FileSettingsStore`), the browser in `localStorage` under `LiveGeometry.<key>`
  (`BrowserSettingsStore`, two imports in `main.js`; `index.html` reads the `Theme` entry
  before the runtime starts, so the splash and the page are already dark). `AppSettings`
  (`[PropertyGridName("Settings")]`, the gear) is the page over it: `Theme` is `System` or a
  theme's name, applied in `App.Initialize` before the first frame. The window placement is
  the `WindowPlacement` line of the same file.
- **Ribbon look**: `ButtonGrid` draws the
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
  "Animation effects" setting, which many machines have off, and a still splash looks stuck.
- **web.config**: `LiveGeometry.Browser/web.config` is hand-written (serves the precompressed
  `.br` files, sets immutable caching on fingerprinted assets, `no-cache` on entry files). The
  wasm SDK drops a project web.config from publish, so the csproj copies it with an explicit
  `AfterTargets="Publish"` target. It sits at the publish root and rewrites into `wwwroot\`.
  There is no IIS locally: verify changes after deploy with
  `curl -s -o /dev/null -D - -H "Accept-Encoding: gzip, br" https://livegeometry.com/_framework/<file>`
  and expect `Content-Encoding: br` and `Cache-Control: public, max-age=31536000, immutable`.
  A malformed web.config takes the whole site down (HTTP 500).

## File format (.lgf)

No compatibility requirements: the gallery drawings are the only files that matter (see
"Backwards compatibility is not a concern"). The upgrade notes below describe what the
loader still does for files from before; none of it needs extending.

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
  (`PointLabel.ClampPosition`); up close, what is kept clear is the box of the letters (from
  the `TextLine` ink metrics), not the line box, so a name can touch the point's outline.
- **`<Drawing IntersectionOrder="Legacy">`**: `Math.GetIntersectionOfCircleAndLine` once swapped
  P1/P2 for a line through the center, and files carried no version then, so old drawings (the
  phone ones) that pick the other intersection opt in to a swap on load
  (`IntersectionPoint.UpgradeLegacyCircleAndLineOrder`). No file in the repo carries the mark
  (`LiveGeometry.Desktop.exe --modernize <folder>` writes into a file what loading it upgrades);
  the code stays for old files from elsewhere.
- **Styles in files** (since 2026-09-29): a figure on the default style of its kind (a free
  point on `FreePoint`, a segment on `Line`) has no `Style` attribute, and gets it on loading
  (`EnsureStyleAssigned`); one on another default names it (`Style="PointOnFigure"` on an
  intersection point drawn green); only a custom style is carried as an element. The default
  names are the constants in `StyleManager` (`FreePoint`, `PointOnFigure`,
  `IntersectionPoint`, `Midpoint`, `DependentPoint` - `DependentPointStyle` in older files,
  `Line`, `SliderTrack`, `Shape`, `OutlinedShape`, `Text`, `Heading`, `Hyperlink`, and the palette ones like
  `RedLine`). Saved are the styles the figures name (`DrawingSerializer.Write` writes the
  figures aside first and collects their `Style` attributes - a figure may name another's
  style, a vector its arrow's) and a default the drawing changed; a default as a new drawing
  has it is left out (`StyleManager.IsUnchangedDefault`). `Name` comes first, and an
  attribute at the value a fresh style has (`IsFilled="true"`, `Dash="Solid"`) is left out; a
  missing one reads as that value. What differs under another theme is a child element
  named after the theme, its attributes (and gradient elements) the properties as in the
  style's own element: `<LineStyle Name="RedLine" Color="#FFD83B3B" StrokeWidth="1.5"><Dark
  Color="#FF00BFFF" /></LineStyle>`. The paper's is the same on `<Viewport>`: `<Dark
  Color="..." />`, `<Dark><Background>gradient</Background></Dark>`, or an empty `<Dark />`
  for that theme's own paper. Loading lays the styles out in a new drawing's order, a file's
  style taking the place of the default of the same name unless it looks the same in Light
  (files from before the themes carry every default they use: those are dropped for the
  default itself, which follows the theme), others after (`StyleManager.AddWithDefaults`).
  A default made to look under a theme as it does in Light ("Same as in Light") and
  otherwise unchanged would be taken for such a copy: it is written with an empty theme
  element (`<PointStyle Name="FreePoint" ...><Dark /></PointStyle>`), and a style read
  with any theme element is kept (`FigureStyle.SaysThemes`).
  A file style under any name that looks like what a default used to be (the numbered
  copies of the phone and CD drawings: the yellow, green and gray 10 px points with a black
  rim, the 18 and 40 pt black Segoe UI text; `StyleManager.LegacyDefaults`) is dropped too,
  and its name resolves to the default (`aliases`). Not the old opaque black line: the
  default line is translucent, and every old drawing would turn gray; those stay file styles
  (`dotnet tools/stylecensus.cs -- <folder>` lists what a folder's files carry, by values).
  The order matters: new lines and shapes take the first line or shape style. Labels leave
  out `DecimalsToShow` at the default. `LiveGeometry.Desktop.exe --rewrite <folder>` loads
  and saves every drawing of a folder, keeping each file's viewport.
- **A figure is asked before it is worked out** while a file is read: whatever a `ReadXml`
  works out (an intersection point does) sees figures that have been read but never
  recalculated nor added. A Bézier curve had no points yet (`Bezier.CurveInfo` works them
  out on demand) and a locus an empty list (`Math.GetPointOnPolylineFromParameter` answers
  "no such point"): both threw for a point on them, and opening the drawing showed an
  error, though the point came out right once the drawing was recalculated. A point on a
  figure no longer works itself out there: it is where the file says until its figure
  comes into the drawing. One that doesn't exist kept the (0, 0) it got from figures that
  were nowhere yet, and wrote it into the next file; in a session it is where its
  parameter says, also off the end of an arc that turned away (`PointOnFigure.Recalculate`),
  so that the same drawing is saved with the same numbers. A figure given by expressions
  asked then compiles what it can: what names a figure doesn't compile yet, and the rest
  must not set its dependencies (`DrawingExpression.Recalculate` keeps the file's until
  all of its expressions compile) - a line by equation `x = P.X` lost P, whenever the
  intersection read before it happened to exist.
- **A file is read as far as it goes.** What the file is comes first
  (`DrawingControl.ParseDrawing`: XML with a `<Drawing>` root; `GeoGebraReader.ReadWorksheet`
  checks for a zip; a `.dgf` must have its `[General]` section): a file that is none, or
  damaged, is said so in the status and leaves the page as it was - it used to switch to
  the editor under the file's name first, so a picture picked by mistake made a gallery
  drawing the user's own and dropped one parked behind the gallery, and any other XML (an
  exported .svg) opened as an empty drawing that Save then wrote over. Within a drawing,
  what can't be read is left out with a line in words (`DrawingDeserializer.ReportError`:
  a kind of figure or style this version doesn't have, a figure built on one the file
  lacks or on itself, a second figure of the same name, a `ReadXml` that threw), and the
  rest comes in; it threw half way, with a bare name or "the given key was not present"
  for a message. Such a drawing keeps the list (`Drawing.LoadErrors`), gets no name and
  so no file to be saved over, and the status says the first line and how many more.
  `IniFile` skips the lines of a `.dgf` it can't read.
- **Label text** is one attribute: a line break is the two characters `\n` and a
  backslash is two backslashes (a typed `C:\notes` came back as two lines); characters an
  XML file can't hold are left out (one pasted in made Save throw and write nothing).
- **A show/hide box is read as a box** (`ShowHideControl.ReadXml`): each figure says in
  the file whether it is hidden. Applying the box to its figures on load hid again one
  that had been shown by hand since the box was last clicked. (The `.dgf` reader does
  apply it: those files have only the buttons.)
- **Saved files declare `encoding="utf-8"`**; files from older builds say `utf-16`, which our
  own loader tolerates but `XDocument.Load` does not.
- **`TranslatedPoint`** without `DistanceSource`/`DirectionSource`/`FreeDistance`/`FreeDirection`
  is the old format (typed values as attributes, roles by position, direction in radians);
  `UpgradeLegacyValues` gives the typed values auxiliary Numbers. A `RotatedPoint` or
  `DilatedPoint` always has its angle or factor as its third dependency (since 2026-09-29;
  the `Angle`/`Factor` attribute before that is not read).
- **Pinned labels** (`Label.Pin`, `Figures/Controls/LabelPin.cs`): `Pin="TopRight" OffsetX
  OffsetY` is the pixel distance from a canvas corner to the *same* corner of the label, so it
  keeps that edge while the plane zooms and pans under it. `Coordinates` always tell where it
  is in the plane right now, so hit testing, dragging and the property grid need nothing
  special. `WrapWidth` (px) and `Backdrop` (a plate of the paper's color) exist for captions.
  Pinned labels draw above figures, zoom to fit ignores them, thumbnails hide them.
- **Parts of a figure** (`IFigureParts`, `Figures/Shapes/DependentPolygonBase.cs`): the
  vertices, sides and inside a regular polygon works out are figures other figures can be
  built on, but not figures of the drawing: unnamed (`PolygonVertex`, `PolygonSide`,
  `InteriorPolygon` skip the naming of `OnAddingToDrawing`), written as
  `<Dependency Name="RegularPolygon1" Part="Vertex3" />` (`Vertex2`... - `Vertex1` is the
  point the polygon is built on - `Side1`..., `Interior`). So a regular polygon makes its
  parts in `ReadXml`, before it is on a canvas (`IsOnCanvas`), and says `Sides="7"` when it
  is not a pentagon. The number of sides doesn't go below what figures built on its vertices
  and sides need: the parts would take them along, from inside a setter, past undo. A
  hidden polygon hides the parts it makes (`Recreate`): read from a file, it made them
  after it was told it was hidden, and its vertices and sides came back to be clicked.
  The parts are listed with what they are built on (the polygon, its first vertex) only
  while the polygon is in the drawing (`RegisterPart`): registered when read, before the
  polygon was added, they failed the consistency check that a point by coordinates runs
  when it is added, and such a file did not load. When the point a polygon is built on is
  replaced (Fix length, Convert to point by coordinates), the polygon takes its parts
  along (`ReplaceDependency` moves their registration; `SubstituteWith` skips them).
  Hit testing hands the parts to the tools, and whatever takes a figure from a click has
  to cope with one that has no name and is not in the drawing's list
  (`RootFigureList.FindTopLevel` gives the figure it belongs to): a tool must not take a
  part of its own preview (`FigureCreator.LookForExpectedDependencyUnderCursor` leaves out
  what is built on the point following the cursor - a double click with the Regular
  polygon tool made a polygon on a vertex of its preview, and the saved file did not
  load); the Transform tools take the side or the inside for the polygon
  (`Transformer.FindTransformSource`: a copy of a side was a segment without a name);
  `TiedValues.Tie` must not add it to the drawing as a figure; a caption calls it by its
  polygon's name (`TiedValues.SourceName`); a translated point says such a source by its
  place among the dependencies (`DistanceSourceIndex="1"`, where a named one is
  `DistanceSource="AB"`).
- **`AngleMeasurement`** says `Radians="true"` when it shows radians (it was not saved).
- **Scenes** (`<Scene Left Top Right Bottom />` under `<Drawing>`, 1 or 2) are opt-in suggested
  views for drawings whose content has no useful bounds (endless ground: Castle, The Falling
  Ladder; a locus: Spiral). Where they exist, fit and tile show the scene nearest in shape to
  the room instead of the content bounds, and a gradient paper spans the scene, not the canvas.
- **`.dgf` paper**: `PaperColor1`/`PaperColor2`/`GradientPaper` of `[General]`, top to bottom.

## GeoGebra files (.ggb)

`Serialization/GeoGebraReader.cs` opens GeoGebra worksheets (Open dialog, command line,
`--check`): a `.ggb` is a zip with `geogebra.xml` inside. The construction is free elements
(`<element type="point">` with `<coords x y z>`, homogeneous: divide by z) and
`<command name="Segment"><input a0="A" a1="B"/><output a0="f"/>` with the outputs' elements
(style, show, coords) following the command. Each known command becomes the figure that
stands for it here, with hidden helpers where the shapes differ (a circle through three
points is the circle around the crossing of two bisectors, an ellipse by foci is center and
axis ends as points by coordinates, a regular polygon is a plain polygon of rotated points so
that its vertices keep the file's names, tangents from a point go through the Thales circle);
an inline command in an input (`Point[Circle[S, 3]]`) is built hidden. A polygon's sides are
outputs of the `Polygon` command and objects of their own in GeoGebra; here the polygon draws
them, and a side becomes a segment only when needed: hidden, the first time a command takes
it (`polygonSides`), or in view and under the file's name when its element shows a name or a
decoration (`ApplyElement`), the name outside the polygon (`FigureLabel.KeepOutside`). Unknown commands and
element types (conics other than circles, pen strokes, buttons, checkboxes, lists) are left
out, along with what is built on them: the status says only that some features aren't
supported, the list of what exactly goes to the console (`GeoGebra: ...` lines, which is
what a `--check` run shows). GeoGebra's expression
language is translated only where it overlaps ours (`x(A)` is `A.X`, `°`, `Name[A]` in texts);
a point whose expression doesn't translate becomes a free point where it was. Names come as
they are (`A_1`, `Q'`; only `A_{12}` loses its braces, which no expression could parse). The
drawing keeps GeoGebra's look, on purpose: every
element carries its color, point size, thickness (stroke width is thickness/2), line opacity
and fill alpha, and the file's defaults are GeoGebra's (blue free points, gray dependent ones
and lines). A point is 2 × `pointSize` across with a rim a shade darker; `pointStyle` cross
and plus are drawn as the characters, ring and hollow diamond as unfilled shapes; a dynamic
color (`dynamicr/g/b`) wins over the static one; a pattern fill (`fillType` hatch, dots...)
becomes a translucent fill. Point names are in the point's color, at GeoGebra's place (the
baseline starts a radius to the upper right, plus `labelOffset`); `labelMode` 1 and 2 show
the coordinates. Not carried over: segment end styles and decorations (arrows, ticks),
captions. The view is the file's (same zoom, same middle). GeoGebra's colors are made for a
white paper, so a worksheet stays light under the dark theme (`GeoGebraReader.KeepLightLook`):
its paper there is `#C0C0C0` (a Dark override; a worksheet with a paper of its own has that
under every theme), and the drawing's default styles are pinned to their Light look
(`FigureStyle.KeepBaseLook`: no theme bindings, each Dark override saying the Light value, so
that a saved `.lgf` keeps them - a default without overrides would be swapped for the
theme's on loading), which is what an element without a color and anything drawn later
gets. A number typed into a command (`Circle[A, 3]`, `Rotate[P, 45°, O]`, `Dilate[P, 0.2 * a,
O]`) becomes an auxiliary `Number`, or a hidden auxiliary `Label` evaluating `[expression]`
when it depends on figures (a label is a length and an angle provider). A named numeric
expression also stays live as a hidden Label (shown or not in the file); later expressions
refer to its `.Value`. One our language can't say, or works out to another value than the
file has (angles up to whole turns), keeps the file's value as a fixed Number, with a line
in the report: left out, it took everything built on it along. Ordinary
numeric inputs to Rotate are radians, while imported angle Numbers already hold degrees.
Centroid uses the polygon's area-weighted centroid, not the average of its vertices.
A bare `°` (Rotate[C, °, A]) is one degree, as in GeoGebra. Helpers are built from the
points, not from expressions over their names, where they can be (the center of a regular
polygon is a rotated and dilated point: written as expressions over A', it stood at the
origin). To check the reader's geometry against GeoGebra's: every element of a worksheet
carries its coordinates (a point's, a line's equation, a conic's matrix), also when it is
worked out - compare them with what the import puts there.
A point given by an
expression (`A + (0, 1)`, `t B + (1 - t) A`) becomes a `PointByCoordinates` through a small
vector-arithmetic translator (`TranslateTerm`). A line typed as an equation (`x = x(P)`,
`y = m x + b`) is a live `LineByEquation` (the variable side read off by substitution at 0 and
1); one the reader can't read that way takes the saved coefficients, static. GeoGebra's angles
are radians in expressions and degrees in texts. `PointIn[region]` becomes a free point: it can
be dragged out of the region, where what is built on it may stop existing. Real-world files to test against: any material on geogebra.org downloads
as `.ggb` from `https://www.geogebra.org/material/download/format/file/id/<id>` (the id is the
tail of a `geogebra.org/m/<id>` link; a 403 means the author didn't allow it), and GitHub has
small sets (`kovzol/gg-art-doc` `ggb/`, `evenjung/geogebra`). Of 58 such files, the classic
constructions (bisectors, circumcircles, nine-point circles, Pythagoras proofs, transformations)
come through whole; what is left out is `If[...]`, lists and `Sequence`, `LocusEquation`, images,
buttons and checkboxes, 3D, custom tools.

## Gallery

- **Gallery and routes** (`LiveGeometry/Gallery/`, `MainView` "Pages" region): three states -
  gallery, a gallery drawing (`CurrentSample`: tour group in the toolbar, Page Up/Down, edits
  dropped silently, Save asks for a file and the drawing becomes the user's own), the user's own
  drawing (parked in `OwnDrawing` with its undo history while they look around; the name of
  the file it came from or was saved to, `OwnFileName`, sits alone in the tour group's
  place, centered, and in the page title - nothing for a new drawing). Paths `/`,
  `/gallery/<slug>`, `/drawing`; `AddressBar` is the abstraction, `BrowserAddressBar` +
  `main.js` do pushState/popstate. `index.html` needs `<base href="/">` for that (all routes
  serve it; `web.config` and `tools/serve.cs` fall back to it). All page changes go through
  `MainView.Show*`. Started at a drawing's address the gallery isn't built at all (the address
  is readable synchronously since `main.js` registers its imports before the runtime starts).
  The gallery page starts with a row of small tiles (`GalleryView.CreateStartTile`): New,
  Open (the editor's Open dialog; the page has no toolbar, so this and Ctrl+O are the ways
  to a file from there) and, while there is a drawing to go back to, My Drawing. They are
  92 wide so that all three fit across a phone 360 wide; the brand beside them gets what
  room is left, which on a phone with three tiles is none. A caption longer than
  "My Drawing" would be cut short.
- **Gallery drawings** are embedded `.lgf` (`Gallery/Drawings`, order and titles in
  `GalleryCatalog`) and are *the* source: edit them directly. They were forked once from the
  Windows Phone samples and the DG 1.0 CD library; `tools/gallerize.cs` made that fork and must
  not be rerun (it would overwrite the hand-edited drawings). To add another CD drawing:
  `--check` it (below), take the `.lgf`, drop the old labels, add `Title`/`Description` labels
  in the `GalleryTitle`/`GalleryText` styles (defaults, so the file needn't carry them), set
  `Grid`, list it in the catalog, and `--rewrite` the folder. Drawings with `Grid="true"`
  (graphs) keep their file viewport in view (`GalleryItem.Plane`), because graphs and lines
  have no bounds. All 47 were rewritten on 2026-09-29 (`--rewrite`: loaded and saved again,
  the file's viewport and the caption's pin, offsets and width kept, since opening lays both
  out for the window at hand): no copies of defaults, no `Style` on a figure with its kind's
  default, no attributes at their defaults, default names of figures as the loader gives them
  (`AB` for `Segment8`). Rerunning `--rewrite` on the folder changes nothing, so a generated
  drawing (below) can be normalized the same way after regenerating.
- **The caption** (`Title` + `Description`, which may contain live expressions - then list the
  points as dependencies) is pinned to the screen; the files carry no X/Y and
  `GalleryDrawing.Fit` recomputes pin, offsets and wrap width at open time: a column at the
  right in a wide canvas, a strip at the bottom in a tall one. Labels are sized in pixels, so
  the figure gets the canvas minus the text and the zoom is computed from that in one go -
  never iterate "place text, zoom to fit": it runs away once the text needs more than its
  share. If the text doesn't fit it runs off the bottom and the reader drags it up: labels of
  a gallery drawing can't be dragged (`Drawing.FixedLabels`, not saved, off once the drawing is
  the user's own), a drag on one pans the view, and one that starts on the caption scrolls the
  pinned labels along (`PinnedLabelScroll`, by pixels, so undo brings both back). Not
  `Locked`: a point counts as locked when anything built on it is, captions included. Refitted on
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
  The Five Platonic Solids (`tools/platonic.cs`, `--alpha` for a translucent variant). Captions
  of these three live in their tools too: change both. Hidden
  `PointByCoordinates` whose coordinates are expressions are the library's variables.
- **Drawings that must not fall apart** when a kid drags the wrong thing (Castle, The Falling
  Ladder): fixed points are `PointByCoordinates` with constant coordinates (a polygon of those
  has no free point, so dragging it does nothing), and the only things that move are
  `PointOnFigure` sliders on hidden rays or segments plus a free point or two.
- **The Spiral's rings** ("Drag to here") are placed for `Locus.StepCount` = 60 samples;
  changing `StepCount` moves them. The samples are taken by their count (61 points, 60
  equal steps of the sliding point's parameter): added up, the step's rounding decided
  whether the last sample but one was taken, and the curve had 60 or 61 points from one
  move to the next.

## Deployment and caching

- A push to main takes about 8 minutes to go live (GitHub Actions: build ~5, deploy ~3); the
  files flip at the very end. Until then every reload, cached or not, gets the previous version.
  The commit in the tooltip of the Octocat at the right end of the toolbar (`BuildVersion`, also
  logged to the console at startup) tells which build is on screen.
- Entry files are `no-cache` with an ETag that changes on every deploy, so a browser holding the
  previous version picks up the new one on a plain reload. If something looks stale, check the
  commit stamp and the Actions run before suspecting the cache.
- **`/history`** is a static page of its own (`history/index.html` at the repo root, no build
  step, everything inline but the Google Fonts; a relative URL in it would resolve against
  `/`, since `/history` has no trailing slash): the Browser csproj links the folder into
  `wwwroot/history/`, and a `web.config` rule answers `/history` with its `index.html` (a folder
  is not a file, so the SPA fallback would otherwise serve the app). `tools/serve.cs` has no
  such rule: locally open `/history/index.html`, or serve the `history` folder itself.

## Running instances

Always build into the normal `bin/` - no scratch output folders. The maintainer rarely has the app
running and doesn't keep state in it that matters, so a `LiveGeometry.Desktop` process that locks
`bin/` is almost certainly a leftover test instance, and killing it is fine in most cases.
Only if a build actually fails on a lock, enumerate and clean up:

`Get-Process LiveGeometry.Desktop | Select Id, StartTime, MainWindowTitle, Path` (pwsh), then
`(Get-Process -Id N).CloseMainWindow()` and, if it is still there a few seconds later,
`Stop-Process -Id N`. Alt+F4 through winauto can miss (it goes to whatever window of the app has
focus, e.g. a popup), so check that the process is really gone. Target test windows by `pid:`/`hwnd:`
rather than by process name when more than one could exist.

## Batch-checking drawings

`LiveGeometry.Desktop.exe --check <folder> <out>` (`MainView.BatchCheck.cs`) opens every
`.lgf`/`.dgf` under the folder (or the one drawing, given a file) in the real editor, zooms to fit (a captioned drawing is laid
out by `GalleryDrawing.Fit` instead, with the ribbon folded when the window is small - so the
sheet shows the gallery layout at whatever size the window was last closed at: `place` it,
close it, then run the check), saves `<out>/<relative path>.png`
and `.lgf` (the conversion) and appends to `<out>/report.txt`: figure counts, `NOT EXISTING`
(only the *root* failures - figures whose dependencies all exist - plus a dump of every point),
load errors. It exits when done. `dotnet tools/contactsheet.cs -- <png folder> <out.png>
[columns] [tile width]` tiles the PNGs into one image: the fastest way to eyeball a whole
folder (the whole gallery fits on one 4-column sheet). The VB6 CD library (221 files, kept
outside the repo, read-only) all loads; what the report still calls "missing" there is second intersections that fall outside
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
  with. "Convert to opposite angle" turns `Interior` off and reads the sides in the order
  (`Flipped`) that points the bisector the other way; the verbs it would inherit from `Ray`
  (Convert to line/segment, Reverse) are vetoed in `CanEdit`. Morley's trisectors are expressions and get the same effect from
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

`dotnet run tools\regression.cs` runs focused construction, property/editor, dragging,
undo/redo, expression-cycle, and LGF/GGB regression cases against the desktop libraries,
including loading and reloading every gallery drawing. Pointer cases briefly open test
windows. This is not an exhaustive UI or importer compatibility suite.

Screenshots are PNGs; image pixels are the click coordinates in both tools.

- `tools/winauto.cs` - any desktop window (the VB6 app, the Avalonia desktop app).
  `list`, `tree <t>`, `menu <t>`, `invoke <t> <menuId>`, `shot <t> out.png [--screen]` (the
  window's own rendering, or what is on the screen there - the only way to see the title bar), `click <t> x y [right|double]`,
  `move <t> x y` (hover, for click previews and `cursor`),
  `drag <t> x1 y1 x2 y2 [steps] [--shift] [--alt]`, `wheel <t> x y <notches>`, `keys <t> "^s"`, `text`, `focus`,
  `cursor`, `place <t> x y w h`. Target = process name | `pid:N` | `hwnd:0x..` | `title:substr`.
  - The VB6 app starts maximized: `place <t> 100 100 1500 1000` first. The Avalonia desktop app
    remembers its window (`LiveGeometry.Desktop/WindowPlacementPersistence.cs`, the
    `WindowPlacement` line of `Settings.txt` in the user's local app data, next to the theme
    choice). It is left at 100,100 1700x1100 for
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
    open, `shot LiveGeometry.Desktop` captures the menu, not the main window (and `list`
    doesn't show it). It opens at the cursor, so click its items by main-window coordinates:
    right-click point + the item's place in the menu shot, with `hwnd:` of the main window
    (by `pid:` a click can land on the main window's system menu).
  - After a figure with a length is made (segment, square, circle...) the length panel opens
    at the right of the canvas (about x 1270-1670, y 405-770 at 1700x1100) and swallows clicks
    there; the next click on the canvas closes it. Take a shot after each construction, or keep
    test clicks left of x 1250.
  - `winauto keys` sends real virtual keys for lowercase ASCII letters/digits (needed for the
    single-letter shortcuts); other characters go as Unicode packets, which KeyDown-based
    shortcuts never see.
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

## VB6 parity backlog

Deliberately out of scope for now: Calculator, step-by-step construction playback.

Still missing compared to VB6, roughly by value: symmetric point (about a point) and inverted
point (in a circle) tools; tracing locus of a point ("Create locus" on a point); "Choose point/figure" disambiguation for
overlapping figures; double-click opens properties (here: double-click = zoom to fit); measurement
label dragging constraints; point shape/size per point and name color; line dash styles per
figure; Show/Hide, message, sound and launch buttons; live cursor coordinates in the status bar;
rulers; undo/redo captions naming the action; unsaved-changes prompt; recent files; print;
"tool select once" option; settings persistence; languages (en/ru/uk/de).

## Not yet verified in the browser

Printing, demo download, and the `Hyperlink` figure (only in old files, no tool makes one),
which fetches the drawing at its URL with `HttpClient` when clicked - not when read, which
put a failure's whole stack trace into its text, saved with it.
Saving through the browser storage provider works (drawings, and PNG and SVG pictures);
opening has only been checked as far as the picker (see "Files in the browser"). Copy image
finishes without an error in headless Edge, but reading the clipboard back is denied there:
what a paste gives was only checked on the desktop.
