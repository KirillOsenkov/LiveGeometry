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
  On a Mac the executable is `bin/Debug/net10.0/LiveGeometry.Desktop` (no `.exe`), and the
  same build runs as is: see "macOS" below.
- Browser: needs the `wasm-tools` workload (`WasmBuildNative=true` links Skia/HarfBuzz; without
  it Skia throws DllNotFoundException). Do a full bin/obj clean when native assets change.
- Browser publish: `dotnet publish Main/Avalonia/LiveGeometry.Browser/LiveGeometry.Browser.csproj -c Release -o <dir>`.
  Output is `<dir>/web.config` + `<dir>/wwwroot/`. CI (`.github/workflows/main_livegeometry.yml`)
  does exactly this and deploys to Azure App Service (IIS). A publish compiles the .NET code
  ahead of time to wasm (`RunAOTCompilation` in the csproj, a few minutes more; see
  "Measuring speed" for what it buys and costs, and how to switch it off).

## The ribbon

Tools are `Behavior` subclasses found by reflection (`Behavior.LoadBehaviors`): `[Category]`
names the tab, `[Order]` the place in it, `[Ignore]` keeps one off the ribbon; the button shows
`Name`, the tooltip adds the letter from `UI/BehaviorShortcuts.cs`, the status bar shows
`HintText`. Tab order is the order of the constants in `BehaviorCategories`; tools the user
defines go onto Misc. Non-tool commands are added in `MainView.InitializeCommands`. Tab by tab
(letter in parentheses):

- **Selection**: Drag (Q) - drags points and figures (with Alt a point snaps onto figures and
  lets go of them, see "Snapping and releasing points"; while a point is dragged the
  status says what Shift and Alt do to it, `Dragger.ModifierHint`); a press on a figure
  selected with others drags them all as one piece (`Dragger.FindSelectionRoots`: the roots
  of the whole selection, each once, by the same offset; no Alt snapping, no point jumping
  under the cursor; a locked figure or root holds them all, a caption stays); also the tool
  every construction returns to. Figure List (toggle): see "The Figure List".
- **Points**: Point (P) - free, on a figure, or at an intersection; Midpoint (M) - two points
  or a segment; Intersection (I) - two figures that cross, the click on the second picks the
  nearer crossing (`PointPlacement.Intersection`, shared with the Point tool); Coordinates (X) -
  the Point tool with its X/Y panel always on (`PointByCoordinatesCreator`), so a point by
  coordinates needs no trip to the Coordinates tab's toggle (that toggle stays for typing
  the points of other figures); Label new points (toggle).
- **Lines**: Segment (S), Ray (Y), Line (L), Vector - two points each; Parallel (N) and
  Perpendicular (E) - a line then a point; Angle Bisector
  (B) - vertex then two side points, an angle's number or mark, an arc, a sector or a
  segment (its central angle: center and the two ends; `AnglePoints` turns any of those
  into the three points and the sweep, which the bisector copies once, as Circle by
  Radius takes a segment for two points), or one click inside an angle
  next to its vertex as the Angle tool takes it (see "An angle at a vertex" under Measure);
  Line at Angle - a point, at the
  angle in the tool's panel (0 until changed, so a horizontal line is one click), or click an
  angle measurement, its arc or a slider first to tie the angle to it, or inside an angle
  next to its vertex (which measures it first, mark and number, and ties the line to the
  measurement); right after the click
  the panel shows the new line's angle instead (see "Tied values"). Perpendicular Bisector
  (two points or a segment; `[Ignore]`d since 2026-10-04: Perpendicular through the midpoint
  does it, and its click on a segment took the ends where every other tool puts a point on
  it - files and GeoGebra imports still make them); Join segments (a point between
  two segments joins their other ends) and Polyline (points, double-click or click an
  existing point to finish) exist but are `[Ignore]`d as rarely used.
- **Circles**: Circle (C) - center then a point on it; By Radius (R) - two points, a segment, a
  distance or a slider, then the center; a first click on empty paper makes a slider for the
  radius (see "Sliders"); Ellipse - center, end of the long axis, end of the short axis;
  Circular Arc (A) - center, start, end (counterclockwise); Elliptical Arc - center, semi-major,
  semi-minor, begin angle, end angle. Right after either arc is made the side panel offers
  the arc's `Sweep` row (counterclockwise, clockwise, under or over 180°: see "Which of
  the two angles"), Convert to sector / segment and OK (`ArcPanel`, 2026-10-09): a sector
  and a circular segment are the arc's three clicks, so they have no tool of their own,
  and the verb gives the shape a filled style in the arc's hue (see "Arcs, sectors and
  segments"). A verb
  clicked in that panel ends in the same panel for the new figure, with the verbs to the
  other two kinds, so one more click goes on or back (`ArcPanel.Convert`); the same verb
  clicked in the figure's own grid ends in the new figure's grid (`EllipseArc.Convert`,
  over the shared `Replace`), which also takes the selection over.
- **Shapes**: Triangle - 3 points; Square - two adjacent vertices; Polygon (W) - points, then
  Enter, a right-click or a click on a vertex closes it; Regular polygon - center then a vertex.
  Triangle and Polygon show no length panel (a side's length means nothing for them). They
  (and Square, for its first side) draw their sides as segments (the polygon's own outline is transparent in the default
  style), except where a visible segment, ray, line or vector on the two points is there
  already (`FindLine`; not any line that depends on both - a perpendicular bisector of the
  two took the side's place and left it undrawn). (Polygon intersection
  exists but is `[Ignore]`d.) Bezier path (last, no letter): see "Bezier paths".
- **Transform** (the tools are verbs, as the tab is; the classes stay `ReflectionCreator`...):
  Reflect (T) - source figure, then a mirror (point, line, segment, ray, or a
  circle for a point source); Rotate - source, center, angle (a figure with an angle, a
  typed value, or a click inside an angle next to its vertex, which measures it first and
  rotates by the measurement: "An angle at a vertex" under Measure); Translate - source,
  distance, direction (a figure with an angle, a line as it points, a vector, a typed value,
  or an angle at a vertex measured the same way; see "TranslatedPoint"); Dilate -
  source, center, factor (a figure with a length or a typed value); after each of the three
  the panel shows the new figure's values (see "Tied values"). Last of the tabs that
  draw with figures alone; the two after it work with numbers. A figure is transformed by
  transforming what it is built on, down to points (`Transformer`), so a source is taken
  only when all of that can be (`CanBeTransformSource`, recursive) - except a radius given
  by a number (a slider, a Number, a measurement: By Radius makes a slider whenever its
  first click is on paper), which the image of a reflection, rotation or translation
  shares and a dilation can't scale. Asked only about the figure itself, the tool threw
  at its last click. What can't be transformed that way but can carry a point (a locus,
  a graph, a line or circle by equation, a line at an angle, a Bézier curve, a circle by
  a number under Dilate) is *traced* (`Transformer.CanBeTraced`, `CreateTracedImage`):
  a hidden auxiliary `PointOnFigure` on it, its image by the same transformation (hidden,
  auxiliary), and a `Locus` of that image, in the source's line style - so
  deleting the locus takes both points along. Reflect in a circle (inversion) traces
  every source but a point: a line's image there is a circle through the center. A
  polygon there gives the images of its sides, not a polygon (no point is on a polygon,
  and a region bounded by arcs is no figure here: `Transformer.CreateInvertedSides`): a
  regular polygon's own side parts are traced, a polygon's side segments, and a side
  without one gets a segment first, as the shape tools would have drawn it. Reflect
  shows no panel after a construction (the length panel came up for such a segment).
  Every other transformation of a polygon or polyline also carries over the
  visible segments along its sides (`Transformer.AddSideSegments`: the shape tools draw
  sides as segments, and the image was a shape without an outline), same style and
  marks; they go into the list before the image, which callers take to be the last. An
  image is a clone of the source given the transformed dependencies; a composite's clone
  made its parts on the source's when it was read, and `Transformer.MovePartsOver` moves
  them (a reflected regular pentagon had two sides and its inside on the source's first
  vertex). The
  GeoGebra reader says `sideSegments: false`: a file has the image's sides as objects
  of their own, and the copies came out hidden.
- **Coordinates**: Grid (G) (command); Function - an expression in x; Line - by
  slope and intercept expressions; Circle - by center and radius expressions; Point by
  coordinates (toggle: gives the point tools an X/Y panel).
- **Measure**: Distance - two points or a segment; Angle (J) - vertex then two side points,
  the angle under 180° whichever side comes first (its `Sweep` row, see "Which of the two
  angles" under "Design decisions", says otherwise), or one click on an arc, a sector or
  a segment, which measures its central angle the way round the arc goes (`AnglePoints`:
  the center and the two ends for the points, the arc's sweep copied once; not an
  angle's own mark or number, measured already), or one click inside an angle
  next to its vertex, within the reach of the mark it would get, where drawn lines,
  segments, rays or polygon sides leave a point (`AngleAtVertex`: the hover shows the mark
  and number faint, halos on the sides; `FigureCreator.TakesAngleAtVertex` is the hook, a
  step's yes or no, and `ClickAngleAtVertex` what the click does: the Angle tool and the
  bisector take the three points, Line at Angle, Rotate and Translate measure the angle,
  `MeasureAngleAtVertex`, and take the measurement, so one undo step holds both; a
  figure the step takes under the cursor comes first). Only when that is the one angle
  the cursor can mean: two vertices in reach, a third line on the same side (a bisector: the half or the
  whole?), a side with no point on it to depend on, an angle measured already - nothing,
  and the click takes a point as before (as it always does on a point); Area (K) - a polygon, ellipse,
  circle or list of points, Enter or a right click when the points are done; Perimeter
  (2026-10-09) - one click on a polygon, circle, ellipse, sector, circular segment or closed
  Bezier path (`IPerimeter`, see "Perimeter" under "Design decisions"; no letter yet);
  Slider - where it sits, then where
  its knob starts (or press, drag, release): a number with a handle, taken wherever a tool
  asks for a length or an angle, named in expressions (a, b, c).
- **Misc**: Bezier - four points; Locus (D) - a point that depends on a point on a figure;
  Text - a label at the click; Show/hide box (`ShowHideCreator`) - the figures selected
  when it is picked are picked already, a click on a figure picks or lets go of it (a
  name for its figure, a vertex for its polygon; not boxes, axes, Numbers or `Auxiliary`
  helpers), a click on a row of the Figure List too, the only way to a hidden figure, which
  then shows as a ghost (`IFigurePicker`: a row click picks instead of selecting, Space
  picks the keyboard's row, Delete does nothing meanwhile; the picks are selected figures,
  and `Drawing.PicksChanged`, not `SelectionChanged`, repaints the list, or the side panel
  would trade the tool's panel for the selection's); a click on the paper puts the box
  there, ticked if any of its figures shows ("Show"), else unticked ("Hint"), so that
  nothing changes on screen, selected with the keyboard in its Caption. The box's Edit
  figures (grid, context menu) starts the tool on it (`ShowHideCreator.Edit`, the ribbon's
  instance: `Behavior.FindTool`); the paper, the box, Enter or OK ends it, one undo step
  (the box moves to the end of the list when it takes a later figure). Under the Drag tool
  a press on a box is the tool's (`ShowHideCheckBox` leaves it to bubble to the canvas):
  a click ticks it (`ShowHideControl.Click`), Ctrl+click selects it, a drag moves it (a
  locked box stays put and still ticks; a box the gallery drawing came with is in
  `Drawing.FixedLabels`, and a drag on it pans); a right click selects it. A figure
  deleted leaves its box (`ISupportRemoveDependency`), the last one takes the box along.
  Define figure - records a construction as a new tool: click
  the figures it starts from, OK, click the figures it makes, Create tool (neither step
  goes on with nothing picked); a halo and the hand cursor show what a click would select
  or let go of (`FigureSelector.FindFigureToToggle`). The new tool lands on Misc, named
  after the first result clicked (Catenary, then Catenary 2: `Behavior.UniqueToolName`
  over the ribbon's tools), and is picked at once with its panel open and the keyboard in
  the name (`ToolNameEditor` refuses an empty name or another tool's; OK or Enter only
  puts the panel away). It is kept between runs, never in a drawing (what it makes is
  plain figures): `StoredTools` stores each macro as a document of the settings store
  (`SettingsStore.GetDocument`: `Tools\<key>.xml` beside `Settings.txt` on the desktop,
  `LiveGeometry.Tools.<key>` in the browser's local storage, a key per tool so that two
  windows or tabs don't write over each other), when made, renamed, and gone with "Delete
  this tool"; read at startup after the ribbon, by key (a time stamp). The macro is
  `<Macro Version Name>` - the version is the drawing format's, and one from a newer
  version, damaged or asking for an unknown kind is left out with a console line
  (`UserDefinedTool.Read`) - then `Inputs`, `Icon`, the `Styles` its figures name (brought
  in as a paste brings them, `PasteAction.BringStyles`) and `Figures`, where a figure that
  had its default name says `DefaultName="true"` and gets the default name where it is
  made (segment AB on P and Q is PQ; a name that reads like the default reads like a
  typed one on other points). It asks for its inputs in the order
  they were clicked (`FigureSelector.GetSelection`; in the drawing's order, a slider
  clicked first was asked for last), and the expressions of what it makes (a point by
  coordinates, a label's [AB]) are rewritten to name what it was given and the copies,
  as a paste's are (`UserDefinedTool.RebindExpressions`): compiled by the macro's names,
  a tool defined on the Catenary drew the first curve again. Its icon is a picture of the
  inputs and results as they were at Create tool (`MacroIcon`, kept in the macro as
  `<Icon>`, in the icon's units): inputs as yellow points and ink, results as
  constructed points and the accent color, lines and graphs cut to the picture's box;
  figures without a simple picture (labels, measurements, angle marks) are left out, and a
  tool with nothing to draw shows a dot and the number of its inputs.

## Avalonia and framework traps

- **Trimming only happens on Release publish**, never in `dotnet run`. The geometry library is
  discovered via reflection (`GetTypes()` for behaviors, serializers, figures), hence
  `TrimmerRootAssembly DynamicGeometry.Avalonia` in the Browser csproj; `LiveGeometry` is a root
  too (the property grid reads `ExceptionReport` by reflection, and `[DynamicallyAccessedMembers]`
  on the class did not keep its rows). A trimming break shows up as a white screen and a
  `CRASH: ...` console line (`Program.cs` prints those). Anything new that is reached only by
  reflection outside those assemblies needs its own root. The functions of the expression
  language that are System.Math's are found by name (`Binder`, a `Type` out of a list, which
  the trimmer can't follow): `[DynamicDependency]` on `Binder`'s static constructor keeps
  them all. Without it the browser had only those the app also calls itself, and `asinh`,
  `sinh`, `cosh` were "Could not find method" - the Catenary drew no curve there, while the
  desktop (never trimmed) drew it.
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
- **The page can have no size**: a browser lays it out at 0x0 (a tab opened in the
  background, DevTools taking the whole height), and a desktop window shrunk to its title
  bar leaves a canvas of a few pixels. Arithmetic on a layout size (`finalSize.Height -
  top`, a room beside a caption) goes negative there, and Avalonia's `Arrange` throws
  "Invalid Arrange rectangle" (the toolbar's tab did, 2026-10-06); a scene chosen for a room
  of no area is null. Code that lays out or fits by sizes must survive 0 and a few pixels,
  and a fit must not depend on the view it came from (`GalleryDrawing.Fit` measured the
  figure's labels at the previous zoom: shrunk and grown back, the figure came back at half
  its size). Checked by laying `MainView` out at a grid of sizes from 0 to 800 px on every
  page and side-panel state, with every first-chance exception logged (a temporary
  in-process probe, as for undo); Windows and headless Edge can't make windows that small.
- **`Shape.Render` is sealed** (Avalonia 12): a Shape can't draw anything but its geometry.
  `PointMarker` draws a character through an `EmojiGlyph` visual child it measures and arranges
  itself; with no geometry, `Shape.ArrangeOverride` returns size 0 and the shape collapses to
  the middle of its place, so the override returns the final size then.
- **A `FontFamily` in a collection that isn't registered yet** ("fonts:Emoji#...") makes text
  layout throw *while rendering*, which ends the desktop app. `EmojiFont.Family` is the default
  family until the font has loaded.
- **`RotateTransform`**: no CenterX/CenterY - Avalonia rotates about `RenderTransformOrigin`
  (the middle by default), and WPF-style centering shifts a tilted ellipse off its center.
- **A TextBlock drops lines that don't fit its arranged height**: it lays its text out again
  at arrange time, in the arranged size, and leaves out every line past its height (they are
  not clipped: dragging the label doesn't bring them back). On an iPhone only (3x pixels;
  never seen on Windows or in headless Edge) a wrapped caption lost its last line, also when
  laid out at least as high as it measured, so the arranged layout wrapped into a line more.
  A label's text is a `LabelTextBlock`, laid out again at the width it measured at and with
  no height limit; that fixed it (2026-10-03). The exact cause wasn't found.
- **In the browser a posted job can keep the page from painting**: Avalonia's dispatcher
  runs what is due (Background jobs, a short timer) in the same JavaScript task as the
  animation frame, before it gives the page back, and the browser paints only between
  tasks. The gallery's first batch of tiles, posted by the first layout, kept the splash
  (which Avalonia closes at the first frame) up a second longer. To let a frame be seen
  first, start the work at the frame after it (`TopLevel.RequestAnimationFrame` twice:
  `GalleryView.OnAttachedToVisualTree`).
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
  key) deletes the selection as Delete does. Text that names a key names it as the
  keyboard does (`KeyNames.Alt`: Option on a Mac); `KeyNames.IsMac` is the desktop's OS,
  and in the browser the page's platform (`isMac` in `main.js`), since a browser says it
  runs on "browser". The browser sets it before Avalonia is up, which is why it is a class
  of its own: touching `Behavior` that early ran its static constructor, which makes
  cursors, and the page died with "Unable to locate ICursorFactory".
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
- **The axes are lines to build on** (`Figures/Lines/AxisLine.cs`, VB6's "active axes"): a
  point on the x-axis, where a circle crosses it, a parallel to it, a reflection in it. The
  grid draws the axes; an `AxisLine` is never drawn (a transparent stroke, for hit testing and
  a selection's halo). A drawing has one object per axis for its whole life
  (`Drawing.GetAxisLine`), and whatever reads an `<AxisLine Axis="X" />` - a file, a paste, a
  tool the user defined, the GeoGebra reader's `xAxis` - gets that one
  (`DrawingDeserializer.ReadFigure`): never two x-axes. They are in the list only while
  something is built on them: a figure added on one brings in both (`AxisLine.AddMissing`,
  from `AddFigureAction` and `PasteAction`, which take them out again on undo; not while a
  file is read, which has its own in their places), and they leave together when nothing is
  built on either, also when one of them is deleted (`FindOrphanedAuxiliaries`). Out of the
  list, hit testing still finds them (`RootFigureList.HitTestCandidates`) - only while the
  grid shows its axes (`AxisLine.IsHitTestVisible`); what is built on them stays when the grid
  is hidden. Real figures win over an axis (lowest z; `PointPlacement.FindOnFigures` puts it
  last), the Drag tool, the context menu and Define figure ignore it (a press pans), and the
  hover halo is drawn from the coordinates (`ClickPreview.CreateHalo`): an axis out of the
  list has no shape on the canvas. As a transform source it is traced.
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
  exist either: it stood at (0, 0). Nor a line by equation without a value (A = B = 0, a
  slope of sqrt(-1)), nor a circle by equation with a radius below 0 or undefined (it was
  a dot), nor a point dilated by a ratio over a length of 0.
- **Overlapping figures: Tab chooses** (`Behaviors/ClickChoice.cs`, VB6's "Choose
  point/figure"). A click takes the first of what it could mean, as hit testing orders it
  (`FigureList.HitTestAll`: the order `HitTest` picks from - topmost ZIndex, the nearer of
  two points, the newer), so a plain click does what it did before there was a choice. A
  tool lists what a click at the cursor could take (`Behavior.FindClickOptions`): for a
  `FigureCreator`, the figures taken instead of a point (`FindFiguresInsteadOfPoint`),
  then the points it could make (`FindPointPlacements`: the points there, else
  `PointPlacement.FindAll` - the default, the other crossings nearby, the midpoint, a point
  on each figure, never a free point next to them, or every hover over a line would be a
  choice) - or, when a step wants a figure, those it takes (`FindExpectedDependencies`,
  where a tool says its own rules: only what can be transformed, a crossing figure). The
  hover offers the list (`ClickChoice.Offer`; the same list again keeps the choice), Tab and
  Shift+Tab step through it (`MainView_KeyDown`, not from a box, the side panel or the
  Figure List; `Behavior.StepChoice` shows the hover anew), and every finder picks from its
  own list through `ClickChoice.Pick`: the first while nothing was chosen, the chosen one
  when its list has it, nothing when another list does. A choice holds for one click
  (forgotten after the press, on leaving the canvas, for typed coordinates). The status
  says it ("A point on circle c (2 of 3): Tab for the next.") through
  `Drawing.ChoiceStatus`, over the hint, which `DrawingHost` brings back when it goes. A
  tap where there is a choice asks in a menu (`Behavior.AskWhatTapMeans`) and is the
  option picked; dismissed, it is nothing. A tool whose click on a figure is a shortcut
  (Midpoint on a segment, Angle Bisector on an angle, Area on a shape, Angle at a vertex)
  offers no choice where it applies, or the status would promise what the click doesn't
  do. Under the Drag tool the status names whatever is under the cursor, with its
  construction ("Midpoint E of AB"; `ClickChoice.DescribesSingle`, `DescribeInFull`), and
  every figure there is an option, in the order a press takes them (a point over the
  segments through it: "Point A (1 of 3)"), so Tab reaches a figure under another and the
  press, the right click and the halo go to the one chosen - not an axis or a caption,
  and nothing while a press is held (the status says what the keys do to the dragged
  point). A finger is offered the handles of a Bezier path only, or every tap on a point
  on a figure asked in a menu. Its context menu also lists the figures under the click
  ("Choose figure") and selects the one picked.
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
  a file is read it waits, `Drawing.IsReading`, since expressions are compiled by
  the file's names). Handed out to whoever asked first, AB2 stayed AB2 after AB was
  deleted until the file was opened again, and undo of a join gave segment AB and line
  AB2 their names back the other way round. The points are read in
  the order that comes first alphabetically (A-Z, then A1-Z1) among the readings that name
  the same figure: a polygon from any vertex either way round (ECBA is ABCE), a segment,
  line, polyline or Bezier either way, a ray or vector only as it goes (`PointOrder`); any
  of those readings counts as a default name. A figure no points name takes a lowercase
  letter (since 2026-10-04; `FigureBase.FirstLetter`, `GenerateLetterName`): from g for
  lines, c for circles, ellipses and arcs, f for functions, p for polygons, u for vectors,
  each counting on through the pool sliders use (`FigureBase.Letters`, no e, x, y) and then
  again with 1, 2..., never round to the start (a function took a, a slider's letter).
  Hidden helper lines and circles take letters too: other figures' constructions name them
  ("of line g and segment AB"), and a PerpendicularLine1 there was a name no row of the
  Figure List showed. Numbers stay n1, n2; measurements, texts, angle marks, loci, parts of
  a composite and hidden points are numbered by type (Circle1; a hidden point doesn't take
  a letter from the points on screen). Old drawings keep their Circle1-style names;
  `LiveGeometry.Desktop.exe --rewrite <folder> --letters` gives them letters (the
  gallery's were, 2026-10-04: only `Name` attributes changed, checked by swapping the
  names back into the old text, comparing every figure loaded both ways, and pixels). `HasDefaultName` (nobody typed a name) is not stored: a name that reads like the
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
  "Midpoint E", "Regular pentagon p"), the kind left off when the name says it already
  (Circle1, Bezier3), the name left off when it is the type and a number on a figure
  nothing refers to by name (Distance, Text, Angle mark, Locus: `NamedByConstruction`) -
  not on others: "on Circle₁" in one row must be found in another.
  After the title comes `Construction`, how the figure is built ("of AB", "to segment AB
  through C", "with center A and radius 3"): faded after the title in the Figure List, a
  smaller line under it in the grid's header. It names what it is built on through
  `ConstructionText`: a point by its name, anything else by its `Noun` and name (segment
  AB, line g - the plain noun, not "parallel line"), a typed value by its number, a tied
  one by what it comes from (a, AB, angle ABC). "angle ABC", not ∠: Inter, the browser's
  only font, has no ∠ glyph.
  A polygon's kind is by vertex count (Triangle, Pentagon, Hexagon, else Polygon); four
  vertices go through `Quadrilaterals.Classify` (Square, Rectangle, Rhombus, Parallelogram,
  Kite, Trapezoid) with a rounding-only tolerance: a shape dragged to look square by eye
  stays a Quadrilateral. The title is refreshed on every move (`PropertyGrid.RefreshNumbers`).
  `ToString()` stays the bare name: messages and dumps use it.
- **A name's index is a subscript on screen** (`Figures/NameDisplay.cs`): `A_1` draws as A₁
  on the canvas (point labels, slider captions) and in `Title` (the grid's header, the
  Figure List), as GeoGebra and TeX read an underscore; the trailing digits of a name
  without one (`A1`, our own default names after Z, `n1`, `Circle1`) draw as a subscript too,
  and so do those of every point in a name of points run together (`G1H1IJ1` is G₁H₁IJ₁).
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
- **An expression is bound once and evaluated many times** (`Expressions/`, 2026-10-07):
  `ExpressionTreeBuilder` binds the parser's `Node` tree into a `BoundExpression` (names
  resolved to figures, numbers, two-point distances, properties and functions, every value
  a double), and `Compiler.Strategy` (`ExpressionStrategy`) says what the delegate run at
  every recalculation is made of: `OwnTree`, the bound tree evaluating itself (the
  default); `LightCompiler`, a System.Linq.Expressions tree run by that library's
  interpreter (what every expression was until then); `LinqInterpreter`, the same tree
  walked by `ExpressionTreeInterpreter`; `Compiled`, the tree compiled to IL (a dynamic
  method for the JIT on the desktop; the browser has no JIT and interprets it). All four
  give the same values: the regression suite's "Expression strategies agree" and the
  benchmark's mismatches column say so. Reflection is asked once (`ExpressionReflection`:
  a function by name and argument count, a method's shape, a method or a property getter as
  a delegate - a getter through an open delegate made by a generic method, whose
  instantiations are named so that Mono's AOT compiles them), and a `Binder` gathers the
  drawing's point names once per expression, not per identifier.
  `LiveGeometry.Desktop.exe --bench-expressions <file>` (the browser build at
  `/?bench=expressions`, read with `webauto console`) loads every gallery drawing under
  each strategy, after a warm-up pass, and prints the load, the compile share of it, the
  evaluation of every expression 200 times, the function graphs sampled, the drawings
  recalculated 20 times, and the mismatches (`ExpressionBenchmark`). 2026-10-07, the
  gallery's 3572 expressions, compile out of the load and the 200 evaluation rounds, in
  ms: desktop JIT - OwnTree 46 (8%) and 96, LightCompiler 60 and 144, LinqInterpreter 61
  and 443, Compiled 528 (68%) and 17; browser AOT - 38 (6%) and 188, 138 and 578, 54 and
  530, 277 (31%) and 96; browser interpreted - 168 (8%) and 246, 626 and 1412, 453 and
  3124, 1044 (23%) and 139 (there the later strategies also pay to warm the interpreter
  up on their own code). So expressions are a small part of a load either way (Stretchy
  Slime's 160 compile in under 2 ms under AOT) and an evaluation is a fraction of a
  microsecond; the bound tree is the cheapest to make and second only to IL at running,
  which is why it is the default. IL wins evaluation everywhere but costs 3 to 7 times the
  compile.
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
  target built on the point. What is built on both collapses first (`PointSnapping.Collapse`):
  C onto B with segments AB, BC, CD deletes BC and leaves AB, BD; a polygon or polyline
  where the two are neighbors loses the point (ABCD becomes ABD), anything else goes as
  Delete takes it. Rewired, segment EF with F onto E was a segment EE, and undo's
  ReplaceDependency swapped its ends; and refusing the join made the snap fall through to
  the segment beside the point. An expression naming both still refuses it. In a join the
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
  the keyboard's row to its dependencies (blue), and with `FigureExplorer.RecursiveArrows`
  (off; no UI for it) on up from those (gray), one trunk per figure, a step left of the
  trunks above it that it runs beside, its branches stopping short of the next trunk: so no
  two arrows cross (`LayOutTrunks`). It rebuilds on `ActionManager.CollectionChanged`
  (posted, coalesced, and through `Throttle` at most every 300 ms: a drag of a figure or the
  view records a move per step, and it rebuilt every row at every frame), never between `ConstructionStepStarted` and a complete step: temporary
  figures of a tool are never recorded and its real steps sit in its transaction. Undoing a
  deletion puts each figure back at its old index (`RemoveFigureAction.Indices`), so it
  doesn't jump to the end of the list. A hidden figure that is selected is shown ghosted
  (`ShapeBase.IsGhost`, half opacity, shape not hit-testable) until unselected; `Visible`
  stays false, so no hit test, snap or drag sees it (they all ask `Visible`;
  `HitTestShape` refuses hidden figures too). Only figures directly in the drawing: a
  composite's parts (a vector's hidden segment) never ghost, so a hidden composite shows
  nothing. `UpdateVisual` overrides that skip hidden figures test `IsShown` instead, or
  the ghost sits at a stale place. A figure hidden when it comes onto the canvas keeps
  its shape off the canvas until it first shows, Visible or a ghost
  (`ShapeBase.OnAddingToCanvas`, `PutShapeOnCanvas`, 2026-10-07; shapes proper only, a
  label is measured as shown while hidden and wants the tree's font): the hidden helpers
  of a construction and the variables of a generated drawing (193 of Stretchy Slime's 264
  figures) cost the canvas nothing at the load and at every frame. The shape is styled all
  the same, since code reads a point's size or a line's thickness off it hidden or not
  (`Arrow`, `CoordinateSystem.LimitZoomByPointReach`). `Actions.ReplacePoint` puts the replacement in the
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
- **The default styles** (2026-10-03) are a palette of eight hues (`Styles/StyleHue.cs`:
  gray, red, orange, brown, green, cyan, blue, purple), a column each in every style picker,
  in full rows of eight: points (the kind defaults and red, blue, purple at 8; beads, 12 with
  a highlight), lines (thin 1.5, thick 2.5, dashed), shapes (outlined with a flat fill - the
  lighter end of the gradient, on either paper - outlined with the gradient, and gradients
  without an outline, but for the brown and green columns of that row: the fill of a new
  polygon, `Shape`, flat yellow from the theme's `ShapeFill` as it always was, and the
  classic flat green, `GreenShape`). Gray is the theme's where a default is (`Line`,
  `ThickLine`, the point fills). A hue is a stroke per paper; fills are diagonal gradients
  worked out from it (`Tint`, light to deeper on the light paper, bright to deep on the dark
  one, translucent; light tints at most 0.75 saturated, or cyan and orange glared). The
  brown column's fill hue is the ribbon's gold: of its own hue it was too near orange. The
  picker is a grid of fixed
  cells, eight to a row (`StylePickerEditor.CellSize`); a style kept for one purpose
  (`SliderTrack`, `GalleryLocus`) is offered only to its figure (`StyleManager.IsOffered`),
  or a row would be one short. A new kind of default wants a full row. "Create new style"
  on a figure that can be filled but has a line style (a circle takes either) makes a shape
  style, so that the fill is there to edit: the same stroke under every theme, a hint of
  its color for a fill, Filled unticked (`ShapeStyle.WithStrokeOf`).
- **Point shapes and emoji** (`PointStyle.Shape` / `Character`; `Size` is the shape's or the
  character's, whichever shows: the style keeps both in memory and the file only the one in
  use, so `Character` must be read before `Size`, which declaration order does): every point
  is a `PointMarker` (circle, triangle, square, diamond, pentagon, hexagon; the polygons reach
  past the circle by eye so they look as big), or one character instead, in the embedded
  Twemoji font (`Main/Avalonia/Fonts`, CC-BY, credited in the Emoji tab) with Inter named as
  the fallback: the browser has no system fonts, and a character neither has is not offered
  (`EmojiFont.CanDraw`). The font is 1.5 MB and loaded on first use, or as soon as the
  gallery is built, whose first rows show emoji (`EmojiFont.Open`: a file beside the desktop
  exe, a fetch of `fonts/` in the browser, brotli via web.config and cached as immutable: a
  different font must get a different file name). In the browser `main.js` fetches it, on
  the gallery's page as soon as the page loads (at a low priority, after the runtime's
  files), and hands it over in one piece (`fetchEmojiFont`, `copyEmojiFont`). The browser
  reads a response on the page's thread a piece at a time, and HttpClient took a turn of
  the UI thread for every step besides: fetched once the gallery was up, the font came in
  between the tiles' loads, seconds after its bytes. A font that fails to come (a network
  error) leaves the characters in the fallback font for good: each `EmojiGlyph` asked to be
  drawn again whenever it was drawn without the font, and the page drew without end. The
  style's editor has Shape | Emoji tabs (`IPropertyGridTabs`: the tab shown is what the style
  is; picking Shape drops the character, undoably; a row on two tabs, Size, gets an editor on
  each). "What the style is" is what it is under the theme on screen (`ShownCharacter`,
  and `ThemedValue` reads a property without an override from the resolved style): by
  the base values an emoji picked under Dark opened on the Shape tab, could not be taken
  off, and showed the shape's size. The Emoji tab searches `Emoji/Emoji.txt`
  (CLDR names and subgroups of the single-character emoji the font has): regenerate it with
  `dotnet tools/emoji.cs -- <emoji-test.txt> <font> <Emoji.txt>` when the font changes.
  A color emoji is painted with the style's `Fill`, so the fill's alpha is the emoji's
  opacity: a style written by hand without a `Fill` takes the shape default's (alpha 100)
  and comes out faded.
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
- **The figure list checks itself in Debug builds only** (`CheckConsistencyInDebug`: after a
  click of a tool, a recalculation of what is built on a figure, a file read): every
  dependency and dependent of a figure is in the drawing and registered both ways. A
  Release build (the site) never runs `CheckConsistency` on its own, so a registration bug
  shows there only by what it breaks later: catch them with the Debug desktop build, the
  regression suite and the harness, which call the check themselves.
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
  threw; `LineTwoPoints.IsThroughTwoPoints`), Convert to segment / sector on an angle's
  arc (`AngleArc`), the Text box of a measurement (`Measurement`, whose text is
  worked out), the style rows and buttons, Visible and Locked of a Number (nothing on
  the paper; Select all leaves numbers out, or a selection with one would lose those
  rows for every figure), a vector's Direction and an
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
  when it has nothing to measure). Convert to polyline (not offered since 2026-10-01: no
  tool makes polylines; a method is a button when it has `[PropertyGridVisible]` at all,
  `(false)` included, so it has none) deletes the polygon, and an area
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
- **Perimeter** (`Figures/Controls/PerimeterMeasurement.cs`, 2026-10-09): `IPerimeter`
  (`Perimeter`) on `PolygonBase`, `DependentPolygonBase`, `EllipseBase` (the arc integral
  `Math.EllipseArcLength` over a whole turn, which `ArcBase.Length` also uses), the sectors
  and circular segments (the arc with the radii or the chord) and `BezierPath` (closed:
  the cubics' lengths by Gauss-Legendre, `BezierInfo.Length`; open: NaN, and the tool
  doesn't take it, a measurement of it stops existing). Not `ILengthProvider`, on purpose:
  a circle's `Length` is its radius and a regular polygon's its side (`IFixableLength`),
  and `c.Length` in an expression reads that - the perimeter is `c.Perimeter`. The
  measurement itself is an `ILengthProvider` (a circumference as a radius), shows the bare
  number as Area does, and has a `Prefix` row ("P = ", saved as `Prefix`) for telling it
  from an Area next to it (a circle of radius 2 says 12.57 twice). Its anchor is on the
  outline, where a Distance sits on its segment: the middle of a polygon's first side, the
  lower right of a circle or an ellipse (the name's corner is the upper left), the middle of
  a sector's arc, the middle of a path's first piece; the default offset pushes it just
  outside (across the side away from the inside, else away from the center), worked out
  at the first `UpdateVisual` with a canvas that has a size (a file without `OffsetX`
  gets it too; before the first layout the anchor could only be the fallback below, and
  the offset came out wrong). A circle or
  an ellipse whose fixed point is off screen anchors at its point nearest the middle of
  the window, which moves with the view as a line's name does. Function graphs and loci
  are left out on purpose: their length depends on the window. A click on a regular
  polygon takes its `Interior` part, as the Area tool's does. The GeoGebra reader maps
  `Perimeter[]` and `Circumference[]`.
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
    if the old one was, takes its selection (outside the transaction, as a paste's: the
    verb was clicked in its grid, and the halo stayed on the figure that had left), and
    is named AB by its place in the list (`SettleDefaultNames`).
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
  `IAngleProvider` in degrees (Rotate). A slider may have a range (2026-10-10:
  `Minimum`, 0 unless said, and `Maximum`, none unless said; `<Slider ... Maximum="3">`,
  the Range group of its grid): the value is the minimum at the anchor plus the track's
  length, and the knob stops at the maximum, also when the value is set from the grid or
  the file. Dragging goes by parts (`IMovableParts`, which the
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
  or Distance measurement on one, and a circle given such an arc for its radius. Convert
  to sector / segment / arc (`EllipseArc.Convert`) trades a default style for its
  counterpart of the same hue (`StyleManager.ConvertStyle`, by the names a hue's styles
  have, `LineStyleNameOf`...): a red arc becomes a red-outlined, flat-filled sector and
  back to a red line, so the shape is seen at once; a style of the user's own stays where
  its kind fits, else the new kind's default (a sector's shape style on an arc didn't). A point
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
  360°), whichever of the two angles a sweep chooses (`AngleSweep.Measure`).
- **Which of the two angles** between two rays out of a point a figure means is one
  property, `Sweep` (`Figures/AngleSweep.cs`, 2026-10-09), on every figure that is such a
  region: the angle's mark and number, the bisector, and the four arc kinds with their
  sectors and segments (`IHasSweep`; `IArc` has it, `EllipseArcBase` keeps it). With θ the
  counterclockwise angle from the first ray (side, begin) to the second, the choices are
  `Counterclockwise` (θ), `Clockwise` (360° - θ: the same as naming the sides the other way
  round, kept as a choice so an arc's start stays the point it was built on), `Smaller`
  ("Under 180°" in the grid) and `Larger` ("Over 180°"). The first two never jump and run
  on past 180°; the last two trade places at 180°, where the number bounces and a bisector
  turns round - right for a triangle's angle, which stays the inside one when the
  triangle flattens and gets reflected; at exactly 180° the counterclockwise one is taken.
  Everything derives from it: the number, the arc drawn (`IsClockwise`, worked out, is
  what the path's `SweepDirection` and `IsAngleBetweenAngles` take), the fill, the
  bisector's direction, an arc's length and its sector's area, a point's parameter domain
  and the hit test; `IAngleProvider.Angle` is the chosen region's measure, so the sweep
  changes the number a Rotate, Line at Angle or Translate takes and never the way a
  rotation goes (GeoGebra's rule; `ang(A, B, C)` in an expression stays the bare
  counterclockwise θ). Defaults: the Angle and Bisector tools make `Smaller` (so the
  hover's wedge is what stays, also dragged past 180°: the creator no longer reorders the
  sides, and the name is the points as clicked), the arc tools `Counterclockwise` (an arc
  follows the cursor while drawn), and a conversion copies it. The mark and the number
  each have the row and keep each other in step (`AngleArc.SyncCompanionSweep`); a
  bisector built on a measurement takes the measurement's and its row is read-only. A
  reflection in a line mirrors it (`Transformer`: ccw and cw trade, the conditional ones
  stay), and a sweep that differs from the kind's default is saved as `Sweep="Larger"`.
  The arc's post-creation panel (`ArcPanel`) starts with the row, so an arc that came out
  the long way round is put right there. This replaced an arc's `Clockwise`, the
  bisector's "Inside the angle" with `Flipped`, and three "Convert to opposite angle"
  verbs, which each encoded a corner of the same table - and the Angle tool's hover
  reordered the sides (as `Smaller`) while the figure it left was oriented. GeoGebra's
  angleStyle 0, 1, 2 are `Counterclockwise`, `Smaller`, `Larger` (3, unbounded, counts
  turns and is read as 0); DG's angles and bisectors are `Smaller`. The player has the
  same property (`figures/angleSweep.js`).
- **Curves have gaps** (`Curve.Gap`, a point that is not one, in what `GetPoints` gives):
  each stretch between gaps is a figure of its own in the geometry. A function graph has
  one where the function has no value - the graph goes on to the very edge of where it
  has one (`FindEdge`), sqrt(x) starts at 0 - and where it jumps between two samples
  (`IsJump`: halving the step towards the larger change, a climb gets smaller and a jump
  stays), so 1/x, tan x and floor(x) have no vertical lines; values far beyond the window
  are clamped (a coordinate of 1e300 pixels is not drawn, and exp(x²) vanished whole). A
  function that throws has no value there (it was 0). A locus has one wherever the traced
  point doesn't exist. A locus samples adaptively (`Locus.SampleAdaptively`): 60 even
  steps, then round after round every step halved where the curve strays more than half
  a pixel from the piece drawn, or a gap begins or ends - at most 12 rounds and 1000
  samples, spread evenly when they run out (sin(40x) is coarse, not stuck); a piece still
  long after every round is a jump (1/x), not joined. Past the open ends of a line, ray
  or graph (and of a locus on one) it steps outward, each step twice the last, and a
  traced point beyond 20 window sizes is a gap. Sampled evenly along the line, the image
  of a line in a circle was a hexagon on its far side with a hole at the center.
  `Samples="60"` in a file asks for that many even steps instead: only the Spiral. Before, a curve was one line through all its points: straight
  pieces across where there is nothing.
- **Vectors** are a `Segment` (`Vector.VectorShaft`, drawn from the start to the head) plus
  an `Arrow` that draws only the head (`DrawsShaft = false`), sized in pixels, filled with the
  line color; both take the vector's style. The shaft was the arrow's filled outline, which
  no dash could break: a dashed style did nothing to a vector. An axis's arrow still draws
  its shaft. `Vector.OnAddingToCanvas` sets the default `LineStyle` before the base call,
  otherwise the polygon default (pale fill) wins. A vector is one figure to the user (a
  click anywhere on it selects the vector), not a figure with parts to select. A vector is an `ILine` (parallel,
  perpendicular, intersection, a point on it all take it) and its hit test asks the segment
  inside after the arrow, since the arrow is a filled polygon with no room around it.
  `Vector.HitTest(Point)` answers whether hidden or not, as a segment's does: a point on a
  vector and an intersection with one exist where that says, and through the composite's
  own test (shown parts only) they all went when the vector was hidden.
- **Bezier paths** (`Figures/Shapes/BezierPath.cs`, 2026-10-06; added beside `Bezier` and
  `Polyline`, which stay as they were): a composite like a regular polygon. The anchors are
  points of the drawing (its dependencies, at least two, each once); each has an in and an
  out handle, a part (`BezierPathHandle`), and the piece from one anchor to the next is
  the cubic through the first's out handle and the next one's in handle (a handle on its
  anchor: that end straight). A handle is an offset from its anchor, or a point of the
  drawing (`BezierPathHandle.Point`; the dependencies after the anchors, in the order in1,
  out1, in2...; they are the truth: a join or a replacement of such a point carries over,
  `OnDependenciesChanged`). Whatever changes what the path is made of (an anchor in or out,
  a hole, a handle's point) is one `LayoutChange`: worked out on a copy of the `Layout`
  (anchors, handles, sides, holes, the parameters of the points on it) and put in place
  whole, the old one back on undo. Closed and Filled are two check boxes (an open
  path fills as if a straight line closed it). Parts: the sides (`BezierPathPiece`,
  selected and styled one by one, "Sides" row), the inside (`BezierPathInterior`, the
  path's own style, "Fill"), the handles (`BezierPathHandle`, the `Handle` point style,
  offered to them alone, "Handles" row; never selected: a click on one keeps the
  selection). A handle shows only next to an anchor that is selected or dragged - its own
  two and the neighbors' that face it (`IsHandleShown`) - with dotted lines in the ink at
  0.3, and only the Drag tool hits one (`HitTest`): nothing is built on a handle but the
  images of a transformation. A side whose stroke draws nothing (a transparent color, as
  the `NoLine` style of a letter or a blob) is never hit: within the cursor's reach of it
  the click takes the filled inside, the path (2026-10-10; `BezierPathPiece.DrawsStroke`,
  the player's `drawsStroke`): the thin letters of Design a Font were all edge, and every
  click on one selected an invisible side instead of the letter. A zero handle sits under its anchor (z just below points):
  Tab takes it (`Dragger.FindClickOptions`, with the other figures there). A dragged handle
  takes the one across the anchor along as its mirror image (snapped to it at the first
  move: a symmetric anchor); with Alt that one stays where it is, for a corner
  (`MirrorsOpposite`, both in the move's undo place); dropped with Alt on a point, the
  handle is that point (`UsePointAsHandle`). A handle that is a point has no part shown
  (the point shows itself, with the dotted line), is never moved by the one across, and
  deleted it leaves an ordinary handle where it was. While drawing, Alt+click takes the
  handles between two anchors: the first the out handle of the one before, the second the
  in handle of the one after (closing: of the first) - a point where the Point tool would
  take or make one, else an ordinary handle at the click; a plain click on a point makes it
  the next anchor. "Convert to path anchor" on a point
  on a path (grid, context menu) makes it a free point there and an anchor between the
  ends of its piece, split by de Casteljau so the curve stays the same (the piece's handles
  that are points become ordinary ones) and the other
  points on it stay put (`ConvertToAnchor`). A point on a path depends on the path, parameter = piece index
  + the cubic's t (`PointPlacement.OrderForPoint` and Snap to take a side for its path); a
  point on the closing piece of a path opened doesn't exist. Delete an anchor and the path
  keeps the rest (two at least, `ISupportRemoveDependency`); Alt-drag an anchor onto a
  neighbor drops it, the target taking its handle on the far side
  (`CanDropAnchorInto`, from `PointSnapping.Collapse`; not onto an anchor further away).
  Holes: "Cut out holes" on a selection of two paths or more (`FigureSelection`) makes the
  largest one's inside leave the others out (`CombinedGeometry` Exclude), the holes being
  last dependencies, unfilled, still paths of their own; a
  deleted hole leaves the path. Transformations but inversion take it through its anchors
  and handles (`CreateImage`): the image's handles are all points, hidden auxiliary ones,
  which can't be dragged; an image isn't split or shortened (`IsImage`), nor is a path
  with an image split or an anchor dropped from it (`HasImages`: the image would keep the
  old pieces). A deleted anchor still leaves the source (its image goes with the anchor's
  image): the layout change runs before the image's helpers leave, so it recalculates
  without the Debug list check (`RecalculateAndUpdateUnchecked`). A deletion counts the
  anchors left after it, also those built on the point deleted (`CanRemoveDependency`).
  The figure the tension comes from is no part of what a transformation transforms. File: the anchors
  as the first dependencies and `Path="C 1,0 -0.5,1 L C #4 0,2 C a a ..."`, one piece per
  anchor, the closing one written whether closed or not: `L` for two handles on their
  anchors, else `C`, the out handle of the anchor and the in handle of the next, each an
  offset, `#k` (the dependency that is its point) or `a` (automatic); the dependencies
  nothing names after the anchors are the holes, but a `Tension="#k"`. No measurements and
  no intersections yet.
  **Automatic handles** (`BezierPathHandle.Auto`, 2026-10-06): a handle can be left to the
  path, worked out from the anchors in `Recalculate` by the path's `Smoothing`
  (`BezierPathSmoother`, a pure function: None - on the anchor; Hobby - METAFONT's, the
  default, four points on a circle give the circle's 0.5523; Catmull-Rom, centripetal;
  Natural spline, chord length). A clicked anchor's two handles are automatic, so each new
  click reshapes the pieces before it; press-drag, Alt+click, a drag of a handle (the one
  across it too: mirrored, or with Alt frozen where it was) make handles the user's.
  Handles the user set bound the automatic ones: across one an automatic handle continues
  straight on, next to one on its anchor (a corner) or at an open end it is a free end
  (Hobby's curl 1, zero curvature for the others). Every method is unchanged by moving,
  turning, scaling and mirroring the anchors (checked by the regression test), which is
  why images could keep automatic handles - they don't: an image's handles stay points
  following the source's handles, so they follow a change of mode too. `Tension` scales
  automatic handles down (Hobby's own tension, also in its equations, at least 0.75 there;
  a plain factor for the others); a slider in the grid from 0.5 to 3, or tied to a figure
  that says a number (`ITiedValues`: the panel right after the tool, a click on a slider;
  not what a point goes on, which starts the next path), the last dependency, typed again
  when that is deleted. A tension that is no number above 0 makes the path not exist.
  "Smooth automatically" and "Sharp corner" on an anchor (grid of a free point or a point
  on a figure, context menu of any point), "Smooth all anchors" on the path. Convert to
  path anchor between two automatic handles gives an automatic anchor (the curve moves a
  little), else the de Casteljau split. The tool's panel has Smoothing and Tension for the
  next path; the enum editor shows a value's `[PropertyGridName]` ("Catmull-Rom").
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
  dragged, and undo of that did not put it back. It hides where its side along the base would
  stick out past the end of a segment, ray or vector (`PerpendicularLineBase.BaseFigure`; a
  vector's inner segment, since its arrowhead has no points while a file is read and threw
  there). It hides when an `AngleArc` sits at the same vertex,
  because a measured angle of exactly 90° draws the same sign itself.
- **An angle is two figures**, `AngleMeasurement` (the number) and `AngleArc` (the mark, 0-3
  arcs), paired by `AngleArc.FindCompanion`. Neither exists while a side has no length (a
  point dragged onto the vertex, `AngleArc.HasSides`): not existing hides a shape and shows it
  again later (`ShapeBase.Exists`), where setting `Shape.Visibility` by hand is forever. The
  mark is an arc in code only: a sign of a fixed size in pixels, so no point goes on it and
  nothing is intersected with it (`PointOnFigure.CanBeOnFigure`,
  `IntersectionPoint.GetAlgorithms`) - a click near the vertex glued the new point to the
  mark, and it moved with every zoom. Its grid says the angle in degrees, like the number.
  Its Radius (`AngleArc.DefaultSize`, 20 px since 2026-10-08) is also the side of the
  square it draws at 90° (the square was `RightAngleMark.Size`, the perpendicular's own,
  whatever the radius said). A mark takes a shape style as a circle does (`[StyleFor(typeof(AngleArc))]` on
  `ShapeStyle`, not `IShapeWithInterior`, which the Area tool takes), the default being
  the line: filled, the angle between the vertex and the first arc (the square of a right
  angle) is filled, by a second path under the arcs (`AngleArc.FillShape`: a path's stroke
  outlines every figure in it, and the sector's radii must not be drawn over the sides;
  the arcs' own figures say `IsFilled = false`, or an open figure fills up to its chord),
  which takes the fill the style put on the shape, follows the shape's visibility and
  opacity through its property changes, goes onto the canvas only when there is something
  to fill, and is hit inside (`IsInsideFill`). The hover's angle preview tints the same
  sector (`ClickPreview.CreateAngleFillGhost`, the shared `CreateSectorGeometry`) and
  draws the ghost arc in the preview's blue: faint and thin, the arc alone was missed. The
  Angle Bisector tool adds the bisector it would make, a faint ray from the vertex
  (`Behavior.PreviewsAngleBisector`, `ClickPreview.CreateBisectorGhost`). `DGFReader.ReadMeasureAngle` creates the arc from
  VB6's DrawStyle / AuxInfo(2) - not tested, there is no sample .dgf with an angle in the repo.
- **Dashes**: `LineStyle.Dash` is put on in `LineStyle.OnApplied`, not through a setter, because
  `StrokeDashArray` counts in stroke widths. Anything that
  applies a style to a shape by hand (sample glyphs) must call `OnApplied` too. Any enum
  property of a style or figure round-trips by name through the generic `EnumSerializer`.
- **Z order** (`Figures/ZOrder.cs`, 2026-10-10): every figure is drawn in its kind's
  `Layer` (the `ZOrder` enum, bottom up: Grid, Axes, Polygons - polygons, the interiors of
  regular polygons and Bezier paths - Labels, Figures - lines, circles, arcs, curves,
  the sides of paths - Vectors, SelectionHalos, Handles, Points, PointLabels, Controls), and
  within the band from Polygons to Vectors (`ZOrders.IsMovable`) its `Z` comes first and the
  layer second: `ZIndex` is `ZOrders.Encode(Layer, Z)` - the layers under the band as they
  are, the band above them with a stride of 1000 per Z, the layers over the band over all
  of it. Every figure's Z is 0, so the gallery draws as it did before there was a Z (checked
  pixel for pixel, `--check` before and after), and within one ZIndex the later figure is
  on top (Avalonia sorts children by ZIndex stably; the player sorts the list the same
  way). Bring to front sets the Z to one above the band's highest, Send to back to one
  below its lowest (`ZOrders.BringToFront`/`SendToBack`, a `SetProperty` per figure in one
  transaction: undo puts the old Z back); a polygon brought to front covers a circle, a
  measurement, a segment and its marks, and nothing in the band ever covers a point. The
  verbs are grid buttons on every figure (`FigureBase.BringToFront`, vetoed through
  `[PropertyGridCondition]`, which names a bool method of the object: the grid asks it
  where `IConditionalProperties` would have to be implemented by every figure), on a
  multiple selection (`FigureSelection`, all of them at one Z, their order among
  themselves kept by the list) and in the context menu (`Dragger`, for the selection as
  Hide and Lock take it); offered only when a figure of the band that is not among them
  is over (or under) one of them, or the step would be an undo step that changes
  nothing; a figure outside the band (a point, a slider, a pinned label) gets none. A
  composite's parts take its Z (`CompositeFigure.OnZIndexChanged`, and on `Children` adds:
  a regular polygon's sides, a path's pieces); a composite drawn by its parts alone says
  its layer itself (a regular polygon and a Bezier path are Polygons), or the verbs would
  not apply to it. Whatever a figure draws beside its shape follows it: the angle's fill at
  `Shape.ZIndex - 1`, the right angle mark at the owner's `ZIndex - 1`, a segment's marks
  at its `ZIndex`, the hover halos at the source's `ZIndex - 1`; a preview visual of a
  figure to come uses `ZOrders.Default(layer)`, a bare layer constant is never a ZIndex
  any more. Saved as `Z="1"` on the figure, left out at 0; the player reads it
  (`figureBase.js`, `zOrder.js`). Since the Z is not the list, the list stays the
  dependency order, and a figure may go behind what it is built on.
- **Selection is a halo, never a change of the figure** (`Figures/Shapes/SelectionHalo.cs`,
  2026-10-03): a striped band drawn from the shape's own `RenderedGeometry` under it (a
  slightly bigger silhouette for a point, a disc for an emoji), in one layer per canvas
  (`SelectionHalo.Layer`, `ZOrder.SelectionHalos`, over every figure of the Z band and
  under the points since 2026-10-10: under its figure it was hidden by one brought to
  front) whose opacity the halos share, so overlaps don't darken. Its brush is the theme's `SelectionHalo` (a gradient edited as
  any on the theme page), not stretched over the figure but repeated, reflected, every
  `StripeWidth` pixels along its angle (`SelectionHalo.MakeRepeating`); the halos redraw
  when it is edited or the theme switches. It is colorless on purpose (grays: white to
  #B0B0B0 in Light, #B0B0B0 to black in Dark): a hue next to the figure changes how the
  figure's own colors read, and those are what is being edited while it is selected. `ShapeBase.UpdateSelectionHalo` keeps it (selected, shown, on a
  canvas); it redraws on any property change of its shape (its bounds included), a change
  of a path's geometry in place (`Geometry.Changed`) or of a polygon's points
  (`IChangesPointsInPlace`), so nothing per figure kind - not on `LayoutUpdated`, which
  Avalonia raises for every listener after any layout pass anywhere. It draws a copy of the
  shape's geometry: a geometry is a compositor resource, and shared with the shape, a
  Bezier curve changed while selected vanished when it was unselected. Selected figures used to be drawn 3 px thicker or bigger and polygons got a striped
  fill, which hid the very thickness, size and fill being edited in the side panel: a style
  must not depend on `Selected`. A part never drawn says `ShowsSelectionHalo = false`.
  Labels keep their own selection plate.
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
  selected tab blends into it. A click on a swatch gives the page the keyboard, and the
  arrow keys move the current swatch through the grid and pick as a click does
  (`SwatchPage.OnKeyDown`; one undo step for the run, the canvas doesn't pan meanwhile).
  On the Spectrum tab the surface pressed last (field, hue, opacity) takes them as fine
  steps, 1% or 1°, ten with Shift (`DragSurface.Stepped`); a step of 1/255 left the color
  as it was once rounded, and a press that changes nothing looks broken.
- **The paper** is `Drawing.Background`, edited through "Drawing background" on the settings
  page (`AppSettings.EditDrawingBackground`; `DrawingHost.ShowDrawingProperties` puts the
  drawing itself in the property grid). "Reset to default" (the theme's paper) shows only while
  the paper is not the theme's, and "Same paper as in Light" only while there is a Dark override: both come and go
  as the paper changes (`[PropertyGridLiveCondition]`: the button is made hidden and the grid
  asks `CanEdit` again on each named `PropertyChanged`; the divider over destructive buttons
  hides with them), since the usual rebuild on a change
  without a name would fold the color picker under the cursor. A gallery tile takes a drawing's paper as its plate and turns its caption white on a
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
  of a segment or a vector still makes a midpoint (`PointPlacement.HasMidpoint`: a vector
  is a figure of its own around a hidden segment, and a check for `Segment` missed it). In a narrow window (a phone, under about 410 px)
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
  `AreaFill`/`AreaHatch`, `LineAccent` for the line or curve a
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
  all (PrintWindow): to check it, `shot ... --screen`. On a Mac Avalonia shades the title bar
  itself.
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
  drawing was behind the gallery), and then not saved. The palette's colored styles (see
  "The default styles") are literal, with a literal Dark override each. A drawing's paper is
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
  `Drawing.Overrides` holds the paper chosen for another theme, "Reset to default" under Dark
  stores a null override (that theme's paper). Both buttons are undo steps.
  `GalleryTitle` (a Dark override in the splash's blue), `GalleryText` (the chrome's text
  color) and `GalleryLocus` are defaults too. The gallery drawings' own styles got their
  Dark overrides from `dotnet tools/darken.cs -- <folder> [--apply]` (2026-09-29): a stroke,
  text or fill darker than lightness 0.35 is lightened to the same hue at 0.82 - 0.4 ×
  lightness (black lands on the ink), a drawing with a paper of its own is left alone, and a
  style that has a `<Dark>` already is never touched, so it is safe to rerun after adding a
  drawing. The three drawings with a light paper of their own (Castle, Conic, Rose) carry
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
  (`~/Library/Application Support/LiveGeometry/Settings.txt` on a Mac;
  `FileSettingsStore`), the browser in `localStorage` under `LiveGeometry.<key>`
  (`BrowserSettingsStore`, two imports in `main.js`; `index.html` reads the `Theme` entry
  before the runtime starts, so the splash and the page are already dark). `AppSettings`
  (`[PropertyGridName("Settings")]`, the gear) is the page over it: `Theme` is `System` or a
  theme's name, applied in `App.Initialize` before the first frame. The window placement is
  the `WindowPlacement` line of the same file (`WindowBounds` on a Mac). Colors tweaked on the "Theme colors" page are
  kept per theme (`ThemeColors.Dark`), only those that differ from `AppTheme.cs`
  (`AppTheme.EditsToText`/`ApplyEdits`, one line, read without throwing), so an untouched
  color follows the code; they are put on before `AppTheme.Register`, and "Built-in colors"
  drops them. Writes wait for a half-second pause (a picker drag sets a color per move) or
  `SettingsStore.Leaving` (the window closing; the page hidden, from `main.js`'s
  `visibilitychange`). In the browser only (`MainView.KeepsOwnDrawing`), the user's own
  drawing is kept too (`MainView.KeptDrawing.cs`: keys `Drawing`, `DrawingName`): saved
  a second after the last change of the undo history (`Throttle`, which only posts the
  save to the UI thread) and on Leaving, never during a construction, removed when empty,
  and not written again when the text is what the store holds (saved at the next idle
  moment, it was saved at every move of a figure's drag, whose moves are merged steps, and
  the drag went a few frames a second);
  read at startup but loaded only when the user goes to it (/drawing, My Drawing). Until
  then it is their drawing and nothing is written over it; New or an opened file replaces it.
- **Ribbon look**: `ButtonGrid` draws the
  hover/pressed/checked plate; `Ribbon`/`TabPanel` replace the Fluent templates in code. To
  bring the group headers closer together, change `ButtonGrid.HeaderOverlap`, not the padding
  inside the tab. An on/off `Command` exposes `IsChecked` (a `Func<bool>`), which its button
  re-reads after any toggle (`CommandToolButton.UpdateToggles`, called by anything that toggles
  from a key) - don't go back to `CheckBox` icons.
- **The splash** (`wwwroot/index.html` + `app.css`) stays up until the app says it has drawn
  what the page opened at (`SplashScreen`, 2026-10-07): the gallery with its tiles in view
  loaded (`GalleryView.LoadNextTiles`) or a drawing laid out (`MainView.HideSplashWhenDrawn`);
  `main.js` adds `app-ready` then and the splash fades. Avalonia's own `splash-close`, added
  at its first frame, hides nothing any more: on the gallery page that frame came seconds
  before the tiles, which then popped in batch by batch. Two deadlines in `main.js` take the
  splash down should the app never say so: 10 s after Avalonia's first frame, 20 s after Main
  returns. It animates Euclid's first construction in HTML, every move a transform or an
  opacity (a ring is drawn by turning a half-colored ring behind a half-window, a line is ink
  sliding into a slot), which the browser plays on its compositor thread while the main
  thread is held by the app's startup; the SVG version's `stroke-dashoffset` and
  `offset-distance` were painted on the main thread and froze with it. The bar under it is
  the downloads for its first 60% (a `withResourceLoader` wrapper counting fetches: the
  loader's `onDownloadResourceProgress` is on the module config, which the .NET 10 host
  builder doesn't expose) and what the app reports for the rest (`SplashScreen.ReportProgress`:
  tiles in view loaded, out of all in view); it is a cover sliding off the gradient, not a
  width, for the same reason. The splash takes the pointer events while it is up (the app
  under it got the clicks). The wrapper hands back a Response only for assemblies, the wasm
  and ICU data; the runtime's own JavaScript modules and its config must be left to the
  default loading (a Response for `dotnetjs` leaves the site spinning on the splash forever).
  A change to `main.js` needs the publish smoke test below before deploying. `?splash` on
  the url shows the splash without starting the app: serve the source `wwwroot` with
  `tools/serve.cs` and open `http://localhost:<port>/index.html?splash`. It animates
  regardless of `prefers-reduced-motion` on purpose: the query follows the Windows "Animation
  effects" setting, which many machines have off, and a still splash looks stuck.
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
  `Line`, `SliderTrack`, `Shape`, `OutlinedShape`, `Text`, `Heading`, `Hyperlink`, and the palette ones
  named by hue: `RedLine`, `ThickRedLine`, `DashedRedLine`, `RedOutline`, `RedShape`,
  `RedPoint`, `RedBead`... - see "The default styles"; `OtherLine`, `DottedLine` and
  `OtherShape` are gone). Saved are the styles the figures name (`DrawingSerializer.Write` writes the
  figures aside first and collects their `Style` attributes - a figure may name another's
  style, a vector its arrow's) and a default the drawing changed; a default as a new drawing
  has it is left out (`StyleManager.IsUnchangedDefault`). `Name` comes first, and an
  attribute at the value a fresh style has (`IsFilled="true"`, `Dash="Solid"`) is left out; a
  missing one reads as that value. A point style's `Size` is compared with a fresh style
  showing the same character (`DrawingSerializer.FreshStyleFor`): a character's default is
  24, a shape's 10, and an emoji of size 10 was left out and came back at 24. What differs
  under another theme is a child element
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
  The order matters: it is the style picker's, and a figure that isn't a point or a slider
  takes `Line` or `Shape` by name, else the first style that fits
  (`StyleManager.GetDefaultStyle`) - but a sector or a circular segment, an arc with an
  inside that draws its own outline (`StyleManager.IsOutlinedShape`), takes `OutlinedShape`
  (2026-10-09: both a line and a shape, it took `Line`, which comes first, and had no
  fill; a circle keeps the line, a polygon's outline is its side segments). Labels leave
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
  for a message. A figure that says too little for its kind - a segment on one point, a
  perpendicular to a point - is caught where it asks (`FigureBase.Point(index)`,
  `IFigureExtensions.Point`/`Line`, 2026-10-09): while the file is read the figure is noted
  (`Drawing.InvalidFigures`), made not to exist (its shape must not be laid out at
  infinity) and given a point that is nowhere, and the deserializer leaves it out with
  what is built on it once every figure is in (`LeaveOutInvalidFigures`); at any other
  time the same access is a bug and throws. A direct `Dependencies[i]` cast is not
  covered. Such a drawing keeps the list (`Drawing.LoadErrors`), gets no name and
  so no file to be saved over, and the status says the first line and how many more.
  `IniFile` skips the lines of a `.dgf` it can't read. A viewport without a size (a file
  saved from a window that had none) keeps the view (`CoordinateSystem.SetViewport`): it
  threw, and the file was refused whole.
- **Label text** is one attribute: a line break is the two characters `\n` and a
  backslash is two backslashes (a typed `C:\notes` came back as two lines); characters an
  XML file can't hold are left out (one pasted in made Save throw and write nothing).
- **A show/hide box is read as a box** (`ShowHideControl.ReadXml`): each figure says in
  the file whether it is hidden. Applying the box to its figures on load hid again one
  that had been shown by hand since the box was last clicked. (The `.dgf` reader does
  apply it: those files have only the buttons.) The box exists whatever its figures do
  (`ShowHideControl.UpdateExistence`): built on them like any figure, it vanished while one
  of them didn't exist (a firework that hadn't burst yet). Its empty box takes the
  caption's color, as the caption does: Fluent's outline was dark on a drawing's own dark
  paper.
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
  polygon was added, they failed the consistency check (`CheckConsistency`), and such a
  file did not load. When the point a polygon is built on is
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
  A vertex or a side is selected by itself (`IFigureParts.SelectableParts`), to be styled;
  a click inside selects the polygon (`FigureParts.SelectionTarget`), and so do its Figure
  List row and the part page's "Select regular hexagon p". The composite's `Selected` is the whole
  only (selecting the whole selects every part, for the halos), so whatever clears a
  selection asks `FigureParts.HasSelection`, `Drawing.GetSelectedFigures` gives the parts
  selected by themselves, and whatever acts on figures of the drawing (Delete, Copy, Hide,
  Lock, the context menu) takes `FigureParts.Wholes` of it: Delete with a side selected
  deletes the polygon. A part's page is its `ICustomPropertyProvider` (style, a side's
  length, which is the polygon's), not the rows of a point or a segment. The polygon's page
  has Fill (its own style), and Sides and Vertices (`PartStylesValue`: every part at once,
  undo giving each its own back). The file says the style most sides have and the odd ones
  (`<Sides Style="RedLine" />`, `<Vertices ... />`, `<Part Name="Side2" Style="BlueLine" />`,
  child elements, so that the `Style` attributes are found when styles are collected); a
  new side takes the style most have, one that comes back keeps its own (undo of fewer
  sides).
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
  place and in the page title - nothing for a new drawing). The tour group follows the
  buttons behind a separator (`MainToolbar.IsGroupLeftAligned`; off, it is centered in the
  room they leave, and the code for that stays). Paths `/`,
  `/gallery/<slug>`, `/drawing`; `AddressBar` is the abstraction, `BrowserAddressBar` +
  `main.js` do pushState/popstate. `index.html` needs `<base href="/">` for that (all routes
  serve it; `web.config` and `tools/serve.cs` fall back to it). All page changes go through
  `MainView.Show*`. Started at a drawing's address the gallery isn't built at all (the address
  is readable synchronously since `main.js` registers its imports before the runtime starts),
  and started at the gallery the editor isn't (`MainView.DrawingHost` builds the drawing
  host, the ribbon's tools and the toolbar the first time anything asks for it; what can run
  on the gallery page - keys, presses, resizes, error reports - asks `IsEditorBuilt` first).
  Built up front it was the most of what came before the gallery's first frame.
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
  share (under the figure, the text may leave it a third of the height; beside it, 40% of
  the width). If the text doesn't fit it runs off the bottom and the reader drags it up: the text
  labels a gallery drawing comes with (`Label`: the caption, text in the plane) can't be
  dragged (`Drawing.FixedLabels`, not saved, emptied once the drawing is the user's own; the
  labels of figures - point names, measurements - and labels added later can: all of them
  were paper at first, and a name or an angle's number could not be moved off a figure
  dragged under it), a drag on one pans the view, and one that starts on the caption scrolls the
  pinned labels along (`PinnedLabelScroll`, by pixels, so undo brings both back). A click
  still selects one (with no way to its properties but a right click, and none at all to move
  it, they were out of reach), and a selected one drags as any label does - but the caption:
  on a phone, a tap and then a thumb to read on would pull the explanation away from its
  heading. Nor does a finger's tap select the caption (the side panel would cover half the
  phone for a reader touching the text). The context menu says "Unlock" for one, which takes it out of the set, caption
  included. Not
  `Locked`: a point counts as locked when anything built on it is, captions included. Refitted on
  resize until the first edit, through `Drawing.SizeChanged` (`MainView.KeepFitted`) and not the
  canvas's event: the coordinate system's own resize handler shifts the origin and has to run
  first. A file with a caption opened from disk is fitted the same way. Explanations have no
  hand line breaks (a blank line separates paragraphs). On an iPhone in portrait (canvas about
  390x565) the long explanations run off the bottom. A third label, `Hint`, hidden and
  listed by a show/hide box, goes under the explanation a paragraph down, its room kept
  while it is hidden (`LabelBase.MeasureSize` measures a hidden label as shown: Avalonia
  measures it as nothing). Under the caption the figure keeps at least a third of the
  height; a catalog item can ask for more (`stackedFigureShare`: the emoji drawings, Steiner's,
  the pentagon), and a second, narrower scene crops a picture at the phone's sides (Treasure
  Island, Magic Tree, Fireworks; full screen picks the wide one). A point's own size limits
  the zoom only where the point is (`CoordinateSystem.LimitZoomByPointReach`): shrinking the
  room on every side by the biggest emoji (Kaleidoscope's blossom in the middle) left the
  figure far smaller than the room.
- **Tiles are live drawings**, not bitmaps (`DrawingThumbnail`): no Behavior, text hidden;
  hovering makes the draggable points drift (the visible ones; where there are none, the
  hidden ones that something on screen is built on - Bubbles shows its points only with its
  hint). Loaded a batch per idle tick and only as they come
  near the view (`GalleryView.LoadNextTiles`: those in view from the top, then half a screen
  around it; none while the page is hidden behind the editor; one a turn while the splash is
  up, so that its bar moves after each): a load holds the UI thread, in the browser a few
  hundred milliseconds for a drawing with hundreds of points by coordinates. Which tiles are
  near the view comes from the grid's arithmetic (`TileGridPanel.GetTileRect`), not from
  asking each tile where it is: that ran at every step of a scroll, and `TranslatePoint` over
  sixty tiles took 30 ms a step on a phone - most of what made scrolling the gallery choppy
  (2026-10-07; `tools/scrollperf.cs` measures it). Tried then and dropped: Avalonia 12's
  `BitmapCache` (`Visual.CacheMode`) on the tiles near the view - in the browser a tile was
  drawn into its bitmap again and again during a scroll, 50-100 ms a time at a 4x CPU
  throttle, and scrolling got worse, not better. A tile whose points are characters waits
  for the emoji font while the others load (`GalleryItem.UsesEmoji`); loaded before it, it
  showed its emoji seconds later. A load error of a tile goes to the console as
  `Gallery: <file>: ...`; every drawing at once is loaded by `tools/regression.cs`.
- **Generated drawings** - regenerate rather than edit the file: Line of Best Fit
  (`dotnet tools/bestfit.cs -- <the .lgf>`), Fibonacci Spiral (`tools/fibonacci.cs`),
  The Five Platonic Solids (`tools/platonic.cs`, `--alpha` for a translucent variant),
  Aperiodic Monotile, Eight Kites Make a Hat and The Hat Family (`tools/hat.cs -- <the
  Drawings folder>` writes all three: it searches the kite grid for a gap-free patch of hats
  around the middle one, the seed picks the patch), Seventeen Squares (`tools/squares17.cs`,
  then `--rewrite` the folder: Bidwell's packing, centers and angles from the Squares
  project's witness file). Captions of these live in their tools
  too: change both. A segment is drawn over every polygon whatever the order of the list
  (`ZOrder`; only a `Z` changed by Bring to front or Send to back is saved, see "Z order"),
  which is why the kites drawing's grid is polygons. Hidden
  `PointByCoordinates` whose coordinates are expressions are the library's variables.
  The Bézier path drawings (2026-10-07) are generated too, each by `dotnet tools/<tool>.cs
  -- <the .lgf>` and a `--rewrite` of the folder, which gives the file byte for byte:
  Stretchy Slime (`slime.cs`), Moon Jelly (`moonjelly.cs`), Poke the Blob (`pokeblob.cs`),
  Pump Up the Balloon (`balloon.cs`), Connect the Dots (`connectdots.cs`). All but the last
  are one recipe: paths with automatic handles whose anchors are hidden points by
  coordinates over a few hidden "variable" points (two numbers each) worked out from what
  is dragged, so a drag moves the anchors and every path smooths itself again; a tension
  tied to a hidden label `[...]`; and whatever appears past a threshold (the pop, the drop,
  the swallowed bubble) is built on a point `X="0 * sqrt(v - limit)"`, which doesn't exist
  below it, and neither does anything built on it. Design a Font (`designafont.cs`,
  2026-10-10) is the other kind: the big t, h and e are closed paths with explicit handles
  over free points, and every t, h and e of the sentence is a dilated image of the big one
  (a `DilatedPoint` per anchor and per handle, the handles as `Part="In3"` dependencies on
  the big path, one hidden center per copy chosen so that the letter's box lands in its
  cell, `Shrink` = 1/8), so dragging an anchor or an arm of the big letter moves the small
  ones. The other letters are a tiny monoline font the tool strokes from skeletons of lines
  and tangent arcs (round caps, 180° as one cubic) into paths over hidden points by
  coordinates; a bowl is a ring with a hole. 726 figures, the biggest drawing of the gallery.
- **The eye-candy batch** (2026-10-10, 19 drawings; the brainstorm they came from and
  what is still open is `docs/gallery-ideas.md`). Generated, each by `dotnet
  tools/<tool>.cs -- <the .lgf>` and a `--rewrite`: Times Tables on a Circle
  (`timestables.cs`), Gears (`gears.cs`: trapezoid teeth as polygons of points by
  coordinates turned by a slider, the phases worked out so that a tooth meets a gap),
  Make Your Own Tessellation (`tessellation.cs`: one closed path whose bottom and left
  edges are its top and right edges shifted by a side, every handle a point so that the
  eleven translated copies can follow), Design a Snowflake (`snowflake.cs`: one twelfth as
  a path with automatic handles, mirrored and rotated), Tangram (`tangram.cs`: a piece is
  a polygon on a free pivot and a point on a unit circle around it, which turns it),
  Spiral of Theodorus (`theodorus.cs`), Spin the Die (`spindie.cs`: 3D points by
  coordinates over two angle sliders, translucent faces so that nothing needs sorting),
  Fold the Cube (`cubenet.cs`), Flower of Life (`flower.cs`), Slice the Pizza (`pizza.cs`:
  sectors moved between two places by a slider), Clock Hands (`clock.cs`: the hands are
  polygons of rotated points, the second hand sweeps once per minute of clock time; the
  frame, bezel and face are 72-gons, not circles, because a polygon draws in the layer
  under circles, `ZOrder.Polygons` under `Figures`, and a filled circle hid the hands;
  since "Z order" (2026-10-10) a circle with `Z="-1"` would do, the polygons stay).
  Hand-written: Cycloid (the inner and the
  flange point slide on a hidden ray along the spoke), Spirograph (the pen slides on a
  hidden ray that turns with the wheel; six loci a turn apart, since a point on a circle
  goes round once), Two Pins and a String (the pencil is a point on the locus, dragged
  along it), A Straight Line from Circles (`Peaucellier.lgf`: the left nail is free, the
  crank's center is a radius slider away from it, the three lengths are sliders styled as
  their rods, and the joint is a point on an arc that covers only the reach of the rods),
  Billiards (the bouncing path is a locus of a point folded with
  `abs` and `floor`), Three-Point Perspective (the third point sits below a `<Scene>`
  that keeps it out of the fit), What Is π? (a wheel 1 across rolls and paints its rim
  onto a ruler: a slider 0 to π, the painted part straight segments clamped to the slider,
  the rest arcs of the wheel), Don't Trust Your Eyes. The sliders' `Maximum` came from
  this batch. Learned on the way: an `AngleArc` and an
  `AngleMeasurement` list the vertex first; a label (`ZOrder.Labels`) draws under a filled
  circle (`ZOrder.Figures`), so a measurement over a clock face is invisible unless it is
  brought to front (`Z="1"`, see "Z order") -
  put the number in a label outside the shape or in the caption; a label as a rotation's
  angle is read in radians (`[rad(-30 * time)]`); the `Parameter` of a point on a line is
  the fraction from the line's first point to its second; a name's trailing digits draw
  as a subscript, so the Theodorus labels are "√17 " with a space.
- **Drawings that must not fall apart** when a kid drags the wrong thing (Castle, The Falling
  Ladder): fixed points are `PointByCoordinates` with constant coordinates (a polygon of those
  has no free point, so dragging it does nothing), and the only things that move are
  `PointOnFigure` sliders on hidden rays or segments plus a free point or two.
- **Emoji drawings** (Kaleidoscope, Golden Angle, Treasure Island, Magic Tree, Fireworks,
  2026-10-02): an emoji keeps its size in pixels while the figure is fitted to the window,
  so they are drawn a few units across: zoom to fit stops at
  `CoordinateSystem.MaxFitUnitLength` (200 px a unit), and on a big screen the picture
  stops growing instead of leaving its emoji far apart. On a phone in portrait they crowd
  unless the figure gets the phone's width (see "The caption": `stackedFigureShare`, a
  narrower scene). Dragging a rotated or reflected copy moves its
  source by the cursor's offset, so the copy goes the mirrored way: the kaleidoscope marks
  the slice whose emoji are the real ones.
- **The Spiral's rings** ("Drag to here") are placed for its locus's `Samples="60"`;
  changing the number moves them, and without it the curve is a smooth spiral. The samples are taken by their count (61 points, 60
  equal steps of the sliding point's parameter): added up, the step's rounding decided
  whether the last sample but one was taken, and the curve had 60 or 61 points from one
  move to the next.

## The JavaScript player

`Main/Player/` (2026-10-07; the plan it follows is `docs/player-plan.md`): a drawing on any
web page, in plain JavaScript with no dependencies, a canvas with the Drag tool and nothing
else - no ribbon, grid, list or gallery, no undo, no selection. The Silverlight-era
`PLAYER`/`TABULA` defines in `DynamicGeometry` are not it (they are leftovers to delete).

- **A port of the library, file by file**: `src/` mirrors `DynamicGeometry/` (`math.js`,
  `drawing.js`, `figures/points/midPoint.js`, `expressions/parser/scanner.js`,
  `serialization/drawingDeserializer.js`, `styles/styleManager.js`, `behaviors/dragger.js`),
  the C# names in camelCase, so that a reader of `MidPoint.cs` finds its twin and a change
  to one wants the same change in the other. **The rule: a new figure kind, a new function
  of the expression language, a new file attribute, a changed `Recalculate`, `HitTest`,
  `ReadXml` or sampler goes into the player in the same change** (a kind the player can't
  have is said so in its deserializer's error, "this version has no figure of the kind").
  What has no C# twin sits apart: `src/render/` (the canvas renderer, `TextMeasurer`) and
  `src/host/` (`Player`, `PlayerCanvas`, the `LiveGeometry` global). Left out on purpose:
  tools, undo, selection, snapping and joining, Convert verbs, layout changes of a Bezier
  path (anchors in or out, holes cut), transformations' images, Hyperlink, user tools.
- **JavaScript conventions**: an interface is a getter that answers true (`isPoint`,
  `isLine`, `isLinearFigure`, `isCircle`, `isLengthProvider`, `isAngleProvider`,
  `isNumber`, `isMovable`, `isMovableParts`, `isFigurePart`, `isBezierPathPiece`...), and
  a figure's own flag must not take such a name (an `isLine` field on the bisector broke
  every line: it is `wholeLine`). Classes that would shadow a browser global are renamed:
  `GeometryMath` (Math.cs), `NumberFigure` (Number.cs; the element is still `<Number>`),
  `SyntaxNode` (Node.cs). Every figure file ends with `FigureTypes.register("<element
  name>", Class)`. A figure has `render(renderer)` besides `updateVisual`; `resolvedStyle`
  is the style resolved for the theme on screen (what the Avalonia shape had applied);
  `stroke`/`fill` are what the renderer takes. Classic scripts, no modules, no bundler:
  `files.txt` is the load order (bases before derived classes), and a file referenced only
  at run time (`Dragger` from a figure's `hitTest`) may come later in it. The Write tool
  emits LF: new files are converted to CRLF afterwards, like the C#.
- **The bundle**: `Player.targets`, imported by the Browser and the Desktop projects,
  concatenates `files.txt` into `obj/player/<version>/player.js` before the build (an
  inline task, dotnet only), wrapped in a function so that the host page sees one global,
  `LiveGeometry` (`play(element, options)`, `playAll`, `players`, and the classes for a
  page that builds on it). The Browser serves it as `/player/1/player.js` and every
  gallery drawing as `/gallery/<file>.lgf` (both `no-cache` and with
  `Access-Control-Allow-Origin: *` in `web.config`, since a page on another site fetches
  the `.lgf` and the browser blocks that otherwise; a script tag needs no CORS); the
  Desktop copies it to `player/player.js` beside the exe. **Versioning**: the folder is the
  player's major version. A snippet pasted into a post pins `/player/1/`; when the format
  breaks that player, `/player/2/` starts and the last `/player/1/player.js` is committed
  as a static file and served forever. Within a version the file revalidates, so fixes
  reach old embeds.
- **Embedding** (`PlayerEmbed.cs`; Export > "Copy embed code" and "Save as .html"): a
  script tag plus `<div class="livegeometry">` with the drawing's XML inside a `<script
  type="text/x-livegeometry">` (inline, works from a disk too; the XML can't contain
  `</script`, since it escapes `<`), or `data-src="<url of an .lgf>"` (the file's server
  must allow CORS; a page opened from `file://` can't fetch at all). Options: `data-theme`
  (light, dark, auto), `data-font`, `data-wheel="zoom"` (else the wheel zooms only with
  Ctrl, so an embed doesn't take a blog's scrolling), `data-fit="content"`. "Save as .html"
  writes the player into the page. Fonts: the host page's, and the browser's own emoji
  font (none shipped; a character the system lacks is a box). The page that explains all
  of this to a reader is `/embed` (see "Deployment and caching"): a change to the options
  or the snippet wants the same change there.
- **Behavior that differs from the app, on purpose**: no selection, so a Bezier path's
  handles show from a press on an anchor until the next press elsewhere
  (`BezierPath.showHandlesWhileDragging` from `Dragger.mouseDown`), and the Tab choice of
  a handle under its anchor doesn't exist. A drawing with a caption (Title and Description
  labels: the gallery's, an export of one) is laid out for the element as the app lays it
  out for its window, by a port of the app's `GalleryDrawing.Fit` (`src/host/galleryDrawing.js`,
  the one twin of a file outside the library), now and on every zoom to fit; an export
  carries the pin and offsets it had on screen, which a small element would clip. Any
  other drawing gets the file's viewport (or its scene) fitted into the element.
- **Working on it**: `dotnet run tools/serve.cs -- Main 5005` (`.claude/launch.json` has it
  as "player-dev") serves the repo's `Main` folder, and
  `http://localhost:5005/Player/dev/index.html#<Drawing>` loads `files.txt` one script at a
  time (`?v=` cache-busted; to reload after an edit change the query, `?r=2`, a changed hash
  alone reloads nothing) with a list of the gallery drawings, Fit, dark and wheel switches,
  and a status line with the figure count, the load time and the load errors.
  `Main/Player/dev/kinds.lgf` has the kinds no gallery drawing uses (a regular polygon with
  parts, a polyline, lines and a circle by equation, a line at an angle, segment marks, a
  name label, a sector, an ellipse arc and segment, a horizontal angle). To check every
  drawing at once, in the page's console: for each file of the listing, `new
  Drawing(player.canvas).addFromXml(text)` and look at `loadErrors` and the visible figures
  that don't exist (Fireworks, Pump Up the Balloon and Poke the Blob have some by design).
  For pictures use headless Edge (`webauto start <url> 1000 700`, `nav`, `shot`, `drag`,
  `eval`): the app's browser pane can be too small to show a drawing, and `--check`'s PNGs
  of the desktop are the reference to put beside them. Headless Edge has a pixel ratio of
  1; `Player.pixelRatio` says what the canvas is scaled by. `Main/Player/dev/bundle.html`
  plays the bundle the Desktop build made, as a page on another site would: an embed by
  URL and an inline one (dark, wheel zooming); `hosted.html?host=http://localhost:5006`
  takes the player and a drawing from another origin (a publish served by `tools/serve.cs`
  on that port, or `https://livegeometry.com` after a deploy: the cross-site fetch for real).
- **The check**: `dotnet run tools/playerparity.cs` (a project reference to the desktop
  head, like `regression.cs`): every figure kind `DrawingDeserializer.FigureTypes` reads is
  registered in the player or in `Main/Player/excluded.txt` with a reason, every function
  of `Functions` has a twin in `functions.js`; then every gallery drawing is loaded by both
  (the library in-process, the player in headless Edge on the dev page, started if they
  aren't running), dumped the same way (`Drawing.dump`: the figures in the list's order
  with the numbers that define them, points, lines, ellipses, polygon vertices, label
  texts, numbers, rounded to 9 decimals) and compared at 1e-6, once as loaded and once
  after every free point is moved by (0.3, 0.2). Left out of the numbers: what is sized in
  pixels (an angle's mark) and what is sampled by the window (a locus, a graph; the points
  on them are compared). Run it after a change to a figure's numbers on either side.

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
  `/history/tabula` (2026-10-09) is a wing of it, `history/tabula.html`, with its pictures
  in `history/tabula/` (a `web.config` rule answers the path with the page, and `serve.cs`
  serves `<path>.html` for an extensionless path that has one); it references them by
  absolute path, since a relative one would resolve against `/history/`. Everything on it
  came from the Wayback Machine's captures of numeracyworks.com and from YouTube.
- **`/embed`** is the same kind of page (`embed/index.html` at the repo root, linked into
  `wwwroot/embed/`, a `web.config` rule): the hub for putting a drawing on another page,
  itself played by the player it describes (`/player/1/player.js`, the gallery's `.lgf`
  files). Its picker lists the gallery from `/gallery/index.json`, which the Browser csproj
  writes at build from `GalleryCatalog.cs` (`WriteGalleryIndex`: the `Item("slug",
  "Title"[, "File"])` lines, in order), so the list never has to be kept in step by hand.
  Its sun and moon read and write the app's `LiveGeometry.Theme` entry, so the choice is
  one for the site, and set every player on the page (`applyTheme`). To work on it it
  needs the player and the drawings at those absolute paths: serve a publish's `wwwroot`
  with the repo's folder overlaid and live reload,
  `dotnet run tools/serve.cs -- <publish>\wwwroot 5006 --overlay embed=embed --reload`,
  and open `/embed/index.html`; every save of the page reloads the browser. A change to
  the player or the catalog still wants a new publish (or its bundle copied into the
  publish's `player/1/`).
- **`/web`** (2026-10-08) is the gallery and its drawings without .NET: `web/index.html` at
  the repo root (linked into `wwwroot/web/`), played by the player. `/web` is the gallery
  page as the app draws it (the New and Open tiles, which lead to the editor at
  `/drawing`, the brand, the tiles: live drawings in a `TilePlayer`, a `Player` whose
  canvas is 560 x 380 scaled down through the pixel ratio, the caption left out of the
  XML as `DrawingThumbnail.LeaveOutCaption` does, drifting on hover, loaded by distance
  from the view a few per frame as `GalleryView.LoadNextTiles` does); `/web/<slug>` is
  the app's toolbar over the drawing - the mark (inert), Gallery, New and Open (links to
  `/drawing`), the tour group (Page Up/Down), the sun and the Octocat - and nothing the
  player can't do (no ribbon, Save, Export, Undo, Settings, status bar). Routes are the
  page's own (`pushState`); `web.config` answers every extensionless path under `/web`
  with the file, and `tools/serve.cs` answers an extensionless path under any folder with
  an `index.html` with that file. The catalog comes from `/gallery/index.json`, which now
  also carries `stackedFigureShare`, passed to the player as an option. The player exports
  `AppTheme`, `GalleryDrawing`, `Label`, `ControlBase` and `LabelPin` for it. The page
  logs its timings to the console (`web: ...` lines: a drawing's fetch and load, each
  tile, and when the tiles in view are in), which is what it is for: a feel for the
  speed without .NET before deciding whether the gallery should move. To work on it:
  `dotnet run tools/serve.cs -- <publish>\wwwroot 5007 --overlay web=web`, with the
  current bundle and `index.json` copied into the publish (`dotnet msbuild
  LiveGeometry.Browser.csproj "-t:BuildPlayer;WriteGalleryIndex"` makes them in `obj/`),
  then `webauto start http://localhost:5007/web/`.

## macOS

The desktop head builds and runs on a Mac (Apple silicon, .NET 10 SDK, no workloads) from the
same project; the regression suite, `--check` and `--rewrite` work there too (the gallery
rewrites byte for byte). What is the Mac's own (2026-10-02):
- **Ctrl+click is a right click** (`Behavior.IsSecondaryClick`): AppKit hands it over as a left
  press with Control, and Avalonia passes it on as one. Its moves and release are swallowed.
  Ctrl's other uses are Cmd's there, as in the browser on a Mac.
- **The menu bar**: `App.Initialize` sets the application's `Name` (it said "Avalonia
  Application") and its own `NativeMenu` ("Live Geometry on GitHub"; Avalonia's default has
  "About Avalonia"); Avalonia adds Services, Hide and Quit after it. Cmd+Q goes through the
  window's Closing, so settings and the window bounds are saved.
- **The Dock icon** (`LiveGeometry.Desktop/MacDockIcon.cs`): Avalonia ignores `Window.Icon`
  on a Mac and a bare executable shows a blank document, so the vector `AppIcon` is rendered
  and given to `NSApplication.applicationIconImage` through the Objective-C runtime.
- **Window bounds** (`WindowBoundsPersistence`, the `WindowBounds` setting): Avalonia's own
  position and client size, those of the last moment the window was normal. A zoom animates
  through a few sizes while the window still says Normal, so bounds count once they have held
  for a second. The green button (Option+click: Zoom) maximizes, but in Avalonia 12.1 it
  doesn't un-zoom a maximized window, also one zoomed by hand: drag its edge instead.
- `.lgf` files are written with CRLF on every system (`DrawingSerializer.XmlSettings`), as
  the gallery's are and the batch rewrite does.
- Testing traps: Cmd+A in the save panel's name box selects the name without its extension,
  so typing `x.lgf` gives `x.lgf.lgf` (the panel, not the app). `winauto.cs` and
  `contactsheet.cs` are Windows only (WinForms); `macauto.cs` is the Mac's (below).

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

## Measuring speed

- The browser ran .NET interpreted until 2026-10-07 (see AOT below; the interpreter is one
  property away): the same code took 3-4 times as long there as in the Debug desktop build
  (2026-10-03, reading the gallery's drawings), and a loop that costs nothing on the desktop
  can take a second there. Measure in the browser before deciding what is slow. Reflection is the worst of it: `GetAttribute` (`Utilities`) and
  the property lists of `PropertyDiscoveryStrategy` are cached per member and type, and a
  save compares the drawing's styles with default styles made once per `AppTheme.Version`
  (`StyleManager.IsUnchangedDefault`). Before that (2026-10-04) a save of a triangle took
  5 ms on the desktop, nine tenths of it in the styles.
- CPU samples of the desktop app need neither admin rights nor a tool: start it with
  `DOTNET_EnableEventPipe=1`, `DOTNET_EventPipeOutputPath=<file>.nettrace` and
  `DOTNET_EventPipeConfig=Microsoft-DotNETCore-SampleProfiler:0:5,Microsoft-Windows-DotNETRuntime:0x4c14fccbd:5`
  in its environment, close it normally, and read the stacks with the TraceEvent package
  (`TraceLog.CreateFromEventPipeDataFile`, then `CallStack()` of each `Thread/Sample` event).
- Better, with an elevated shell (ETW needs admin on Windows; not on a Mac): the Ultra
  profiler (`dotnet tool install -g Ultra`) and the UltraMcp server (`dotnet tool install -g
  ultramcp`, an MCP server that reads its traces). Build Release, then
  `ultra profile --duration 15 -o <name> -- <...>/bin/Release/net10.0/LiveGeometry.Desktop.exe [--gallery <slug>]`
  in a scratch folder: it samples every thread 8190 times a second (kernel, native and
  managed stacks, JIT and GC markers) and writes `<name>.json.gz` in the Firefox Profiler
  format; the app stays open when the time is up (close it). The MCP tools read the `.gz`
  as it is: `list_threads` (thread 0 is the UI thread), `call_tree` with `focus` and
  `startMs`/`endMs`, `call_tree_inverted` for who calls a hot function, `find_hotspots` to
  compare named functions between two traces of the same scenario. A sample is 0.122 ms of
  CPU. `NtTraceEvent` at the top of a leaf list is the profiler's own cost. The site's
  speed is what counts, and the desktop is only a proxy for it: read a trace for the work
  our code does, not for the JIT (the browser interprets), ReadyToRun, or the start of the
  GPU (ANGLE, Direct3D: the browser has WebGL), and check a fix in the browser.
  2026-10-03, startup into the gallery at 1700x1100: the window's first frame at about
  1.7 s (half of it the JIT: neither our assemblies nor Avalonia's are ReadyToRun), then
  the tiles in view, 0.85 s of the UI thread's CPU (1.73 s before that day's fixes).
- In the browser: a `Stopwatch` around what is suspected and `Console.WriteLine`, in a
  Release publish served by `tools/serve.cs`, read with `webauto console`. Each `webauto`
  call is a process that takes CPU from the page: compare runs made the same way.
- How a page loads, as F12 shows it: `dotnet run tools/loadperf.cs -- <url> <out folder>
  [--seconds 20] [--size 1700x1000] [--trace] [--phone] [--wifi]`. A first visit (a new Edge
  profile, nothing cached) and a returning visitor's (a new browser on the same profile):
  the downloads, when .NET runs (the build line in the console), Avalonia's first frame and
  when the app took the splash down, and the long tasks of the page's thread - each load of
  a gallery tile is one. `--trace` records the first visit for F12's Performance panel (Load
  profile) and saves what was on screen every half second (`contactsheet` puts it on one
  image); `--phone` is a phone's screen, a 4 times slower CPU and Fast 4G (`--wifi` leaves
  the network alone: a local publish has no brotli, so its 4G download says nothing about
  the site's; the CPU part does). Nothing else may run meanwhile: a browser busy on the same
  machine (a page refreshed by hand, a headless one left behind) made every visit 1.5 to 2
  times slower. livegeometry.com on 2026-10-03, 1700x1000: on a first visit the splash
  closed at 2.7-3.0 s and the tiles in view were drawn by 6.2-6.6 s; for a returning
  visitor, 1.3-1.4 s and 4.6-4.8 s.
- How the gallery scrolls: `dotnet run tools/scrollperf.cs -- <url> <out folder> [--phone]
  [--wheel] [--passes 3] [--wait 15] [--dpr 1]`: the page loaded and left alone until its
  tiles are in, then scrolled down (loading the tiles it reaches) and back up (over loaded
  tiles: the scrolling alone), each with its frame intervals, the main thread's time in
  timer and animation-frame callbacks (Avalonia's dispatcher and render pass run in those)
  and its long tasks. `--phone` is 390x645 and a 4x CPU throttle, at a pixel ratio of 1:
  under the emulation Avalonia measured its canvas in physical pixels and divided by the
  emulated ratio, and laid the page out at a third of its width.
- Where the time went on 2026-10-07 (a local publish, headless Edge, returning visitor,
  1700x1000; the runtime is up at 0.2 s and everything below is after that): Avalonia's
  platform init 0.07 s, FluentTheme 0.03 s, the gallery page built 0.1 s, its first layout
  and frame 0.2 s (the frame itself 0.12 s of text shaping and Skia, native: the same under
  AOT), then the tiles in view 2.6 s in all, Stretchy Slime 0.47 s of it (the XML 0.03 s,
  reading the figures 0.13 s - the expressions compile there - adding them to the drawing
  0.19 s). The emoji font's parse is 1 ms. Under the 4x throttle every managed part is 4x:
  the first frame 3.1 s after the runtime, Stretchy Slime 2.6-3.2 s. Scrolling over loaded
  tiles at the 4x throttle: a p95 frame of 14 ms, where it was 40 ms before the tile
  distances were made arithmetic (above).
- AOT (`RunAOTCompilation` in the Browser csproj, on since 2026-10-07; set it to false, or
  publish with `-p:RunAOTCompilation=false`, to go back to the interpreter): the download is
  9.9 MB brotli instead of 5.3 (`dotnet.native.wasm` 35 MB raw, 7.2 MB brotli, which the
  browser also has to compile), the tiles load 2-3.5x faster (Stretchy Slime 0.24 s), the
  startup 1.3-1.7x, the gallery under the phone throttle is in at 6 s after the runtime
  instead of 12, and scrolling over loaded tiles takes 15% of the main thread instead of
  27%. A first visit on 4G pays about 4 s more for the download; a returning visitor has it
  cached. Only a publish compiles ahead of time (about 3 minutes more on this machine);
  `dotnet build` and `dotnet run` never do. Switching the property between two publishes
  on one machine wants the Browser project's `bin` and `obj` deleted in between: published
  the other way into the same `obj`, the runtime's native image and the assemblies didn't
  match, and the page died while loading the runtime ("MONO interpreter: NIY encountered",
  `mono_wasm_load_runtime () failed`). CI starts clean. A profile-guided partial AOT (only
  the methods a gallery load runs) would be the next thing to try for most of the speed at
  less of the size.

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
  the reader leaves the new bisector's `Sweep` at `Smaller` (the angle under 180°, see
  "Which of the two angles") and sets `IsLine` ("Whole line", saved as `Line="true"`); the
  verbs it would inherit from `Ray` (Convert to line/segment, Reverse) are vetoed in
  `CanEdit`. Morley's trisectors are expressions and get the same effect from
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
- A polygon lists its first point again at the end, to close it; the reader drops repeated
  vertices (kept, the inner triangle of Morley's was a "Quadrilateral" named JJKL). The
  gallery's polygons had the same repeat, also the hand-made and generated ones: removed
  2026-10-01.

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
  `drag <t> x1 y1 x2 y2 [steps] [--shift] [--alt]`, `wheel <t> x y <notches>`, `keys <t> "^s"`, `text`, `focus`,
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

## VB6 parity backlog

Still missing compared to VB6, roughly by value: unsaved-changes prompt; live cursor
coordinates in the status bar; undo/redo
captions naming the action; recent files; print; message, sound and launch buttons (the `.dgf`
reader leaves button types 1-3 out); "Create locus" on a point (the Locus tool does the
tracing); measurement label dragging constraints; rulers; "tool select once" option (every
construction returns to Drag); languages (en/ru/uk/de); step-by-step construction playback;
Calculator. Not a gap but a choice: double-click
zooms to fit, where VB6 opened properties (here selecting a figure shows them).

Done since the list was made: the point symmetric about a point and the inverted point are
Reflect with a point or a circle for the mirror; point shapes and sizes, name colors and dashes
are per style; show/hide buttons are `ShowHideControl`, made with the Show/hide box tool
(2026-10-07; VB6 dragged its buttons with the right button, here a drag is a drag and a
click ticks); settings persist (`SettingsStore`);
the axes are lines to build on (`AxisLine`); a click among overlapping figures is chosen
with Tab, a tap's in a menu, a selection in the context menu (`ClickChoice`).

## Not yet verified in the browser

The `Hyperlink` figure (only in old files, no tool makes one),
which fetches the drawing at its URL with `HttpClient` when clicked - not when read, which
put a failure's whole stack trace into its text, saved with it.
Saving through the browser storage provider works (drawings, and PNG and SVG pictures);
opening has only been checked as far as the picker (see "Files in the browser"). Copy image
finishes without an error in headless Edge, but reading the clipboard back is denied there:
what a paste gives was only checked on the desktop.
