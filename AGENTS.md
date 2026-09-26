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

## Gotchas

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
  `Polygon/Polyline.Points` only rebuild geometry when the property gets a different list.
  Figures keep their allocation-free point cache and call `Shape.PointsChanged()`
  (`PolygonShape`/`PolylineShape` in `WpfCompat.cs`, created by `Factory`). Suspect the same
  thing whenever something "renders once and then never updates".
- **Right button / hover** go through `Behavior.MouseRightClick` and `Behavior.GetCursor`
  (virtual, per tool). Tool letter shortcuts live in `UI/BehaviorShortcuts.cs` (also feeds the
  toolbar tooltips); plain-key handling (letters, arrows, +/-, G, H) is in `MainView.HandlePlainKey`.
  G toggles the grid through `DrawingHost.ToggleGrid`, which calls `CommandToolButton.UpdateToggles`
  so the ribbon button follows; `Command.Shortcut` puts the letter in a command button's tooltip.
- **The view** (`Figures/Coordinates/CoordinateSystem.cs`) is an origin in pixels plus `UnitLength`
  (pixels per unit). Everything goes through `Zoom(factor, focus)` (the point under `focus` stays
  put: the cursor for the wheel, the canvas middle for +/- and the menu), `Fit`/`SetView`
  (`ZoomExtend` = zoom to fit with a pixel margin, `CenterContent` = Home, `SetViewport` for files)
  and the resize handler, which keeps the middle of the canvas in the middle. "Content" for fit is
  `TryGetContentBounds`: points, whole ellipses (but a circular arc only its own extent: its
  ends and the compass points on it), labels - not lines. Panning (`MoveTo`) must not round the
  origin: a drag is many sub-pixel steps (0.5 px at 200% scaling) and rounding each one makes the
  plane run faster than the cursor. When testing pans with winauto, start the drag on an empty
  spot (a drag that starts on a figure moves the figure), use many steps
  (`drag <t> x1 y1 x2 y2 300`), and compare against the axis numbers. Labels have a fixed *pixel*
  size, so fit re-measures and refits a few times. `winauto wheel <t> x y <notches>` tests the wheel.
- **The grid adapts to the zoom** (`CoordinateSystem`, "Grid step" region): the labeled lines
  are 1, 2 or 5 times a power of ten apart, the smallest step that keeps them
  `MinimumMajorGridSpacing` (40 px) apart; between them a fainter tier (`MinorStyle` of
  `RectangularGridLinesCollection`) cuts a step into 5 (4 for a step of 2) when those would
  be at least `MinimumMinorGridSpacing` (10 px) apart. Both in DIPs. Shift-snapping lands on
  the labeled step (`MajorGridStep`), not on a fixed 1. A drawing that must keep its unit
  squares whatever the zoom says `<Viewport GridStep="1">` (`CoordinateSystem.GridStep`, a
  floor for the step; Pick's Theorem has it). Values come from an integer index times the
  step, rounded to 10 decimals, so labels never read 0.6000000000000001. Whether the grid
  shows is the drawing's own (`CoordinateGrid.Visible`): a new drawing starts without one
  and a file says. There is no global setting: one would leak from every loaded file (each
  gallery tile included) into the next new drawing.
- **Ellipses** (`Figures/Circles/Ellipse.cs`, `EllipseCreator`): center, the end of the long
  axis, the end of the short axis. The third point sets the short semi-axis by its *distance
  from the long axis* (`Math.GetDistanceToLine`; `EllipseArcBase` the same), so a point on the
  short axis is on the ellipse. The tool puts a free third click there: once the first two
  points are known it adds a hidden segment (the long semi-axis) and a hidden
  `PerpendicularLine` through the center, and the third point becomes a `PointOnFigure` on
  that line (`EllipseCreator.FindPointPlacement`, so the hover ghost and the preview ellipse
  show it); a third click on an existing point or another figure keeps that and the two
  helpers are removed again. The perpendicular's second point is the rotated axis, so the
  point's parameter is a fraction of the long semi-axis and scaling the long axis scales the
  short one with it. `ClickPreview` doesn't halo hidden sources: a hidden line's shape is
  never updated. `EllipseBase.UpdateVisual` puts no CenterX/CenterY on its `RotateTransform`:
  Avalonia rotates about `RenderTransformOrigin` (the middle by default), and WPF-style
  centering shifts a tilted ellipse off its center.
- **Vectors**: `Vector` = an invisible `Segment` + an `Arrow` (a 7-point polygon: shaft + head).
  The arrow is computed in pixels (`Arrow.HeadLength/HeadHalfWidth/HeadGrowth`, shaft = the
  style's stroke width), is filled with the *line* color, has no outline, and stops at the rim
  of its end point. `Vector.OnAddingToCanvas` gives it the default `LineStyle` before the base
  call, otherwise the polygon default (pale translucent fill) wins.
- **A point with an infinite coordinate does not exist** (`Math.Exists(Point)` rejects NaN
  and infinity; `IsValidValue` for a double). Avalonia's layout throws "Invalid Arrange
  rectangle" for a canvas child placed at infinity, unhandled: the desktop app dies and the
  browser app freezes on the spot. `CenterAt`/`MoveTo` in `Utilities` also refuse a non-finite
  place. Anything new that positions a control from figure coordinates must not hand it NaN
  or infinity.
- **`PathFigure.IsClosed` defaults to true in Avalonia** (false in WPF), and `IsFilled` to true.
  Every hand-built `PathFigure` must set both; a forgotten one draws a line from the end of the
  path back to its start.
- **No pixel snapping of figure geometry.** Lines are exact and point shapes have
  `UseLayoutRounding = false`, so lines hit their own points. Anything new that must line up
  with figures: same two rules. To check alignment, `winauto shot ... --region x,y,60,60 --zoom 16`
  around a point.
- **Right angles**: `Figures/Lines/RightAngleMark.cs` draws the two-sides-of-a-square sign. It is
  a passive visual owned by figures that are perpendicular by construction (`PerpendicularLine`
  at the foot, when the foot is on the base figure; `SegmentBisector` at the midpoint), gray and
  translucent, switchable per figure ("Right angle mark", saved as `RightAngleMark="false"`).
  Both figures derive from `PerpendicularLineBase`, which owns the mark. Which of the four
  corners it sits in is *stored* (`RightAngleMark.Corner`, 0-3 counterclockwise from the base
  line's P1->P2 direction, saved as `RightAngleCorner`), chosen once when the line first shows up
  and never derived from the geometry again - deriving it makes the mark flip-flop on rounding.
  A click on the mark with the Drag tool moves it to the next corner (undoable; the mark has a
  transparent square as its click area and marks the press handled so the Dragger never sees
  it). It hides when an `AngleArc` sits at the same vertex, because a measured angle of 90°
  (within 0.005°, i.e. exactly when the label reads "90°") draws the same sign itself in its
  own style instead of an arc. Deliberately not an auto-created angle figure.
- **Angle marks**: an angle is two figures - `AngleMeasurement` (the number) and `AngleArc` (the
  mark), same three dependencies. `AngleArc.ArcCount` 0-3 draws nothing, ), )) or ))) (saved as
  `Arcs`, default 1; spacing = stroke width + 2 px like VB6) and `Size` is the radius (`Radius`,
  default 16). The label's panel has the same "Arcs" box and forwards to its arc
  (`AngleMeasurement.FindArc`), because with 0 arcs there is little left to click. The pair is
  found by `AngleArc.FindCompanion`; "Convert to opposite angle" on either goes through
  `AngleArc.ConvertToOpposite(figure)`, which swaps the sides of both. `AngleArc.HitTest` covers
  the whole band from the first arc to the last (the base class only knows the first radius).
  The path keeps its first `PathFigure` even for "nothing" (no segments): Avalonia doesn't
  repaint a path whose figure list became empty. `DGFReader.ReadMeasureAngle` creates the arc
  from VB6's DrawStyle / AuxInfo(2) - not tested, there is no sample .dgf with an angle in the repo.
- **Any figure is draggable**: dragging a dependent figure finds its root free points and moves
  those, so gallery text can say "drag the circle" even when the points it is built on are
  hidden (Bubbles). Measurement labels are the exception (`Measurement.AllowMove`): a drag only
  changes the label's `Offset` from its anchor, although the figure has dependencies.
- **Dashes**: `LineStyle.Dash` is a `LineDash` enum (`Styles/LineDash.cs`: Solid, Dash, Dot,
  DashDot, DashDotDot - the VB6 DrawStyle 0-4, which `DGFReader` maps onto it). Inherited by shape
  and point styles, so circles, arcs and polygons dash too. Saved as `Dash="DashDot"` by the
  generic `EnumSerializer` (any enum property of a style or figure round-trips by name; an
  unknown name keeps the default). The pattern is put on in `LineStyle.OnApplied`, not through a
  setter, because `StrokeDashArray` counts in stroke widths and a selected figure is thicker.
  Anything that applies a style to a shape by hand (sample glyphs) must call `OnApplied` too.
- **Default point styles are by kind** (`StyleManager.AddDefaultStyles` / `AssignDefaultStyle`):
  `FreePoint` (yellow), `PointOnFigure` (green) - the draggable ones, size 10 - and
  `IntersectionPoint` (blue), `Midpoint` (orange), `DependentPointStyle` (gray, every other
  constructed point) at size 8. Looked up by *name*; a drawing from a file brings its own styles,
  and when the named one is missing (older files) the first point style is used. The click
  preview ghost uses the same lookup.
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
  on the source figures, equal-halves ticks for a midpoint. A tool that needs a figure rather
  than a point gets a halo on the figure a click would pick (`Behavior.GetFigureToPick`,
  `FigureCreator.FindFigureToPick`; an existing point the click would take counts too and gets
  a disc behind it; halos exist for points, lines, circles/ellipses, arcs, polygons and labels
  (a plate) - add a case to `ClickPreview.CreateHalo` for anything else). A tool that takes a
  figure where it would otherwise expect a point (Distance: a segment to measure; Circle by
  Radius: anything with a length as the radius, then the center) overrides
  `FigureCreator.FindFigureInsteadOfPoint` and uses the same test in its click handling;
  the base class then shows a halo and a hand and no ghost point over that figure. Circle
  by Radius unwraps a segment, a vector or a distance measurement into its two points
  (`CircleByRadiusCreator.FindRadiusEnds`), so the circle is the plain three-point one and
  outlives what was clicked; a polyline, an arc or an expression label stays the radius
  figure itself (`CircleByRadius.Radius` reads its `ILengthProvider.Length`). The preview is
  plain canvas visuals, never figures. Hit testing uses the *snapped* coordinates, so with snap
  to grid on the grid wins. Typed coordinates always give a free point. A point must never be
  placed on the figure being constructed (`FigureCreator.CanPlacePointOn`), that would be a
  dependency cycle.
- **Side panel (property grid) look**: the surface is a Border in `DrawingHost.CreatePropertyGrid`
  using `RibbonTheme` colors; the editors are restyled by scoped Avalonia styles in
  `PropertyGrid/PropertyGridTheme.cs` (no per-editor styling code). The Fluent theme paints
  hover/selected states on the template's ContentPresenter, so overrides must target that part.
  Rows share the label column width via `SharedSizeGroup`. Numbers are *shown* with
  `Settings.DisplayDecimals` (2; also the default of a label's `DecimalsToShow`) by every
  editor (`DoubleEditor`, `UpDownEditor`, `SliderEditor`); values keep their digits and a
  typed number is taken as typed. Avalonia raises a TextBox's TextChanged through the
  dispatcher, after the editor's guard is gone, so each text editor remembers the text it
  put in the box itself (`StringEditor.ShownText`) and ignores a TextChanged carrying it -
  otherwise showing a figure records a property set per text row (and writes the rounded value
  back). `LabeledValueEditor.SetValue` also skips a value equal to the current one. Layout is
  declarative: `[PropertyGridGroup("Name")]` on properties/methods boxes them together
  (editors, then their buttons in a row); `[PropertyGridDestructive]` on a method puts its
  button last, under a divider, with a trash can (`PropertyGrid.Arrange`, `MethodCallerButton`);
  `[PropertyGridIcon(PropertyGridIcon.Pencil)]` puts a small drawn icon (`PropertyGridIcons`,
  14 px, same grid as the trash can) in front of the caption. Every verb button has one
  (pencil/plus for edit and add, a padlock for fix/free, a green check for OK-like verbs, a
  cross for Cancel, glyphs for the conversions); give a new verb one too.
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
- **Numbers with a range** (`[Domain(min, max)]` doubles: stroke width, point size, font size) are
  edited by `SliderEditor`: text box + `Controls/UpDownControl` (repeat buttons, also Up/Down keys
  and the wheel over the box; `UpDownControl.Step` lands on whole steps and clamps) + slider.
- **Colors in files are always `#AARRGGBB`** (`ColorText.ToArgbHex`). Never `Color.ToString()`:
  Avalonia writes the *name* of a known color. `ToColor()` accepts names too, for old files.
  A gradient fill is a child element of the style (`<Fill><LinearGradientBrush>`), not an attribute.
- **The paper** is `Drawing.Background` (solid or gradient, white by default), pushed onto
  whatever canvas the drawing is attached to; edited in the side panel through the "Background"
  button on the Coordinates tab (`DrawingHost.ToggleDrawingProperties` shows the drawing itself
  in the property grid, undoable like any property). Saved on `<Viewport>`: a solid color as
  `Color="#AARRGGBB"` (white left out), a gradient as a `<Background><LinearGradientBrush>`
  child. `.dgf`: `PaperColor1`/`PaperColor2`/`GradientPaper` of `[General]`, top to bottom. A
  gallery tile takes a drawing's paper as its plate (pastel only for white ones) and turns its
  caption white on a dark plate. Point sizes in the gallery are standard (10 for a draggable
  point, 8 otherwise): `dotnet tools/pointsizes.cs -- <folder> [--apply]` lists and raises
  undersized point styles. The Rose's 90 control points stay at 5 px on purpose (at 10 they
  swallow the flower).
- **Numbers are figures** (`Figures/Values/Number.cs`): a `Number` is a root in
  the dependency graph with no shape and no hit test, a `Value` the property grid edits
  (undoable, recalculates dependents), a length (units) and an angle (degrees; `Angle` gives
  radians like every angle provider). Named n1, n2, n3. An expression names one by that
  name (`[n1*2]`, `TreeBuilder.CreateIdentifierExpression`) and the label depends on it.
  Saved as `<Number Name="n1" Value="3" Auxiliary="true" />`. `Auxiliary` (on every
  figure, `IFigure.Auxiliary`) marks one created on demand for another figure: it is
  removed with its last dependent (`RemoveFigureAction.FindOrphanedAuxiliaries`) and comes
  back on undo. Every deletion goes through `RemoveFigureAction`: the Delete key
  (`Drawing.DeleteSelection`) removes the selected figures one by one in a transaction, so
  point labels come back on undo and a polygon loses a vertex rather than dying, like the
  grid's Delete button. The Dragger ignores Numbers among the roots of a dragged figure (they
  have no place to go). There is no tool that makes a Number yet; a typed value in the
  Translation tool becomes one.
- **TranslatedPoint** (`Figures/Points/TranslatedPoint.cs`) is the point at a distance and
  a direction from its source. Each quantity is either *tied* to a figure it depends on (a
  vector for both, a length provider; for the direction also any `ILine` - segment, ray,
  line - taken the way it points, from its first point to its second, or an angle
  provider; a Number for a typed value) or *free*: the
  parameter dragging changes, kept in the file. Free direction = slides on the circle
  around the source; free distance = on the line through it, signed. Roles are stored by
  index into the dependency list, never inferred from types (a `Label` is both a length and
  an angle provider). Property grid: Distance / Direction (degrees) rows editable when free
  or held by a Number, read-only with the source's name otherwise; "Free distance" /
  "Free direction" checkboxes convert (a freed Number is retired, not deleted, so undo
  brings the same one back; freeing both is disabled). A point with a free
  quantity takes the green `PointOnFigure` style. File: `DistanceSource="n1"` /
  `DirectionSource="..."` for tied, `FreeDirection="true" Direction="30"` for free (degrees).
  A file with none of those is the old format (typed values as attributes, roles by
  position, Direction in radians); `UpgradeLegacyValues`, called by the deserializer,
  gives the typed values auxiliary Numbers. `Transformer.CreateTranslatedFigure` takes the
  two sources, shared by every point of a translated figure.
- **The Translation tool is stepwise** (`TranslationCreator`): source figure,
  then distance, then direction, then - when something was left free - a click that
  places the point. The side panel shows only the step at hand (`PropertyBag` is the
  step's `ValueStep`: an up/down number box (`UpDownEditor`, chosen by
  `[PropertyGridPreferredEditor("UpDown")]`), OK (Enter too) and, for a point source and at
  most once, Free), remembers the last values for the session, and goes away when the
  construction ends: `DrawingControl` re-shows the behavior's `PropertyBag` on "construction
  complete" as well, and the creator clears its panel in `Stopping` because that event is
  raised before `Started` resets the state. A click on a figure during a value step takes
  it (a vector at the distance step gives both and ends the construction); a click on
  empty paper does nothing. The riding point of the placement step is the real
  `TranslatedPoint` as a temp result, moved by `MoveTo` on every mouse move; the click
  stores the place and the final one is created there.
- **Rows that are editable only sometimes** (`PropertyGrid/ConditionalPropertyValue.cs`):
  `[PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]` on the property,
  and the figure implements `IConditionalProperties` (`CanEdit(name)`, `Caption(name,
  default)`). A read-only text row is grayed (`StringEditor` disables the box). Used by
  `TranslatedPoint` (Distance/Direction/Free rows) and `Segment.Length`: setting a
  segment's length stretches it once (not a constraint) by moving the second end away from
  the first, or the first if the second can't; an end can take it when it is a free point
  or a translated point sliding along this segment's own line, and the other end isn't
  built on it (a fixed-length segment's far end would just follow). Otherwise read-only.
  The same interface vetoes *buttons* by method name (`PropertyGrid.GetCallableMethods`).
- **Fix length / Free length** are verbs on a segment, a vector, a regular polygon and the
  circles (`IFixableLength`: `Length`, `FixLength`, `FreeLength`; the rules and the surgery
  are in `Figures/Lines/LengthConstraint.cs` for any end kept at a distance from a pivot -
  a segment's end from the other end, the polygon's vertex from its center, the vector's
  through its inner segment, a circle's rim point from its center or the two radius points
  (`CircleBase.RadiusPivot`/`RadiusEnd`, null for a circle by equation or a radius taken
  from a figure - then the Radius row is read-only and no verb shows); the captions come
  from `IConditionalProperties.Caption` ("Side", "Radius", "Fix radius"); only the
  applicable verb shows). Fix length turns the end that could take a new length (same rule
  as above) into a `TranslatedPoint` from the pivot with an auxiliary Number at the current
  distance and a free direction; nothing moves, the end turns green, and the Length (or
  Side) row then edits the Number. Free length puts a `FreePoint` back where the fixed end
  is (the Number goes with it). A sliding end (translated, distance free) is fixed or freed
  by toggling its `FreeDistance` instead. **Right after such a figure is made** the side
  panel shows a `LengthPanel` (`FigureCreator.ShowCreatedFigure`): the length, a Show
  checkbox that puts a `DistanceMeasurement` on the figure or takes it off
  (`IFixableLength.MeasuredFigures` says what it measures: the segment itself, a circle's
  two radius points; null for the regular polygon, whose sides are its own children, so no
  checkbox), the verb, and Done, which just closes the panel - plus a line in the status
  bar. The tools have no length box; a square's panel is its base side's. Not shown when
  nothing can move (a figure built on existing dependent points: `CanEdit("Length")` is
  false). The Show setter changes the figure list directly, like the point's Free toggles:
  the grid records the property set as the undo step and the undo library refuses an action
  recorded from inside another. A creator's undo transaction spans one construction, opened
  at the first click (`FigureCreator.EnsureTransaction`, before the point that click makes)
  and committed with the figures, so an edit made in that panel between constructions is an
  undo step of its own. The swap is `Actions.ReplacePoint`: add the replacement, regraft
  dependents, hand over the point's name label (which `ReplaceFigureAction` skips on
  purpose), remove the old point, give the new one the old name - one transaction, undo
  restores the old point with its Number. Undo of a *drag* of a constrained point is by
  offset and lands only near the old parameter, as for a `PointOnFigure`.
- **The library's own types shadow framework ones**: `Math`, `Ellipse`, `Polygon`, `Path`...
  In a file-scoped-namespace file a `using X = ...;` alias does NOT win over a type of the
  enclosing namespace - write `System.Math.Max`, `Avalonia.Controls.Shapes.Ellipse` in full.
- **`Style` and `Setter` are ambiguous in the library**: `DynamicGeometry.Style`/`Setter` are the
  WPF shims and shadow Avalonia's. For real Avalonia styles alias them
  (`using AvaloniaStyle = Avalonia.Styling.Style;`).
- **Gallery and routes** (`LiveGeometry/Gallery/`, `MainView` "Pages" region). `MainView` shows
  either `GalleryView` (the start page: "New Drawing" + a tile per drawing) or the editor. Three
  states: gallery, a gallery drawing (`CurrentSample`; the toolbar shows ◀ n/N ▶ + title, Page
  Up/Down, edits are dropped silently, Save = save as and the drawing becomes the user's own), or
  the user's own drawing (only the Gallery button). The user's drawing object is parked in
  `OwnDrawing` (with its undo history) while they look around, and re-attached by "Back to My
  Drawing". Paths: `/` gallery, `/gallery/<slug>`, `/drawing`; `AddressBar` is the abstraction,
  `BrowserAddressBar` + `main.js` do pushState/popstate, so Back works and links can be shared.
  `index.html` needs `<base href="/">` for that (all routes serve it; `web.config` and
  `tools/serve.cs` fall back to it). All page changes go through `MainView.Show*`. Started
  at the address of a drawing (a shared link), `MainView` comes up in the editor and the
  gallery isn't built at all (`IsDrawingPath` in the constructor; the address is readable
  synchronously since `main.js` registers its imports before the runtime starts).
- **Gallery drawings** are embedded `.lgf` (`Gallery/Drawings`, order and titles in
  `GalleryCatalog`) and are *the* source: edit them directly (by hand as XML, or open, edit,
  save). They were forked once from the Windows Phone samples and the DG 1.0 CD library;
  `tools/gallerize.cs` made that fork and must not be rerun (it would overwrite the
  hand-edited drawings). To add another CD drawing: `--check` it (below), take the `.lgf`,
  drop the old labels, add `Title`/`Description` labels in the `GalleryTitle`/`GalleryText`
  styles (copy from any drawing here), set `Grid`, list it in the catalog. Each has two
  labels, `Title` and `Description` (may contain live `[AB^2]` expressions - then list the
  points as dependencies), together "the caption". The caption is *pinned to the screen* (see
  "Pinned labels" below): the files carry `Pin`/`OffsetX`/`OffsetY`/`WrapWidth`/`Backdrop`
  and no X/Y, and `GalleryDrawing.Fit` recomputes them at open time: a 400 px column at the
  right in a wide canvas (`TopRight`, wider only for a heading that needs it, and only if that
  leaves the figure 40% of the width), a strip at the bottom in a tall one (`BottomLeft`,
  canvas width minus margins). Labels are sized in pixels, so the figure gets the canvas
  minus the text and the zoom is computed from that in one go - never iterate "place text,
  zoom to fit": it runs away (zooms out a little more per round) once the text needs more
  than its share. The figure keeps at least 40% of the canvas; if the text doesn't fit it
  runs off the bottom, and the reader drags the caption up (a pinned label is dragged by its
  offset). Refitted on resize until the first edit, through `Drawing.SizeChanged`
  (`MainView.KeepFitted`) and not the canvas's event: the coordinate system's own resize
  handler shifts the origin, and it has to run first. A file with a caption opened from disk
  (`MainView.OpenDrawing`) is fitted the same way. The explanations have no hand line breaks
  (the column wraps them; a blank line separates paragraphs). On an iPhone in portrait
  (canvas about 390x565) the long explanations run off the bottom. Drawings with
  `Grid="true"` (graphs) keep their file viewport in view (`GalleryItem.Plane`), because
  graphs and lines have no bounds.
- **Pinned labels** (`Label.Pin`, `Figures/Controls/LabelPin.cs`): a label is either a place
  in the plane (X/Y, the default) or pinned to a corner of the canvas: `PinOffset` is the
  distance in pixels from that corner to the *same* corner of the label, so a right-pinned
  label keeps its right edge and a bottom-pinned one its bottom while the plane zooms and
  pans under it. `Coordinates` always tell where it is in the plane right now
  (`UpdateVisual` sets them), so hit testing, dragging (`MoveToCore` turns the new place into
  an offset; undo works) and the property grid ("Pinned to", radio buttons; switching keeps
  the label where it is on screen) need nothing special. `WrapWidth` (px, 0 = none) gives the
  TextBlock a fixed width and wrapping, so two labels of one column line up at the left;
  `Backdrop` puts a plate of the paper's color (solid papers only) with an 8 px padding behind
  the text. Pinned labels draw at `ZOrder.Controls`, above figures; zoom to fit ignores them;
  thumbnails hide them (`GalleryDrawing.HideText`). Saved as `Pin="TopRight" OffsetX OffsetY
  WrapWidth Backdrop`. `Label.MeasureSize` invalidates the Border before measuring: its measure
  stays valid when only the TextBlock inside changed, and answers with the old size. Dragging
  any label moves the label, not the figures its expressions name (`Label.AllowMove`, like
  measurements).
- **Scenes** (`<Scene Left Top Right Bottom />` under `<Drawing>`, 1 or 2, `Drawing.Scenes`) are
  opt-in suggested views for the few drawings whose content has no useful bounds - ground
  that goes on forever (Castle, The Falling Ladder) or a locus (Spiral). Only where they
  exist: the gallery fit and the tile show the scene nearest in shape to the room
  (`Drawing.ChooseScene`, landscape vs portrait) instead of the content bounds; every other
  drawing is fitted as before. The fitted one is `Drawing.ActiveScene`, and a gradient paper
  then spans the scene, not the canvas, solid beyond it (`Drawing.PlaceBackground`, re-pinned
  on every view change), so the sky stays put when the view moves.
- **Generated gallery drawings** - regenerate rather than edit the file: Line of Best Fit
  (`dotnet tools/bestfit.cs -- <the .lgf>`), Fibonacci Spiral (`tools/fibonacci.cs`),
  5 Platonic Solids (`tools/platonic.cs`, `--alpha` for a translucent variant). Techniques
  they use: hidden `PointByCoordinates` whose coordinates are expressions are the library's
  variables (Best Fit keeps its means, sums and slope/intercept in three of them); a
  `PointByCoordinates` expression of one free base pair (`A + x·right + y·up`) keeps a
  tiling from coming apart.
- **Drawings that must not fall apart** when a kid drags the wrong thing (Castle, The Falling
  Ladder): fixed points are `PointByCoordinates` with constant coordinates (a polygon of those
  has no free point, so dragging it does nothing), and the only things that move are
  `PointOnFigure` sliders on hidden rays or segments (a foot along the ground, a top up a
  vertical ray from that foot, a ladder length on a segment whose length is the wall) plus a
  free point or two (sun, cloud).
- **The Spiral's polygons**: a `Locus` samples its driver in `Locus.StepCount` = 60 steps, so
  the Spiral's orange "Drag to here" rings (a square spiral at 90°, a triangle at 120°...) sit
  at D.Y + 60·2π/n for n = 3, 4, 5, 6. Changing `StepCount` moves the rings.
- **Tiles are live drawings**, not bitmaps (`DrawingThumbnail`): a `Drawing` on its own 560x380
  canvas inside a Viewbox, no Behavior attached, text hidden, loaded one per idle tick. Hovering
  makes the draggable points drift (`IsAnimated`). A load error of a tile goes to the console as
  `Gallery: <file>: ...` - that is the quickest way to check all drawings at once.
- **Old files and circle-line intersections**: `Math.GetIntersectionOfCircleAndLine` once
  swapped P1/P2 for a line through the center. Drawings from before (the phone ones) pick the
  other intersection, so squares built with perpendicular + circle flip inward. Files carried
  no usable version then, so the fix is opt-in: `<Drawing IntersectionOrder="Legacy">` makes
  `DrawingDeserializer` swap the algorithm of every intersection whose line passes through the
  center (`IntersectionPoint.UpgradeLegacyCircleAndLineOrder`, numeric test, in construction
  order). No file in the repo carries the mark (`LiveGeometry.Desktop.exe --modernize <folder>`
  writes into a file what loading it upgrades); the code stays for old files from elsewhere.
- **Label offsets are pixels, and files have a version.** A `LabelWithOffset` (point
  labels, distance/angle/area measurements) sits at `Offset` pixels from its `Anchor` in the
  plane (the point, the segment's middle, the vertex), so the gap keeps its size at every
  zoom; in units of the plane it would shrink onto the point on a phone and drift away zoomed
  in. The base class does the placing (`PlaceFromOffset`, `MoveToCore`, `UpdateVisual`);
  subclasses give the anchor and set the text. `<Drawing Version="1">`
  (`Settings.CurrentDrawingVersion`) says the offsets are pixels; a file without it is
  upgraded on load, after the viewport, at the zoom it opens at (`UpgradeOffsetFromUnits`) -
  the best guess for where the author had it, since files don't say. Anything new that
  measures something in the plane but places text should keep its distance in pixels the
  same way. A point label lives in an orbit (`PointLabel.ClampPosition`, applied on every
  move): no further from the point than about its own size, and no edge or corner of its text
  box nearer to the point's rim than `PointLabel.Clearance` (4 px) - the box, not the glyphs,
  so the letters sit a little further out than that.
- **Saved `.lgf` declare `encoding="utf-8"`** (`DrawingSerializer.Utf8StringWriter`); files
  from older builds say `utf-16`, which our own loader tolerates but `XDocument.Load` does not.
- **Zoom to fit** counts the vertices of visible segments/polygons/Béziers even when the points
  are hidden, measures labels itself (`Measure`, their Bounds are stale right after a load), and
  takes an optional rect to keep in view.
- **No menu.** New, Open, Save | Undo, Redo are one row (`LiveGeometry/MainToolbar.cs`) above
  the ribbon, in the ribbon's blue-gray tint a shade darker than its header row (#DDE0E3), with
  the header row's kind of line along its bottom; at its right a faint Octocat links to the
  repository, and its tooltip is the build. Icons are drawn in code (`MainToolbarIcons`, 20x20
  grid drawn at 24 px); Undo/Redo follow `DrawingControl.CommandUndo/CommandRedo` as command
  observers. `MainToolbar` lays its three parts out itself: buttons left, stamp right, and the
  tour group (◀ n/N ▶ + title, bigger; `MainToolbarGroup`) in the room between: the arrows are
  centered there and the title hangs off their right, so the arrows don't move with the title;
  only when the title wouldn't fit whole do the arrows move left, as far as it needs. When the
  group doesn't fit beside the buttons at all it wraps to a second row, left-aligned; when
  only the stamp doesn't fit, the stamp is hidden. The first button is the app's mark
  (`AppIcon`, the logo in the corner) and folds the whole ribbon away (Ctrl+F1,
  `MainView.UpdateRibbon`): folded by default when a gallery drawing opens on a small screen
  (under 700x500), open otherwise; once pressed, the user's choice holds for the session.
  While the ribbon is open the mark is drawn as a tab (`TabOutline`, the selected group
  header's shape) opening through the strip's bottom line into the header row, filled with
  the tools strip's gray that turns into the header row's gray at the bottom
  (`MainToolbar.IsTabOpen`); when the tour group has wrapped to a second row the tab would
  span both, so the button shows the blue checked plate instead. Everything else is keys:
  Ctrl+N/O/S/Z/Y/A/C/V are handled on key *down* (`MainView.HandleControlShortcut`; on key up
  Ctrl may already be released and a bare S is the Segment tool), plain keys in `HandlePlainKey`.
  Lost their menu entry and are unreachable for now: Lock, Figure List, the settings page.
  Not on the Selection tab (obscure for the audience; the commands and settings are still
  there in `DrawingHost`): Ortho, Polar, Snap to grid, Snap to point, Snap to center. Shift
  while dragging or clicking still snaps to the grid (Pick's theorem says so), and a click
  near the middle of a segment still makes a midpoint.
- **The app runs under the invariant culture** (`App.UseInvariantCulture`, first thing in
  both `Main`s; the Browser project also builds with `InvariantGlobalization`, so no ICU
  data is downloaded). Numbers use a point everywhere: the expression language, `.lgf` and
  `.dgf` files, Avalonia path markup, labels. The spots that format or parse numbers for a
  parser (`AppIcon`, the angle tool icon, expression literals, `IniFile`) say
  `InvariantCulture` explicitly all the same; a `$"M{x},13.5"` under a decimal-comma culture
  makes `Geometry.Parse` throw. To test a culture: `webauto stop`, then
  `webauto start <url> 1280 800 --lang de-DE`.
- **Every exception is shown** (`MainView.CurrentDomain_FirstChanceException`, as in Helix
  and the structured log viewer): the message goes to the status bar ("Error: ...") and the
  whole text to the side panel as an `ExceptionReport` page ("Something went wrong": Error,
  Details), also printed to the console. Reported on the UI thread; exceptions the report
  itself throws are dropped (`reportingException`), the same text thrown again only
  refreshes the status bar, and OperationCanceled/Aggregate/TargetInvocation are ignored.
  Add to `IsBenign` when a framework exception turns out to be noise (the browser's file
  picker throws `JSException` "AbortError..." on cancel, which Avalonia turns into null).
  A `JSException` has no .NET stack, so the handler captures `Environment.StackTrace` at
  the throw for the report. Text boxes of the property grid wrap at 480 px
  (`PropertyGridTheme`), so a caption or a stack trace doesn't stretch the panel across the
  window.
- **Files in the browser** (`MainView.SaveDrawingToFile`/`OpenDrawingFromFile`): the File
  System Access API takes a file type only as a MIME type with its extensions, and Avalonia
  drops a `FilePickerFileType` without `MimeTypes`, so the types carry `application/xml`
  (.lgf) and `text/plain` (.dgf). The stream from `OpenWriteAsync` has only `WriteAsync`; a
  `StreamWriter` flushes synchronously on dispose and throws "Browser supports only
  WriteAsync", so the text is written as bytes with `WriteAsync`/`FlushAsync`. To test saving
  headless (no native dialog): `webauto eval` a fake `globalThis.showSaveFilePicker` returning
  a handle with `kind`, `name`, `getFile`, `queryPermission`, `requestPermission`,
  `createWritable` (`write`/`close`/`seek`/`truncate`) *before the first picker use* -
  Avalonia's picker polyfill captures the global when its storage module is first imported,
  later replacements are ignored - then click Save and read back the bytes; throw a
  `DOMException(..., "AbortError")` for cancel.
- **Keyboard focus drifts into tool panels.** A tool's PropertyBag panel (e.g. "Point by
  coordinates") takes focus into its TextBox after every construction step, so neither the canvas
  KeyDown nor `MainView_KeyUp` (which skips TextBox focus) sees keys then. Anything that must
  always work (Escape) belongs in the `MainView_KeyDown` tunnel handler. The "Point by
  coordinates" panel (X/Y boxes of `FigureCreator.Dialog`, `ShapeCreator.ShapeDialog`,
  `FreePointCreator.CoordinatesDialog`) only exists while the "Point by coordinates" toggle on
  the Coordinates tab is on (`Settings.EnablePointByCoordinates`, off by default); the Point
  tool has no panel otherwise.
- **Toolbar look** is centralized in `UI/Ribbon/RibbonTheme.cs`; `ButtonGrid` draws the
  hover/pressed/checked plate. `Ribbon` and `TabPanel` replace the Fluent TabControl/TabItem
  templates with their own (in code): the header row has a bottom line *behind* the headers and
  the selected group header paints a tab shape over it (`TabOutline`, laid out wider than the
  header by its flare via negative margin). To bring the group headers closer together, change
  `ButtonGrid.HeaderOverlap` (neighboring headers share their flare zones - only the selected
  one draws feet), not the padding inside the tab. Group headers show the icon of the group's
  active tool. An on/off `Command` exposes `IsChecked` (a `Func<bool>`), which its button
  re-reads after any toggle is clicked - don't go back to `CheckBox` icons.
- **Browser has no system fonts.** Text renders only because `Avalonia.Fonts.Inter` is embedded
  (`.WithInterFont()`). Inter reaches text by *inheritance* from the window (the theme sets it
  there); an explicit family that doesn't exist - and `FontFamily.Default` and
  `FontManager.DefaultFontFamily` too - renders as Noto Mono in the browser. So `TextStyle` sets
  the font a drawing names (Arial, Segoe UI) only when it really resolves, and naming "Inter"
  in a drawing doesn't work (embedded, not installed).
- **The splash** (`wwwroot/index.html` + `app.css`, shown until Avalonia adds `splash-close`)
  is Euclid's first construction drawing itself: a 7.5 s CSS animation loop on inline SVG,
  with a progress bar that `main.js` feeds from a `withResourceLoader` wrapper counting
  fetches (the loader's `onDownloadResourceProgress` is on the module config, which the
  .NET 10 host builder doesn't expose). The wrapper hands back a Response only for assemblies,
  the wasm and ICU data; the runtime's own JavaScript modules and its config must be left to
  the default loading (a Response for `dotnetjs` leaves the site spinning on the splash
  forever). A change to `main.js` needs the publish smoke test below before deploying.
  `?splash` on the url shows the splash without starting the app, for working on it: serve
  the source `wwwroot` with `tools/serve.cs` and open `http://localhost:<port>/index.html?splash`.
  It animates regardless of `prefers-reduced-motion` (deliberately no rule for it: the query
  follows the Windows "Animation effects" setting, off on this machine, and a still splash
  looks stuck).
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
  `drag <t> x1 y1 x2 y2`, `keys <t> "^s"`, `text`, `focus`, `place <t> x y w h`.
  Target = process name | `pid:N` | `hwnd:0x..` | `title:substr`.
  - The VB6 app starts maximized on a 4K/200% monitor: `place <t> 100 100 1500 1000` first. The
    Avalonia desktop app remembers its window (`LiveGeometry.Desktop/WindowPlacementPersistence.cs`,
    Get/SetWindowPlacement as in Helix; `%LOCALAPPDATA%\LiveGeometry\MainWindowPosition.txt` =
    flags,showCmd,min x,y,max x,y,left,top,width,height). It is left at 100,100 1700x1100 for
    testing - `place` is only needed again if someone resized it. Close test instances with
    `(Get-Process -Id N).CloseMainWindow()`: that goes through the normal close path, which is
    what saves the placement (and it is more reliable than Alt+F4).
    `LiveGeometry.Desktop.exe --arrange` opens the gallery with its tiles draggable: a tile
    takes the place of the tile under the pointer as it goes, and every drop rewrites the
    `Items` block of `GalleryCatalog.cs` in the new order (`GalleryCatalog.SaveOrder`, by
    `[CallerFilePath]`; the drawings keep their lines, comments between them go). Rebuild to
    see it in the app; the running window already shows it.
    `LiveGeometry.Desktop.exe --gallery <slug>` opens a gallery drawing the way the browser's
    `/gallery/<slug>` does (tour group, caption fit, ribbon folded on a small window) - a file
    on the command line opens as the user's own drawing instead. An iPhone 13 Pro in
    portrait gives Safari a 390x645 viewport, of which the two toolbar rows take 82 and the
    canvas gets 390x563; landscape is about 844x310 with a 844x250 canvas (estimated). For
    the browser build that is `webauto start <url> 390 645`; for the desktop app at 200%
    scaling `place <t> 100 100 780 1352` (the chrome and toolbar take 110 DIPs) and
    `1688 800` for landscape.
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
