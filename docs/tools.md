# The ribbon

Tools are `Behavior` subclasses found by reflection (`Behavior.LoadBehaviors`): `[Category]`
names the tab, `[Order]` the place in it, `[Ignore]` keeps one off the ribbon; the button shows
`Name`, the tooltip adds the letter from `UI/BehaviorShortcuts.cs`, the status bar shows
`HintText`. Tab order is the order of the constants in `BehaviorCategories`; tools the user
defines go onto Misc. Non-tool commands are added in `MainView.InitializeCommands`. Tab by tab
(letter in parentheses):

- **Selection**: Drag (Q) - drags points and figures (with Alt a point snaps onto figures and
  lets go of them, see "Snapping and releasing points", and a text goes onto a figure or
  off it, see "Texts attached to figures"; while a point is dragged the
  status says what Shift and Alt do to it, `Dragger.ModifierHint`); a press on a figure
  selected with others drags them all as one piece (`Dragger.FindSelectionRoots`: the roots
  of the whole selection, each once, by the same offset; no Alt snapping, no point jumping
  under the cursor; a locked figure or root holds them all, a caption stays); also the tool
  every construction returns to. A drag with the right button draws a lasso, and a right
  click opens the context menu at the release, not the press (`Behavior`, "The right
  button": a tool takes a right drag through `MouseRightDragStarted`; one that doesn't gets
  the click at the release wherever it is). Box, Lasso (2026-10-11; `RegionSelector`:
  `BoxSelector`, `LassoSelector`) - a drag anywhere, on figures too, adds what the region
  takes to the selection, with Alt takes it out; a click selects as Drag's does. While the
  drag lasts (`RegionSelection`) the region is a band of the halo's stripes in the halo
  layer, what it takes is selected live so the halos show the outcome, and Escape puts the
  selection back (`Behavior.AbandonGesture`, asked by `MainView` first). What a region takes
  is judged in pixels on the shapes on screen, fill and stroke as drawn (`SelectionRegion`:
  Skia path operations through `CombinedGeometry`, after a bounds test), by AutoCAD's rule
  (2026-10-15; `TakesTouchedFigures`): a box dragged to the left of where it began takes
  every figure it touches (a dashed band), to the right only those wholly inside (a solid
  one); a lasso goes by the way it first went, left or right. Wholly inside: a figure the
  window shows whole must be inside whole; one that goes past the window or under the side
  panel only with what of it lies `EdgeMargin` (50 px) in from the window's edges and the
  panel (`VisibleArea`) - or, when nothing does, what shows at all - so a big circle is
  taken without a drag along the very edge. Lines, rays and graphs are never shown whole,
  so they go by that rule too, and a line or ray also goes with the points it is drawn
  through when the region takes them all (line AB with A and B, a parallel with its
  point: `IsTakenWithItsPoints`). (Until 2026-10-15 every region took what it touched;
  GeoGebra takes only what is wholly inside, and no line at all.) Never hidden figures,
  names, Numbers, the axes or a gallery drawing's text. Figure List (toggle): see "The
  Figure List".
- **Points**: Point (P) - free, on a figure, or at an intersection; Midpoint (M) - two points
  or a segment; Intersection (I) - two figures that cross, the click on the second picks the
  nearer crossing (`PointPlacement.Intersection`, shared with the Point tool), a graph or a
  Bezier curve among them (see "Crossings with curves"); Point inside
  (O; 2026-10-11) - one click inside a polygon, a circle or an ellipse, a sector or a
  circular segment, a closed or filled Bezier path, filled on screen or not (the hover's
  halo and ghost say which; `IRegion` is the contract, see "A point inside a shape"), the
  smallest of nested regions first and Tab for the next; nothing on a point or inside
  nothing; Coordinates (X) -
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
  the panel shows the new line's angle instead (see "Tied values"). Both panels and the
  line's grid have Horizontal and Vertical (2026-10-12; the nearer way, as a segment's: see
  "Setting a segment's length"), which type the angle again when it was taken from a figure
  (`LineAtAngle.Align`; the panel after the tool calls it itself, since the line's own verbs
  show its grid and that would end the panel). Perpendicular Bisector
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
  Enter, a right-click or a click on a vertex closes it; Regular polygon - center then a vertex;
  Polygon on a side - two adjacent vertices, the polygon to the left of the first to the
  second (2026-10-13). The last three make plain figures (see "Squares and regular
  polygons"); the two regular polygon tools have a Number of sides panel, shared.
  Triangle and Polygon show no length panel (a side's length means nothing for them). They
  (and Square, for its first side) draw their sides as segments (the polygon's own outline is transparent in the default
  style), except where a visible segment, ray, line or vector on the two points is there
  already (`FindLine`; not any line that depends on both - a perpendicular bisector of the
  two took the side's place and left it undrawn). (Polygon intersection
  exists but is `[Ignore]`d.) Bezier path (no letter): see "Bezier paths". Difference,
  Intersection, Union (U) (2026-10-23; `RegionCreator` and its three subclasses, Venn icons):
  a click on each shape with an inside (inside it, the smallest of nested ones first and Tab
  for the next, or on its edge), then Enter, a right click, a click on the paper or on a
  shape taken already makes a `Region` of them (see "Regions"), once there are two; the
  first click of Difference is what the rest is cut out of. From the first click on the
  hover tints what the region would be with the shape under the cursor
  (`Behavior.GetRegionPreview`, `ClickPreview.ShowRegionGhost`), and from the second the
  side panel says what the region would read as ("square ABCD without circle c") and
  offers Done (`RegionCreator.RegionPanel`, the tool's `PropertyBag` while there are two
  and the transaction is open). The region's page shows after it (the Operation row, its
  area and perimeter).
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
  polygon's side segments are traced, and a side without one gets a segment first, as the
  shape tools would have drawn it. Reflect
  shows no panel after a construction (the length panel came up for such a segment).
  Every other transformation of a polygon or polyline also carries over the
  visible segments along its sides (`Transformer.AddSideSegments`: the shape tools draw
  sides as segments, and the image was a shape without an outline), same style and
  marks; they go into the list before the image, which callers take to be the last. An
  image is a clone of the source given the transformed dependencies; a composite's clone
  made its parts on the source's when it was read, and `Transformer.MovePartsOver` moves
  them (a reflected regular pentagon, a composite then, had two sides and its inside on
  the source's first vertex). The
  GeoGebra reader says `sideSegments: false`: a file has the image's sides as objects
  of their own, and the copies came out hidden.
- **Coordinates**: Grid (G) (command; a split button, its arrow a menu of Axes and Grid
  lines: `Command.DropDown`, `ButtonGrid.AddDropDown`); Function - an expression in x; Line - by
  slope and intercept expressions; Circle - by center and radius expressions; Point by
  coordinates (toggle: gives the point tools an X/Y panel).
- **Measure**: Length (2026-10-15; replaced Distance and Perimeter) - two points, or a
  first click on a segment, vector, polyline, Bezier curve or path, arc, sector, circular
  segment, circle, ellipse or polygon, which measures the figure's first length, and the
  panel after it offers the others (see "Lengths" under "Design decisions"); once the
  first point is down the next click is the other end, on a figure if it is on one (it
  measured the figure and left the first point behind); Angle (J) - vertex then two side points,
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
  circle, sector, circular segment, Bezier path with an inside (closed or filled;
  `AreaMeasurement.IsMeasurable`), or a list of points, Enter or a right click when the
  points are done. Both tools take a figure through `FindFiguresInsteadOfPoint`, where the
  hover looks too, so the cursor is the hand and the figure has its halo (the Area tool
  took shapes in `Click`, and showed the cross of a free point over a filled polygon),
  and a path is hit by a part, mapped to the path (`FigureParts.Whole`);
  Number (2026-10-15) - one click where the number goes, then a value or a formula typed
  in the panel's Formula box (see "Numbers are figures"; while the box has the keyboard a
  click on a figure puts a reference to it there: "Building a formula by clicking");
  Slider - where it sits, then where
  its knob starts (or press, drag, release): a number with a handle, taken wherever a tool
  asks for a length or an angle, named in expressions (a, b, c).
- **Misc**: Bezier - four points; Locus (D) - a point that depends on a point on a figure;
  Text - a label at a click on the paper, or a text attached to the point, segment, line,
  ray, vector, circle or polygon clicked (see "Texts attached to figures"; the hover's halo and hand
  say which, Tab chooses among overlapping figures); Show/hide box (`ShowHideCreator`) - the figures selected
  when it is picked are picked already, a click on a figure picks or lets go of it (a
  name for its figure, a vertex for its polygon; not boxes, axes or `Auxiliary`
  helpers), a click on a row of the Figure List too, the only way to a hidden figure, which
  then shows as a ghost (`IFigurePicker`: a row click picks instead of selecting, Space
  picks the keyboard's row, Delete does nothing meanwhile; the picks are selected figures,
  and `Drawing.PicksChanged`, not `SelectionChanged`, repaints the list, or the side panel
  would trade the tool's panel for the selection's); a click on the paper puts the box
  there, ticked if any of its figures shows ("Show"), else unticked ("Hint"), so that
  nothing changes on screen, selected with the keyboard in its Caption. The box's Edit
  figures (grid, context menu) starts the tool on it (`ShowHideCreator.Edit`, the ribbon's
  instance: `Behavior.FindTool`; the ribbon, following a tool a verb started, makes its tab
  remember it before opening the tab, `Ribbon.SelectBehavior`, or the tab clicked the tool
  it remembered and stopped the editing); the paper, the box, Enter or OK ends it, one undo step
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
  typed one on other points). An input is asked for as any point, line or ellipse unless
  a result takes more from it (`MacroSerializer.TakesMoreThanKind`: its length, area,
  inside) or an expression of a result reads a number off it that not every figure of the
  kind has (u.X of a vector, Q.Parameter; a point's X and Y every point has): then only
  its own type. It asks for its inputs in the order
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
