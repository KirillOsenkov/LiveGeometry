# The canvas

The view, what a click, a drag or a key does on the canvas, the cursor, the order figures are drawn in, selection.

- **The view** (`Figures/Coordinates/CoordinateSystem.cs`) is an origin in pixels plus `UnitLength`
  (pixels per unit); every zoom goes through `Zoom(factor, focus)` (the point under `focus`
  stays put), every fit through `Fit`/`SetView`. "Content" for fit (`TryGetContentBounds`) is
  points, whole ellipses (a circular arc only its own extent), labels and the vertices of visible
  segments/polygons/Béziers even when their points are hidden - not lines. Panning (`MoveTo`)
  must not round the origin: a drag is many sub-pixel steps (0.5 px at 200% scaling) and
  rounding each one makes the plane run faster than the cursor. Labels have a fixed *pixel*
  size, so fit re-measures and refits a few times. What the user asks for as "zoom to fit"
  (double click, but on a text label, H, Home, the context menu) is `Drawing.ZoomToFit`: the drawing's
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
  and ticked it brought the axes up without the grid until the file was opened), and so are
  its lines (`ShowGridLines`, `Lines="false"`, 2026-10-18). A grid that shows shows at least
  one of the two: the menus (Grid's arrow, the paper's context menu) set them through
  `CartesianGrid.SetShown`, where hiding the last one hides the grid and leaves both on, and
  showing one of a hidden grid shows it alone; G and the button (`Toggle`) hide the grid
  and bring back what it showed. An axis is
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
  tool that takes a figure where it would otherwise expect a point (Length, Circle by Radius)
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
  regular polygon's sides, when it was a composite: it threw), and expressions that named the point name the target
  ([A.X] becomes [C.X]: `ExpressionRenamer` with the target under the point's name, once
  the point has left; undo puts the texts back as they were,
  `IRenamableExpressions.ExpressionTexts`) - so a join into a point without a name is
  refused when expressions name the point (no tool is given one any more). A point
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
- **Cursor philosophy** (`Behavior.GetCursor`): cross = a new *free* point appears here; hand =
  the click picks something already there - a figure the tool needs, an existing point, or a
  place defined by figures (intersection, midpoint); arrow = everything else, including a new
  point sliding along a figure and clicks that do nothing. `winauto cursor` prints the cursor
  showing now (screenshots don't include it).
- **Any figure is draggable**: dragging a dependent figure moves its root free points, so gallery
  text can say "drag the circle" even when the points it is built on are hidden. Labels are the
  exception (`AllowMove`): a drag moves the label, not what it measures or names. A point by
  coordinates is no root to move (`PointByCoordinates.AllowMove` is false, and `Dragger`
  leaves such roots out): the drag goes to the other roots, and a figure built on such
  points alone doesn't move, nor does the view. (They used to be "moved", to no effect but
  an undo step that undid nothing - every drag of the Ladder's wall.) A drag whose drop
  throws still ends its transaction (`Dragger.MouseUp`): left open, it took in everything
  done afterwards.
- **Tool letters are the only plain keys**: `Behavior.KeyDown` (what a tool without a key
  handler of its own gets: Point, Coordinates, Slider, Text) used to toggle "Label new
  points" on A and "Snap to grid" on G, from before those were the letters of the Arc tool
  and the grid: G with the Point tool on showed the grid and, silently, made every new
  point snap to it.
- **Z order** (`Figures/ZOrder.cs`, 2026-10-10): every figure is drawn in its kind's
  `Layer` (the `ZOrder` enum, bottom up: Grid, Axes, SelectionHalos, Polygons - polygons,
  the interiors of Bezier paths - Labels, Figures - lines, circles, arcs, curves,
  the sides of paths - Vectors, Handles, FixedPoints, ConstrainedPoints,
  FreePoints, PointLabels, Controls). Two bands of layers are ordered by the figure's `Z`
  (`ZOrders.IsMovable`), the Z first and the layer second: the figures' band, from Polygons
  to Vectors, and the points' band (2026-10-13), from FixedPoints to FreePoints - points and
  sliders, ordered among themselves. The three point layers go by what a drag does to the
  point itself, as the default styles do (`DefaultLayer` of each kind): free points (yellow)
  over those dragged along or within a figure (green: on a figure, inside a shape, a
  translated point with a free quantity, which changes layer as it is freed or fixed, and
  sliders) over those that don't move by themselves (gray: midpoints, intersections, images,
  points by coordinates - the corners a square or a regular polygon adds). All at Z 0, so a file says nothing
  of it; Bring to front gives Z 1, over the free points, Send to back Z -1, under the fixed
  ones. `ZIndex` is `ZOrders.Encode(Layer, Z)`: the layers under the figures' band (the
  selection halos among them) as they are, the figures' band with a stride of 1000 per Z,
  the handles over all of it, the points' band over those with the same stride (so a hover
  halo at a point's `ZIndex - 1` is still over the points of the Z below), the layers over
  it over all of it. Every figure's
  Z is 0, so the gallery draws as it did before there was a Z (checked pixel for pixel,
  `--check` before and after; the point layers changed it only where the Kaleidoscope's
  real emoji now cover their twins), and within one ZIndex the later figure is on top (Avalonia
  sorts children by ZIndex stably; the player sorts the list the same way). Bring to front
  sets the Z to one above the highest of the figure's band, Send to back to one below its
  lowest (`ZOrders.BringToFront`/`SendToBack`, a `SetProperty` per figure in one
  transaction, a selection of both kinds band by band: undo puts the old Z back); a
  polygon brought to front covers a circle, a measurement, a segment and its marks, a
  point brought to front the points it overlaps (and a click there takes it first: of two
  points the nearer wins only at one ZIndex, `FigureList.HitTest`), and nothing in the
  figures' band ever covers a point. The
  verbs are grid buttons on every figure (`FigureBase.BringToFront`, vetoed through
  `[PropertyGridCondition]`, which names a bool method of the object: the grid asks it
  where `IConditionalProperties` would have to be implemented by every figure), on a
  multiple selection (`FigureSelection`, all of them at one Z, their order among
  themselves kept by the list) and in the context menu (`Dragger`, for the selection as
  Hide and Lock take it); offered only when a figure of the band that is not among them
  is over (or under) one of them in the order of drawing - overlapping or not, as for the
  figures - or the step would be an undo step that changes nothing; a figure outside the
  bands (a label of a point, a pinned label) gets none. A swapped point
  (`Actions.ReplacePoint`) keeps the Z. A
  composite's parts take its Z (`CompositeFigure.OnZIndexChanged`, and on `Children` adds:
  a path's pieces), but only parts of its band
  (`ZOrders.PassOn`): a Z is a place in its own band, and a regular polygon (a composite
  then) brought to front among the figures would have put its vertices over every point, a
  slider brought to front among the points its track over every figure; a composite drawn
  by its parts alone says its layer itself (a Bezier path is Polygons), or the verbs would
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
  (`SelectionHalo.Layer`, `ZOrder.SelectionHalos`, under every figure whatever its Z since
  2026-10-15, so the stripes never cover the stroke or fill being edited; a filled shape
  over a selected figure hides that much of its halo. From 2026-10-10 the layer was over
  the figures' band and under the points', and the stripes lay over every line; before
  that each halo was under its own figure, hidden by one brought to front over it; the
  band of a region being selected is in the same layer) whose opacity the halos share, so
  overlaps don't darken. Its brush is the theme's `SelectionHalo` (a gradient edited as
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
