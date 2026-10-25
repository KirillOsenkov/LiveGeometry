# Figures

How figures are built and worked out: the dependency graph, crossings, regions, squares and regular polygons, lengths, ellipses, arcs and angles, curves, vectors, marks, parts of a figure.

- **Crossings with curves** (`Figures/Points/Intersection/CurveIntersections.cs`,
  2026-10-15): a function graph, a Bezier curve or a Bezier path crosses a line, ray,
  segment, vector or axis, a circle, arc or ellipse, or a graph (graph with graph, Bezier
  with graph; not two Beziers, no loci). Any number of crossings, so the `IntersectionPoint`
  has no `Algorithm` but a `Parameter` (x on a graph, t on the cubic itself, piece + t on a
  path; saved as `Parameter="3.16"` instead of `Algorithm`), and each recalculation takes
  the root nearest it of the other figure's equation along the curve (a signed distance;
  `y - g(x)` for a graph) and moves the parameter there: the point follows its own root.
  The search walks out on both sides in steps growing by 1.1, halves a sign change, and
  takes a minimum of next to nothing for a touch (y = x² on the x-axis); a sign change
  across a pole or a jump is no root. A graph is not hit-tested (its samples end at the
  window); a ray, segment or arc is. The hover and the click seed the parameter from the
  cursor, the GeoGebra reader's `Intersect`, `Root` and `Roots` from the saved place.
  Tracking depends on the way the figures went, so undo would put two crossings onto one
  root after a long drag: `CurveIntersectionHistory` keeps every crossing's parameter for
  each state of the history (by the last undoable action) and puts them back on undo, redo
  and a rolled-back construction. The player has the twin (`curveIntersections.js`), and
  `Main/Player/dev/kinds.lgf` has crossings with a graph and a path.
- **A point inside a shape** (`Figures/Points/PointInRegion.cs`, 2026-10-11; GeoGebra's
  "Point on Object", `PointIn[]`): a `FreePoint` by inheritance, like a point on a figure,
  with one dependency, an `IRegion` (`Figures/IRegion.cs`: `HasInside`, `Area`,
  `ContainsPoint`, `NearestPointInside` - on `PolygonBase`, `EllipseBase`, the four sectors and segments through helpers on
  `EllipseArcBase`, and `BezierPath`, closed or filled, holes left out; a circle's nearest
  point is along the ray from its center, a path's is on the edge of its inside as drawn,
  holes taken out: `Math.GetNearestPointOnRegionEdge`, the pieces of every side, cut where
  the other polygons cross it, with the region on one side and not the other - judged by
  a point just off each side, since a middle on another polygon's outline is inside it or
  not by chance, and the seam of two holes was taken for an edge; the nearest point of the
  whole outline could be in a hole that reached past it). A sector or segment with such
  a point offers no Convert to arc (`IsUsedForArea`). Its state
  is an `Offset` from the region's `Center` (saved as `OffsetX`/`OffsetY`; a hand-written
  file with X and Y alone gets it at the first recalculation): `Recalculate` works the
  point out from the region as it is now, however it got there - dragged, driven by a
  slider, the image of a transformation - so a polygon dragged carries the point, a vertex
  dragged moves it by a third, a turned piece moves it with the center but doesn't turn it.
  Where the place falls outside, the point sits at the nearest point inside and the
  offset keeps the wished-for place, so undo of the change, or the shape growing back,
  gives the point its place back; only a drag of the point (`MoveToCore`, which clamps)
  or the file sets the offset, and `CapturePlace`/`RestorePlace` carry it. It doesn't
  exist while the region doesn't or has no inside. Its X and Y are not typed (read-only,
  as a point on a figure's; decided 2026-10-12): a place typed outside went to the edge,
  the other coordinate with it. It takes the `PointOnFigure` green,
  "Free point" releases it, it is no end to stretch or fix (`LengthConstraint.CanStretch`),
  the Dragger records where it lands as for a point on a figure, and the player has the
  twin (`pointInRegion.js`, `isRegion` on the same classes). The tool is
  `PointInRegionCreator` (a `Behavior`, as the Point tool): `PointInRegion.FindRegions`
  lists the visible regions containing the click by area, each a
  `PointPlacement.InRegion`, so the preview, the choice and the status come for free.
  Why an offset and not a Dragger rule: a drag of a triangle moves A, then B, then C, each
  with its own recalculation, so a point that watched its parent for a translation never
  saw one; and a rule in the Dragger would miss every move that isn't a drag.
- **Regions** (`Figures/Regions/`, 2026-10-23): a `Region` is a set made of figures with an
  inside (`IRegion`: polygons, circles, ellipses, sectors, circular segments, Bezier paths,
  other regions) by an `Operation` - Difference (the first without the others), Intersection,
  Union - with no edge of its own: the inside alone, filled (the palette's `BlueShape` by
  default, so that it is told from the yellow polygon it lies over; an outlined style
  draws its true edge), in the Polygons layer, `<Region Operation="Union">` with its
  operands as dependencies. Disjoint parts are one region (a union of two circles apart);
  a set that comes to nothing (an intersection pulled apart) exists with area 0 and draws
  nothing, and a point inside it doesn't exist. An operand that doesn't exist (a sector
  whose radius came to 0) is the empty set: a union goes on without it, a difference
  without that subtrahend; a difference whose first operand is gone, or an intersection
  missing any, is gone (`Region.UpdateExistence`; the goat's reach around a corner is
  such a sector until the rope is long enough). Everything comes from one tracer,
  `RegionBoundary.Trace`: every operand's boundary in pieces (`BoundaryPiece`: segments,
  and arcs of ellipses in at most quarter turns - a circle is an ellipse, a sector is its
  arc and two radii, a path is the polygon through its samples as `BezierPath.Area` has
  it, a nested region its own kept pieces; a circle's inclination, which a sector or a
  segment has from its first point, is folded into the angles, or an arc of a circular
  segment and an arc of its circle were taken for different ellipses and the search for
  crossings between them cut slivers at random - the yin-yang flashed straight edges while
  it was spun), cut at every crossing (segment-segment,
  segment-ellipse and circle-circle through `Math`, with its tangency tolerance, so a
  touch by construction is one point and no crossing; an ellipse against another arc by
  sampling and bisection; two stretches on one line or circle cut at each other's ends),
  each piece turned so that its own operand is on its left (judged a hair off its middle,
  inside and out: a self-crossing polygon's even-odd lobes come out right), and kept when
  the operation holds on one side of it and not the other - the other operands judged at
  the piece's middle itself, which lies on no other boundary once the pieces are cut, but
  for a stretch two operands share, which goes by the other's side of it (`IsSameAs`);
  two kept copies of one stretch are one. The kept pieces, the result on their left, give
  the area by Green's theorem (exact for segments and arcs: `AreaTerm`, taken from a
  nearby origin), the perimeter, the loops to draw (`Chain`, chained within 1e-6 of the
  scale where the pieces are cut within 1e-9: a sliver too short to be judged leaves no
  gap that a straight line would close; any pairing where several meet will do, even-odd
  goes by the pieces), the centroid (`Center`, what a point inside
  keeps its offset from) and the nearest point of the edge. `ContainsPoint` is the
  operation over the operands' own. Checked by the regression suite on the classic
  puzzles (the square and its inscribed circle, the lens, the Reuleaux triangle, an
  annulus, a notched square, a bow tie), on the identities every pair obeys (union plus
  intersection is the sum, for tilted ellipses too) and on a circle dragged through a
  side, touching from outside, crossing and touching from inside, where the area must not
  jump. A region is transformed through its operands (`Transformer`: a region on their
  images; not in a circle), measured by the Area and Length tools, read in expressions
  (`r.Area`, `r.Perimeter`), and holds its operands' kinds (`IsUsedForArea`: no Convert to
  arc under a region). The player has the twin (`regions/regionBoundary.js`, `region.js`;
  `Main/Player/dev/kinds.lgf` has three), and `playerparity` compares areas and perimeters.
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
- **The figure list is in dependency order**: a figure comes after what it is built on. The
  file is written in the list's order and read back dependencies first, so a list out of
  order comes back in another order, and `Drawing.Recalculate` goes down the list. What
  breaks it is rewiring: a join (the target may be later than what gets built on it) and a
  tool that records a figure while its last point is still the one following the cursor
  (the Angle tool added its arc with the preview; both figures are made at the end now).
  `Actions.MoveBefore` is the repair, and knows parts (`RootFigureList.FindTopLevel`).
- **A walk over dependencies visits each figure once** (a visited set, or an answer kept
  per figure): the graph is a DAG full of diamonds, and a walk that goes down every path
  takes a time that doubles with each link of a chain of them. Numbers worked out from
  numbers (a ray traced through a prism, 2026-10-24) hung the window on loading in
  `DependsOn`, the cycle check of every name an expression compiles, and took seconds at
  each press in `DependencyAlgorithms.FindRoots`. The player's `dependsOn` is the same.
- **The figure list checks itself in Debug builds only** (`CheckConsistencyInDebug`: after a
  click of a tool, a recalculation of what is built on a figure, a file read): every
  dependency and dependent of a figure is in the drawing and registered both ways. A
  Release build (the site) never runs `CheckConsistency` on its own, so a registration bug
  shows there only by what it breaks later: catch them with the Debug desktop build, the
  regression suite and the harness, which call the check themselves.
- **Squares and regular polygons are plain figures** (`Figures/Shapes/RegularPolygonBuilder.cs`,
  2026-10-13): the vertices they add are points by coordinates written over the two points
  clicked - named like any point, so expressions say them and figures are built on them -
  then a segment per side not drawn already (`FindLine`), and a `Polygon` on all the
  vertices. Square and Polygon on a side take two adjacent vertices, the polygon to the
  left of the first to the second; Regular polygon takes its center and a vertex, the
  others following counterclockwise. A vertex is P + R(Q - P), its X and Y sorted by point
  name, X before Y, without parentheses (the square's C is `A.Y + B.X - B.Y`,
  `-A.X + B.X + B.Y`), its coefficients exact where the polygon tiles - halves and halves of
  `sqrt(3)` for 3, 6 and 12 sides, of `sqrt(2)` for 8, whole numbers for 4 - and 15-digit
  decimals otherwise. What the drawing has is taken, never what is there by chance: a
  vertex is a point the drawing shows when the two are the same point however the figures
  move (`RegularPolygonBuilder.FindSamePoint`: their X and Y worked out exactly through the
  points by coordinates they name come out equal, `LinearForms`, see "Exact forms"), then
  the sides between such vertices (`FindLine`), and a polygon on the same vertices is no
  second polygon ("Square ABCD is there already", no undo step). In a lattice of squares
  grown on each other's sides the fourth corner of a square is the third of its neighbor,
  and a fan of squares round a point closes on the first one's free corner: one point
  each, one segment per side (2026-10-13; the regression suite grows a 2 x 2, a 3 x 3, a
  fan, six hexagons round a hexagon and six triangles round a point). A free point at the
  place, one typed there or a corner of a square on other points stays a point of its own.
  A deleted side takes only its segment. Until then a square was a
  circle, a perpendicular, their crossing, a midpoint and a reflection per square, and a
  regular polygon a composite whose other vertices were unnamed parts nothing could name.
  The two regular polygon tools share a Number of sides (the pre-create `SidesPanel`,
  `RegularPolygonCreator.NextNumberOfSides`); right after one is made its own panel
  (`RegularPolygonPanel`) has the number of sides - a new count builds the polygon again,
  what was made for it removed and new figures added in one transaction recorded by the
  row itself (`ISelfRecordingValue`: a property set can't record actions), refused while
  something else is built on what was made - the side (moving the second click's vertex)
  with Fix length and Free length, and OK, which brings the tool's panel back. After that
  the polygon is a plain polygon, with no count to change. The Square shows its base side's
  `DimensionsPanel`, as other figures with a length do, and so does Polygon on a side
  (2026-10-13), with the number of sides on top and the polygon's title and OK
  (`DimensionsPanel.RegularPolygon`, `RegularPolygonPanel.Shown`): the side's length and
  direction, Fix, Horizontal, Vertical. A Show ticked there is built on the side, so a new
  count is refused until it is unticked. Those verbs may replace either point clicked (a
  side the drawing had may run the other way, and then its end is the first click), so the
  panel reads both off the polygon: a new count built on the first click as it was put the
  polygon on a point that had left.
- **Lengths** (`Figures/Controls/LengthMeasurement.cs`, 2026-10-15): one figure,
  `LengthMeasurement`, for every length - the distance between two points, or one length
  of one figure, which one being its `Quantity` (the grid's Measure row): Distance (two
  points), Length (segment, vector, polyline, Bezier curve, open Bezier path), Arc length,
  Radius (circles, circular arcs, sectors, segments), Diameter (circles), Circumference
  (circles, ellipses), Perimeter (polygons, sectors, circular segments, closed paths: a
  path's one quantity follows its Closed, so its number is its length open and its
  perimeter closed). `QuantitiesOf` lists what a figure
  has, the first being what the tool's click gives (a circle's radius, a sector's arc
  length); the row offers only those (`IOffersValues`, which the enum editors ask, so a
  circle's list is three long) and is read-only where there is one. It replaced the
  Distance and Perimeter measurements and tools, and a file says `Quantity` only when it
  is not the figure's first (`<LengthMeasurement Quantity="Circumference">`). A quantity
  the figure doesn't have, or one without a value (an open path's perimeter, a segment
  turned into a ray), makes it not exist. Right after the tool a figure with more than one
  length gets `LengthMeasurementPanel` (the Measure row and OK), as an arc gets its Sweep
  row. A prefix comes with the quantity ("r = ", "d = ", "C = ", "P = "; nothing for a
  distance, a length, an arc length, which sit on what they measure), so that a radius, a
  diameter and an area of the same circle are told apart; a typed one (`Prefix`, saved,
  `""` for none) stays whatever the quantity becomes. Where it sits: a distance or a
  length at the middle with the file's offset (the default offset is 0, as it always
  was; a polyline's, a curve's or a path's middle is the middle of its parameter, on it,
  not its center); the rest are placed once, at the first `UpdateVisual` with a canvas that has a size
  (and again when a quantity is picked in the row; a file without an offset gets it too).
  The place is decided at that first redraw whatever the quantity, so a quantity that
  changes by itself (an open path's length becoming its perimeter when it is closed)
  doesn't move the number; it did, outside the undo step and only until a reload. The row
  is `LengthMeasurement.QuantityValue`, an `IRestorableValue`: undo puts back the quantity
  and the offset (a number dragged away, then given a new quantity, came back elsewhere).
  The Prefix row is `PrefixValue`, restoring the prefix as kept (typed, or null for the
  quantity's own): through the setter, a typed prefix equal to the new quantity's own came
  back as that quantity's, and went with the quantity on the next undo.
  Default places: an arc
  length at the arc's middle and pushed away from the center, a radius at the middle of
  the radius to the point a circle goes through (an arc's start; else the one to the
  right) and beside it, a diameter under the center, a circumference or a perimeter on the
  outline - the middle of a polygon's first side (pushed across it away from the inside),
  the lower right of a circle or an ellipse (the name's corner is the upper left), the
  middle of a sector's arc, the middle of a path's first piece. A circle or an ellipse
  whose fixed point is off screen anchors at its point nearest the middle of the window,
  which moves with the view as a line's name does. `IPerimeter` (`Perimeter`) is on
  `PolygonBase`, `EllipseBase` (the arc integral `Math.EllipseArcLength` over a whole turn,
  which `ArcBase.Length` also uses), the sectors and circular segments (the arc with the
  radii or the chord) and `BezierPath` (closed: its `Length`, the cubics' lengths by
  Gauss-Legendre, `BezierInfo.Length`; open: NaN). `Bezier.Length` and `BezierPath.Length`
  are plain properties, not `ILengthProvider` either (other tools would take a curve for a
  radius). Not `ILengthProvider`, on
  purpose: a circle's `Length` is its radius (`IFixableLength`), and `c.Length` in an
  expression reads that; an expression says `c.Radius`, `c.Diameter`, `c.Circumference`
  (= `c.Perimeter`), `s.ArcLength`. The measurement itself is an `ILengthProvider` (a
  circumference as a radius). Only a perimeter measurement keeps a sector from becoming an
  arc (`IsUsedForArea`); only a length of a segment is unwrapped into the segment's two
  points by By Radius. Angle marks are arcs in code and are left out, as are numbers and
  labels; lines, rays and function graphs go on without end and loci are sampled by the
  window, so a click on one is a point on it, the end of a distance (the arrow cursor of a
  point on a figure). The GeoGebra
  reader maps `Distance[]`, `Length[]` (the figure's first length, but a circle's or an
  ellipse's circumference, as GeoGebra means it), `Perimeter[]`, `Circumference[]` and
  `Radius[]`; an expression or a text naming such a result reads `r.Length` (an area's
  `a.Measure`, `ValueExpressionOf`): bare, the name was no number to our language, and
  what was built on it kept the worksheet's value.
- **Ellipses**: center, end of the long axis, end of the short axis, the last by its distance
  from the long axis, so a point on the short axis is on the ellipse. The tool puts a free third
  click on a hidden perpendicular through the center (`EllipseCreator.FindPointPlacement`), so
  scaling the long axis scales the short one with it.
- **Arcs, sectors and segments of circles and ellipses** (`Figures/Circles/ArcBase.cs`) get
  their areas from the sweep t in the angle that parametrizes the ellipse (x = a cos t,
  y = b sin t; the central angle on a circle): sector a·b·t/2, segment a·b·(t - sin t)/2;
  the length of an elliptical arc is Simpson's rule over t. They said r²·angle/π (a quarter
  disc of radius 2 had the area 4), and "not a number" for every ellipse - so did an Area
  or length measurement on one, and a circle given such an arc for its radius. Convert
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
  rotation goes (GeoGebra's rule; `directedangle(A, B, C)` in an expression stays the bare
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
  steps, then round after round every step halved where its middle sample strays more
  than half a pixel from the straight line between its ends, or a gap begins or ends -
  at most 12 rounds and 1000 samples, spread evenly when they run out (sin(40x) is
  coarse, not stuck). The middle is drawn either way, so what is drawn strays about an
  eighth of a pixel (2026-10-16: thrown away, it left the inverse of a circle 600 px
  across with a corner every 35 px; a tolerance of an eighth cost 1.6 to 2 times the
  samples for the same look, the middle nothing). A piece still long (20 px) after every
  round is a jump (1/x), not joined; so that a jump that is not through infinity is found
  too, a piece that long is never done while its middle sample lies within a tenth of
  its length of an end (`Locus.IsLopsided`). Spirograph's sliding point drives an angle
  that wraps from 2π to 0 at the end of every lap: the middle landed next to one end,
  on the straight line, and every lap ended in a straight line back to its start, a
  triangle over the petal at the right (until 2026-10-17; no other gallery locus takes a
  sample more for it). Past the open ends of a line, ray
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
  own test (shown parts only) they all went when the vector was hidden. Its X and Y
  (2026-10-12) are the second point less the first, rows of its grid and of the panel after
  the tool, not of a selection of several (`[PropertyGridDisallowMultiEdit]`): typed, they
  move the end the length and direction rows move, and a translated end (fixed length or
  direction) gets its Numbers or free parameters set to match (`Vector.SetComponents`); an
  expression reads `u.X`, in the player too. Segments have none, on purpose. "Convert end to
  point by coordinates" (grid, a free end) makes the end `A.X + 3`, `A.Y - 1.5` - the
  components as the grid shows them, so the end may move by a rounding - which then follows
  the start; offered only where both texts compile without a cycle (`EndCoordinateTexts`).
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
  **The number places itself** (2026-10-16; `AngleLabelLayout`, `AngleMeasurement.PlacesItself`,
  the grid's "Place automatically"): at every redraw, for all the numbers at its vertex at
  once, each inside its angle past its mark, clear of the others and of the sides where
  there is room (a few turns off the bisector, a few pushes out; the cheapest of overlap,
  the length of the sides through it, its middle outside the angle, the turn, the push).
  It remembers its place relative to its angle (`LayoutPlacement`, not saved) and starts
  from there: it slides by small steps, and takes a place further away only when that is
  cheaper by `Hysteresis` and `Proximity` per pixel of the way. Chosen afresh at every
  redraw, the numbers jumped between two places at every move of a drag (measured in the
  player: 44 jumps over 12 px in a sweep of a side through 110 degrees and back, 1 now). A drag makes it
  a number placed by hand, which the others keep clear of; undo of the drag, or the row
  ticked, makes it place itself again. A file says `OffsetX`/`OffsetY` only for one placed
  by hand (`LabelWithOffset.SavesOffset`); every file from before says them, so the
  gallery's stay as they were. The Angle tool's hover ghost goes where the layout would put
  the number. The player has the twin (`angleLabelLayout.js`). Under the Drag tool, hovering
  a number lights up its mark (`Dragger.GetFigureToPick`), and any tool that picks a number
  lights its mark with it (`ClickPreview.WithAngleMark`): not with the faint halo band, which
  was hard to see around a mark this small, but a copy of the mark drawn over it 6 px wide
  in strong cyan (`ClickPreview.MarkHighlightBrush`, the hue of the coincident line), with
  the mark's own flat ends (round caps stuck out past the sides).
- **Parts of a figure** (`IFigureParts`, `Figures/Shapes/FigureParts.cs`): the sides,
  handles and inside of a Bezier path are figures other figures can be built on (only the
  images of a transformation are), but not figures of the drawing: unnamed, written as
  `<Dependency Name="p" Part="In3" />`. They are listed with what they are built on only
  while the path is in the drawing: registered when read, before it was added, they fail
  the consistency check. Whatever takes a figure from a click has to cope with one that
  has no name and is not in the drawing's list (`RootFigureList.FindTopLevel` gives the
  figure it belongs to). A side is selected by itself (`IFigureParts.SelectableParts`), to
  be styled; a click inside selects the path (`FigureParts.SelectionTarget`). The
  composite's `Selected` is the whole only, so whatever clears a selection asks
  `FigureParts.HasSelection`, and whatever acts on figures of the drawing (Delete, Copy,
  Hide, Lock, the context menu) takes `FigureParts.Wholes` of it. The file says the style
  most parts have and the odd ones (`<Sides Style="RedLine" />`, `<Part Name="Piece2"
  Style="BlueLine" />`). Regular polygons were such a composite until 2026-10-13, their
  vertices and sides unnamed parts (`RegularPolygon`, `DependentPolygonBase`); much of the
  code that copes with parts was written for them (`TiedValues.SourceName`, a translated
  point's `DistanceSourceIndex`, a macro's part inputs, `FigureCreator` leaving out the
  preview's parts) and stays for any composite.
