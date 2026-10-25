# Bezier paths

- **Bezier paths** (`Figures/Shapes/BezierPath.cs`, 2026-10-06; added beside `Bezier` and
  `Polyline`, which stay as they were): a composite (see "Parts of a figure"). The anchors are
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
  which can't be dragged; an image isn't split nor an anchor of it dropped into another
  (`IsImage`), nor is a path with an image split or an anchor dropped from it (`HasImages`:
  the image would keep the old pieces). A deleted anchor leaves the source and every image
  (2026-10-18; the image of the anchor goes, and each image loses it as the source does),
  and an image's own anchor deleted leaves the image, the source untouched: the helper
  points of the lost anchor's handles go in the same deletion
  (`ISupportRemoveDependency.GetReleasedDependencies`, which `FindOrphanedAuxiliaries` takes
  as no longer used, also two images deep). The layout change runs before the helpers
  leave, so it recalculates without the Debug list check (`RecalculateAndUpdateUnchecked`;
  a polygon losing a vertex does the same). A deletion counts the
  anchors left after it, also those built on the point deleted (`CanRemoveDependencies`:
  `RemoveFigureAction` asks a figure that can lose dependencies about everything the
  deletion takes at once, so a show/hide box holding A and AB goes when A is deleted).
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
  "Smooth automatically" and "Sharp corner" on an anchor (grid and context menu of any
  point, `PointBase`: an anchor may be a midpoint or a crossing, whose grids had neither
  until 2026-10-17), "Smooth all anchors" on the path. Convert to
  path anchor between two automatic handles gives an automatic anchor (the curve moves a
  little), else the de Casteljau split. The tool's panel has Smoothing and Tension for the
  next path; the enum editor shows a value's `[PropertyGridName]` ("Catmull-Rom").
