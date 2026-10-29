# Numbers, sliders and tied values

Lengths and directions set or fixed, Numbers, sliders, the typed or tied values of transformations, translated points.

- **Setting a segment's length stretches it once** (not a constraint); **Fix length** makes the
  constraint: the end becomes a `TranslatedPoint` from the pivot with an auxiliary `Number` at
  the current distance and a free direction (`IFixableLength`, `Figures/Lines/LengthConstraint.cs`;
  segments, vectors, circles; a regular polygon's vertex from its center or its first vertex,
  through its panel). The Number holds the length as typed when the end is still where the
  typed length put it, but for rounding (`TypedDimensions`, for Fix direction too: a typed 8,
  worked back out of the stretched end's coordinates, was fixed and saved as
  7.999999999999999). An end on a line that runs through the pivot
  takes the line as its direction instead (the Number signed along it, so the end is fully
  determined) and Free makes it a point on the line again; an end on any other figure can't be
  stretched or fixed at all. The **direction** of a segment, a vector, a ray or a line
  through two points works the same way (`IFixableDirection`, the `Direction` property in
  degrees - a vector's own `Direction` is radians, so its row stays `Angle` and the
  interface is implemented explicitly; `Figures/Lines/DirectionConstraint.cs`, 2026-10-12;
  an expression reads `AB.Direction`, the player through its `directionConstraint.js`):
  the row turns an end once, Fix direction makes the end a `TranslatedPoint` from the pivot
  whose direction is an auxiliary `Number` (distance free: it slides along the line through
  the pivot) - the same point Fix length makes, so the two are fixed and freed one by one
  through its Free distance and Free direction, and fixed both ways it is fully determined.
  Freeing the last one held makes it another kind of point, and so does fixing a point on a
  line, so either is refused (verbs, panels and the point's own boxes alike:
  `LengthConstraint.CanFree`/`CanFix`, `DirectionConstraint.CanFree`) while an expression
  reads off it what that kind lacks (B.Distance, B.Parameter). A verb that swaps a point
  for another kind (`Actions.ReplacePoint`; Convert to path anchor too) asks
  `IsReadByExpressionsBeyond` of the new kind first, unless the point is a plain free point,
  whose X and Y every point has.
  The end turned is the second if it can be, else the first (whose own direction is the
  figure's + 180°). Horizontal and Vertical fix the direction along the axis the nearer way
  (0° or 180°, 90° or 270°), one undo step. Lines built otherwise inherit the members and
  say no (`DirectionConstraint.CanEdit` wants two points; `LineTwoPoints.IsThroughTwoPoints`
  for a segment bisector, which is on two points), and hide the row: a parallel, a
  perpendicular, an angle bisector (whose `Angle` is its half angle, why the property is
  not `Angle`). Right after such a figure is made the side panel shows a `DimensionsPanel`
  (`FigureCreator.ShowCreatedFigure`): the length and the direction, whichever the figure
  has, each in its own box with its verbs; the tools themselves have no length box. A
  regular polygon has a panel of its own (`RegularPolygonPanel`, see "Squares and regular
  polygons"). The swap is `Actions.ReplacePoint`, which hands over the point's name label
  (`ReplaceFigureAction` skips it on purpose). A creator's undo transaction spans one
  construction, opened at the first click (`FigureCreator.EnsureTransaction`), so an edit made
  in that panel between constructions is an undo step of its own.
- **Numbers are figures** (`Figures/Values/Number.cs`): usable in expressions by name, a
  length and an angle at once. Since 2026-10-15 a Number is a `CompositeFigure` whose one
  part is a caption "a = 18.85" like a slider's (`Number.NumberCaption`, a
  `LabelWithOffset` anchored at the number's `Position`, which a drag moves whole,
  `IMovableParts`), and its value is typed (`Value`) or worked out by a formula over the
  drawing (`Expression`, a `DrawingExpression` as a point by coordinates has, the Formula
  row; `IExpressionOwner`, so renames follow): a Number with a formula is to a typed one
  what a point by coordinates is to a free point, and the Slider is a Number with a handle.
  The Number tool (Measure tab) makes one at the click, shown, named a, b, c from the
  sliders' pool, and opens its own panel (`NumberCreator.NumberPanel`: the Formula box
  with the keyboard in it, and OK), where what is typed goes into the number as soon as it
  compiles and OK, Enter, Escape or a right click are "done" (the tool goes down, the
  panel with it; the placing is the one undo step, the typing none); the helpers tools make (a
  typed distance of a translation, a fixed length's number) stay n1, n2 and hidden: a
  Number is hidden by default (`FigureBase.VisibleByDefault`, false for a Number: the
  file says `Visible="true"` for a shown one and nothing for the others, so every file
  from before reads as it did), and ticking Visible on one puts its caption at the middle
  of the window; hidden, it is a variable of the Figure List. Typing a Value into a number
  with a formula makes it typed again (`ValueState` restores the formula on undo); a
  formula that doesn't compile makes the number not exist, one without a value right now
  says "a = undefined". "Show formula" (2026-10-28, a row only while there is a formula)
  puts the formula in the caption: "a = angle(A, B, C) = 27". File: `<Number Name="a"
  Formula="d.ArcLength + f.ArcLength" X Y Visible="true" ShowName="false" ShowFormula="true"
  Decimals />`, `Value` instead of `Formula` for a typed one. The player has the twin (`number.js`). `Auxiliary` (any figure) marks one
  created on demand for another: it is removed with its last dependent and comes back on undo.
  Every deletion goes through `RemoveFigureAction`, one figure per action inside a transaction,
  so undo restores labels and a polygon loses a vertex rather than dying.
  **A polygon loses a deleted vertex** while three are left (2026-10-18; `Polygon`'s
  `ISupportRemoveDependency`), its images too (each a plain polygon on the images of the
  vertices, so deleting B takes B' and B'' out of the images, deleting B' leaves ABCD whole).
  The side segments along the gap go with the vertex, and a segment in the style of the
  first of them closes it (none where a line on the two points is drawn already, or the
  sides weren't segments). Not a tool's preview (`ExecuteImmediatelyWithoutRecording`): it
  goes with the point following the cursor, or it drew its closing side outside the tool's
  figures. `Drawing.Delete` deletes what is built on others first, so a polygon selected
  with a vertex goes whole and leaves no closing side. `RemoveFigureAction` works out what
  goes once and runs the same actions on redo, so redo brings back the same closing segment.
- **Sliders** (`Figures/Values/Slider.cs`) are one `CompositeFigure` whose parts are library
  figures: a `FreePoint` anchor, a `TranslatedPoint` knob kept on the horizontal through it
  (free distance, direction a Number saying 0, clamped at the anchor), the `Segment` track, a
  caption label "a = 2.00" a fixed few pixels above the anchor. The drawing and the file see
  one figure (`<Slider X Y Value>`), and no part is ever handed out: `HitTest` answers with the
  slider, so a tool can't come to depend on the knob. The value is the track's length:
  `INumber` (expressions say `a`), `ILengthProvider` (By Radius, Dilate, Translate),
  `IAngleProvider` in degrees (Rotate). A slider may have a range and a step (2026-10-10:
  `Minimum`, 0 unless said, `Maximum`, none unless said, and `Step`, 0 for smooth unless
  said; `<Slider ... Minimum="-10" Maximum="10" Step="0.25">`, the Range group of its
  grid): the value is the minimum at the anchor plus the track's length, the knob stops at
  the maximum and lands on a multiple of the step from the minimum
  (`Slider.ConstrainDistance`), also when the value is set from the grid or the file. The
  GeoGebra reader takes them from `<slider min max>` and `<animation step>` where those
  are plain numbers (GeoGebra lets them be expressions, left unread), in degrees for an
  angle slider. The gallery's sliders have maximums and no steps: the morphing between
  whole numbers is the point of Times Tables. Dragging goes by parts (`IMovableParts`, which the
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
  **What a Number is to its figures** (2026-10-22): the helper Number stays n1 (what a
  formula says), but its `Kind` and `Construction` say its role: each owner says it
  (`ITiedValues.DescribeValue`: "Rotation angle", "Translation distance" / "direction",
  "Dilation factor", "Angle" for a line at an angle, "Fixed length" / "Fixed direction"
  for a `TranslatedPoint` with the other quantity free, "Tension"), and
  `TiedValues.RoleOf` takes it when every dependent takes the number as the same value
  (none when an expression reads it too: then "Number"), so the Figure List reads
  "Rotation angle n₁ = 45° of triangle ABC about O", "Fixed length n₅ = 4 of segment DE"
  (the degree sign from `IsAngleValue`; the subject from `FigureTransformation` or the
  segment the end is on). A user's own number (a, from the Number tool) gets the role
  too once only a transformation takes it. **A figure that is the image of a
  transformation** (`Figures/Transformations/FigureTransformation.cs`): a plain polygon,
  segment, circle, path or traced locus on points that are all rotated, translated,
  dilated or reflected by the same sources (the `Siblings` test; a locus only when its
  sliding point is the Transformer's auxiliary one). Its construction line says it,
  "rotation of triangle ABC about O by 45°" (the default `FigureBase.Construction`;
  Circle, CircleByRadius, Ellipse, ArcBase and Locus put it before their own words), the
  source found as the figure on exactly the pre-images (`FindFigureOn`, a path also by a
  subset: its handles may be offsets), else the points by name when three or fewer. Its
  page gets the transformation's rows and buttons hoisted from one of its points
  (`PropertyGrid.CreateObjectControls`, after the Definition rows, before the
  quantities: `TiedValues.Rows`/`Buttons`, the same the `TiedValuesPanel` shows), since
  the points share their sources an edit through one moves the whole image. A value held
  by a hidden Number has "Show the angle on the drawing" (`TiedValues.ShowNumber`, one
  undo step: Auxiliary off, a letter name as the Number tool gives, Visible), there and
  in the panel after the tool. An image of an image shows the last transformation only.
- **TranslatedPoint** (`Figures/Points/TranslatedPoint.cs`): distance and direction are each
  *tied* to a figure (vector, length/angle provider, `ILine` as it points, a Number) or *free*
  (the parameter dragging changes); both are tied values (above), `SetSource` does the swap
  since a vector serves both at one index. Roles are stored by index into the dependency list, never
  inferred from types (a `Label` is both a length and an angle provider). A point with a free
  quantity takes the green `PointOnFigure` style. Both can't be free: ticking the second Free
  box makes the point a free point where it is (`FreedomValue.TryRecord`, through
  `ISelfRecordingValue`, which `Actions.SetProperty` asks first: a point swapped for another
  is an undo step of its own, and can't be recorded from inside the property set; the box
  was disabled until 2026-10-12). The Translate tool is stepwise (source,
  distance, direction, placement); its panel is cleared in `Stopping` because that event is
  raised before `Started` resets the state, and `DrawingControl` re-shows a behavior's
  `PropertyBag` on "construction complete" as well.
