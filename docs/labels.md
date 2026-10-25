# Names and labels

Default names, subscripts, labels that say numbers, wrapping, names of lines and circles, attached texts, names keeping clear, pinned labels, editing a text in place.

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
  "Midpoint E", "Hexagon ABCDEF"), the kind left off when the name says it already
  (Circle1, Bezier3), the name left off when it is the type and a number on a figure
  nothing refers to by name (Distance, Text, Angle mark, Locus: `NamedByConstruction`) -
  not on others: "on Circle₁" in one row must be found in another.
  After the title comes `Construction`, how the figure is built ("of AB", "to segment AB
  through C", "with center A and radius 3"): faded after the title in the Figure List, a
  smaller line under it in the grid's header (`ConstructionOnPage`: not on the page of a
  text or a show/hide box, where the Text box says it all). It names what it is built on through
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
- **A label is a length and an angle only when it says a number** (`Label.IsNumber`: its
  text is one, usually the value of an expression, "[AB * 2]" - what the GeoGebra reader
  makes for a computed value). Every `Label` has `ILengthProvider` and `IAngleProvider`,
  so wherever a tool or a tied value takes "a figure with a length" or "with an angle" it
  asks `Label.GivesNumber` too (By Radius, Rotate, Dilate, Translate, the
  `Accepts` of the tied values, `FigureCreator.LookForExpectedDependencyUnderCursor`): a
  caption, taken for 0, tied a circle's radius or a rotation to a piece of text with one
  stray click. An expression without a value shows "undefined" (`LabelBase.UndefinedText`),
  not "NaN" (nor "Infinity"), and such a label's `Value` is still not a number. A label
  that is one expression and nothing else gives that expression's value, not the text it
  shows (`LabelBase.ExactValue`): `[AB / 3]` as a radius was 0.33, and changed with the
  label's Decimals. A circle whose radius comes out negative or undefined (a label, a
  Number) doesn't exist (`CircleByRadius.UpdateExistence`). A label's text set (the grid,
  undo, redo) recalculates what is built on it (`LabelBase.Text`); a file read or a
  re-binding doesn't come through the setter. Without it, a circle and a point on the
  label's value kept the old one until the next recalculation, and undo left them there.
- **Wrapped text keeps an expression on one line** (`Figures/Controls/LineBreaks.cs`,
  2026-10-17): a label with a wrap width (a caption) shows each space next to a word that
  is all operators (`= ≠ ≈ < > ≤ ≥ + − - ± × · ÷ * /`) as a no-break space, so "AP = 8 − r"
  no longer breaks into "AP = 8 −" and "r."; a chain of more than `LongChain` (24)
  characters may still break before its second and later relations ("x² = AP² − AN²",
  then "= (8 − r)² − r²"). Only on screen (`LabelBase.ShownText`, which `Label` overrides):
  the text keeps the spaces it was typed with, and nobody types `&#160;` into a caption any
  more. A chain with no second relation that is wider than the line is broken inside, as
  any word that long. The player's twin is `lineBreaks.js` (`Label.shownText`).
- **Names of lines and circles** (`Figures/Controls/FigureLabel.cs`): "Show name" on a line,
  ray, segment or circle (`LineBase`/`CircleBase.ShowName`, over `FigureBase.HasNameLabel`)
  adds a `FigureLabel` the way a point's name is a `PointLabel`: a label depending on the
  figure, `Offset` pixels from an anchor on it (`FigureAnchor`, which a text attached to a
  figure shares) - a segment's middle, the upper left of a
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
  **One family of labels** (2026-10-18): `PointLabel`, `FigureLabel` and `AttachedLabel` are
  `AnchoredLabel`s (hung on their first dependency through `FigureAnchor`, keeping clear,
  hidden with the figure - but a point's name, which the clock face's numbers and the
  like are - and kept on screen beside a line); the first two are `NameLabel`s, part of
  their figure (`FigureBase.NameLabel`, a point's too, with `HasNameLabel` and its retired
  label), and whatever leaves names out (the Figure List, Delete, a region, a show/hide
  box, a formula, the context menu) asks `is NameLabel`. Every `LabelWithOffset` is placed
  once the same way: an `Offset` set from anywhere (a file, a drag, undo, a reader) is its
  place (`Placed`); one without gets `DefaultOffset()` at the first redraw with a sized
  canvas outside a file read (`FirstOffset` where it keeps clear), and a file without
  `OffsetX`/`OffsetY` leaves it unplaced. A reader that moves a new label on from its
  default calls `PlaceIfNew` first. A point's name goes over to a replacement like any
  dependent (`SubstituteWith` no longer skips it), then `Actions.HandOverNameLabel`.
- **Texts attached to figures** (`Figures/Controls/AttachedLabel.cs`, 2026-10-17): the
  Text tool clicked on a point, segment, line, ray, vector, circle or polygon (inside, where
  it is filled; on its sides a click takes the side segment) makes an `AttachedLabel`, a
  text that hangs where a name would (`FigureAnchor`: a point itself, to its right at
  first; a segment's middle, beside it; a polygon's centroid by area, `FigureAnchor.Centroid`,
  centered on it, where a new area number goes too - it hung its corner at the average of
  the vertices, which lies outside an L; one read from a file keeps its offset, from the
  centroid now, the same point for the squares of Pythagoras) and moves with its figure, but says
  whatever is typed, so that a name isn't made into a comment (a segment named r to show
  "r"). It is a figure of its own - in the Figure List, deleted by Delete, gone with its
  figure, hidden with it - edited in the grid's Text box, a plain one (not the free
  label's, which puts references to figures in [...]: the text is plain, since [...] parts
  would make it depend on the figures they name, and it depends on its figure alone) or by
  a double click (`InPlaceTextEditor`). File: `<AttachedLabel Text OffsetX OffsetY>` with
  its figure as the dependency; a file written by hand may leave the offsets out, and the
  text goes where a new one would once the drawing is read. The player has the twin
  (`attachedLabel.js`, `figureAnchor.js`). Detach (grid, context menu) makes it a free
  text where it is drawn, and a text dragged with Alt goes onto the figure under the
  cursor at the drop, or, attached, lets go of its figure over the paper, as a point
  snaps and is released (`Dragger.AimText`: the halo shows the figure, the status says
  what the drop does; nothing changes before it). Both are a swap of one label for the
  other in place, style, Visible and Locked along (`TextAttachment`), inside the drag's
  undo step. A free text with [...] parts stays free (with Alt the status says why: it
  would show them as typed), and so does a pinned caption.
- **Names keep clear** (`Figures/Controls/NameLayout.cs`, 2026-10-17): a point's name, a
  figure's name and an attached text are drawn where their offset puts them unless
  something covers that place - another name, measurement or angle's number overlapping it
  by more than `HomeDepth` (2 px) both ways (boxes have room around the letters: the Clock
  Hands' numbers touch on a phone and stay put, but 10 and 11) or another point. Then they
  are drawn at the nearest free place around the anchor, `Spacing` (6 px) from the other
  names (at 2, A and B side by side read as AB), steps of 15° each way at the same
  distance, then a ring further out, not across the lines and circles through the anchor
  (a point's own, `FiguresThroughAnchor`; a segment's name, its segment) and clear of their
  own point. Labels keep their size in pixels at every zoom, so names apart on a big screen
  met on a phone. Only where they are drawn changes (`LabelWithOffset.Nudge`, not saved):
  the offset and the file stay as the user left them, so a drawing looks as it did
  wherever nothing collides. A name stays where it went while that is free (no flicker
  during a drag), and goes back once its place is free: every label drawn elsewhere than
  before has the names drawn away look again (`NameLayout.Recheck`). A name being dragged
  is drawn under the cursor whatever it covers (`movedByHand`). A new name's first place
  is the nearest free one to its default that no line through the anchor crosses
  (`NameLayout.FirstPlace`: a point's name above it, but not on a vertical line through
  it); after that, its place is the user's. Measurements and angle numbers are obstacles,
  not moved. The player has the twin (`nameLayout.js`).
- **Pinned labels** (`Label.Pin`, `Figures/Controls/LabelPin.cs`): `Pin="TopRight" OffsetX
  OffsetY` is the pixel distance from a canvas corner to the *same* corner of the label, so it
  keeps that edge while the plane zooms and pans under it. `Coordinates` always tell where it
  is in the plane right now, so hit testing, dragging and the property grid need nothing
  special. `WrapWidth` (px) exists for captions, and so does `Backdrop` (a plate of the
  paper's color, "Background" in the grid; selected, the plate takes the selection's tint,
  which it covered) - since 2026-10-23 on every label (`LabelBase`), so a measurement or
  a name can sit on one over a busy spot; a free text writes the attribute after its own,
  as the gallery's files have it (`WritesBackdropLast`). Pinned labels draw above figures,
  zoom to fit ignores them, thumbnails hide them.
- **A label's text is edited where it is** (`Behaviors/InPlaceTextEditor.cs`, 2026-10-17): a
  double click on a text label with the Drag tool (not a locked one, nor a gallery
  drawing's own text) puts a box over it in its font, size, padding and wrap width, so the
  lines break where they did (but for an expression `LineBreaks` keeps together: the box
  holds the text as typed, and breaks at any space), the label hidden under it and the
  caret where the click was.
  The label is the one the first click pressed (`Dragger.lastPressed`): that click
  selected it, and the view may have moved since to bring it out from under the side
  panel. Enter or leaving the box (a click elsewhere, another tool) puts the text in, one
  `SetProperty` undo step; Shift+Enter starts a line; Escape (`Dragger.AbandonGesture`,
  which `MainView` asks first) leaves the label as it was. The side panel is away while the
  box is up and comes back with the label after; the view and the panel's side stay as
  they were meanwhile. The box follows the label when the view moves under it.
