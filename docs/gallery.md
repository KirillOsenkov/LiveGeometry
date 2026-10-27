# Gallery

- **Gallery and routes** (`LiveGeometry.App/Gallery/`, `MainView` "Pages" region): three states -
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
- **The puzzle libraries** (2026-10-27): the "Gallery" heading shares its line with a chip per
  library of the site (`GalleryView.Libraries`: title, address, the chip's pastel, the count
  and a miniature), a gap to its right, where the room beside the heading was empty. A chip
  is a `GalleryTile` in its chip mode (a pill: the picture at the left, the caption beside
  it, the count as a faint `Badge`), so it has the tiles' plate, border and hover. The
  miniature of Catriona Agg's library is the three hexagons of her "Three Regular Hexagons
  and a Semicircle" (without the semicircle) drawn as vector shapes from the drawing's own
  construction (`HexagonsPicture`: the overlaps as `CombinedGeometry` intersections, the
  hole in the middle white; under the light theme the maintainer's fills with a hairline of
  ink on every edge, under the dark theme deep vivid fills and no ink, the colors telling
  the regions apart - `hexagonLightFills` / `hexagonDarkFills`); a mosaic of four colored
  cells was tried first and read as the Microsoft logo, and the semicircle was dropped as
  clutter at that size. A click opens the library's page through the launcher: a new tab in
  the browser head, the browser on the desktop. `HeadingRow` lays the line out: the chips
  in a line after the heading while all fit, else under the heading, one under another (a
  phone). A second library is one more line of the table and a miniature of its own.
- **Gallery drawings** are embedded `.lgf` (`Gallery/Drawings`, order and titles in
  `GalleryCatalog`) and are *the* source: edit them directly. They were forked once from the
  Windows Phone samples and the DG 1.0 CD library; `tools/gallerize.cs` made that fork and must
  not be rerun (it would overwrite the hand-edited drawings). To add another CD drawing:
  `--check` it (below), take the `.lgf`, drop the old labels, add `Title`/`Description` labels
  in the `Heading`/`Explanation` styles (defaults, so the file needn't carry them), set
  `Grid`, list it in the catalog, and `--rewrite` the folder. Drawings with `Grid="true"`
  (graphs) keep their file viewport in view (`GalleryItem.Plane`), because graphs and lines
  have no bounds. All 47 were rewritten on 2026-09-29 (`--rewrite`: loaded and saved again,
  the file's viewport and the caption's pin, offsets and width kept, since opening lays both
  out for the window at hand): no copies of defaults, no `Style` on a figure with its kind's
  default, no attributes at their defaults, default names of figures as the loader gives them
  (`AB` for `Segment8`). Rerunning `--rewrite` on the folder changes nothing, so a generated
  drawing (below) can be normalized the same way after regenerating.
- **The caption** (`Title` + `Description`, which may contain live expressions - then list the
  points as dependencies) is pinned to the screen. Its labels are the ones named Title and
  Description, else the first in the `Heading` and the `Explanation` style
  (`GalleryDrawing.FindCaptionLabel`, the player's `findCaptionLabel`; 2026-10-16: a label
  given the style and not the name was no caption, and a user had to do both). The files
  carry no X/Y and `GalleryDrawing.Fit` recomputes pin, offsets and wrap width at open
  time: a column at the
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
  the pentagon) or less (Cycloid, a strip), and a second, narrower scene crops a picture at
  the phone's sides (Treasure Island, Magic Tree, Fireworks; full screen picks the wide one).
  A phone scene smaller than the figure lets it run into the 40 px margins and fill the
  width (Honeycomb Facts, Cycloid); its shape must stay nearer the phone's room than the
  wide scene's does, and the wide one nearer the full screen's and the tile's (560 x 380),
  or the wrong one is picked. On a phone the caption lies over whatever the plane has under
  the figure, and a gradient paper gives it no plate: Design a Snowflake's snow fades into
  the paper's blue just below the scene, or its light text was white on white. `--check` of
  a single file lays it out with the default share, not the catalog's: check a phone layout
  with `--gallery <slug>` in a window `place`d to phone size. A point's own size limits
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
  angle is read in degrees (`[-30 * time]`; radians until 2026-10-18); the `Parameter` of a point on a line is
  the fraction from the line's first point to its second; a name's trailing digits draw
  as a subscript, so the Theodorus labels are "√17 " with a space.
- **The regions batch** (2026-10-23, five hand-written drawings at the end of the gallery,
  the classic shaded-area puzzles made with the region tools): Two Moons and a Triangle
  (`LunesOfHippocrates.lgf`: each lune is a circle on a leg without the circumcircle, C a
  point on the circle, the caption adds the two areas live), The Shoemaker's Knife
  (`Arbelos.lgf`: the big semicircle without the two small ones, all three `CircleSegment`s
  on the midpoints, hidden; D by coordinates as the geometric mean along the normal, so
  it needs no intersection), Yin and Yang (`YinYang.lgf`: three regions deep, the dark
  half out of a half disc, the two half-size circles and the eyes, the light half the
  circle without it; R turns and sizes it; the half disc's chord runs through R and the
  small circles' centers, which makes the S-curve - across them it was two stacked rings;
  the regression suite spins that wrong one too, since it is the harder geometry), Shaded Squares (`ShadedSquares.lgf`: the leaf
  as an intersection of two quarter discs, the flower as a union of four petals that are
  intersections of semicircles, the star as a square without four quarter discs, all over
  a slider for the side), The Grazing Goat (`GrazingGoat.lgf`: a union of three sectors,
  two of which have radius `max(r - w, 0)` and so don't exist until the rope reaches
  the corner; the goat is a point kept inside, an emoji), Arc Plus Arc (`ArcPlusArc.lgf`,
  the maintainer's: an S of two half circles over AD and DB as long as the half circle
  over AB, D sliding; the two colored parts are a half disc without one bump plus the
  other), and five more the same day, the ones first weighed and left out: The Ring and
  the Stick (`RingAndStick.lgf`: an annulus with a chord tangent to the hole, a circle on
  the chord beside it with the same area), The Salt Cellar (`Salinon.lgf`: the big half
  disc without the two small ones, in a Union with the middle one; the circle on its
  axis drawn as an outline over it, filled it hid the shape), Archimedes' Twin Circles
  (`TwinCircles.lgf`: the arbelos again, the twins by their closed-form radius r₁r₂/R
  in a hidden `Number` and centers by coordinates, so no tangent-circle tool), Counting
  Overlaps (`CountingOverlaps.lgf`: the seven pieces of three circles as regions, each
  an Intersection cut by a Difference, the three circles' sum against the pieces
  counted twice and three times, then inclusion-exclusion down to the union, which is
  the trivial reading and no caption line of its own), The Pizza Theorem
  (`PizzaTheorem.lgf`: a point kept inside the pizza, a knife emoji on the crust that
  sets the first cut - a slider in degrees was 90 units long, a slider's value being its
  track - the eight slices as the disc cut with wedge triangles reaching far out, and
  emoji toppings kept inside their slices; no olives, Twemoji's 🫒 is a sprig like the
  herb). Square Around Four Points (`SquareAroundFourPoints.lgf`, the same day, hand-written:
  the CD library's Problem 6.21.1 rebuilt - DF is AC turned a quarter turn, hung from D,
  and the line BF is a side; the construction lines are a faint `Guide` style, the two
  transferred lengths a blue `Transfer` one, and the square's sides `ThickLine` segments
  over the yellow polygon, since in the default line they vanished on the guide lines). A semicircle is a
  `CircleSegment` with the center at the side's middle, begun at the end the
  counterclockwise way round reaches the inside from (bottom side: from the right end to
  the left). A region's extent isn't content for the fit, so the lunes, the knife, the
  goat and Arc Plus Arc have a `<Scene>`. The flower's area number is offset onto its
  upper right petal (28 px up and right of the middle, where the four arcs crossed under
  the digits; below the square it fell off a phone's screen, and a plate of the paper
  behind it hid the middle of the flower). A point the reader is to drag must have its name's box
  clear of the point: a `PointLabel` is hit before its point (the labels' layer is over
  the points'), and C's name at `OffsetX="-4" OffsetY="-22"` covered C, so a drag on C
  dragged the name. Check with a hover under the Drag tool: the status must name the
  point first.
- **The optics batch** (2026-10-24, five drawings at the end of the gallery, all written by
  `dotnet tools/optics.cs -- <the Drawings folder>` and a `--rewrite`): The Upside-Down
  Candle (`UpsideDownCandle.lgf`: a thin lens whose sides are two circles, their
  Intersection the glass, focus = radius for n = 1.5; the picture is the candle scaled by
  m = f / (f - u) at the image distance, shown real past the lens or as a dashed ghost
  behind it, each gated by `0 * sqrt(...)`), Light It with the Sun (`BurningGlass.lgf`:
  15 sunbeams refracted exactly at both circular faces by Snell's law in vector form, each
  step a hidden `Number`; a 🔥 exists once enough beams pass the wick), Make a Rainbow
  (`Prism.lgf`: heavy flint indices per color, a beam traced to whichever side it meets
  first, reflected inside where it can't get out, out through the next side), Inside a
  Raindrop (`Raindrop.lgf`: the chords of a ray in a drop by their central angles, a
  show/hide box of many rays and one of the second bounce), Chase the Rainbow
  (`ChaseTheRainbow.lgf`: the view ahead in units of 10°, the bands Differences of circles
  around the antisolar point, clipped at the horizon by a hidden polygon; the bands are
  wider than a real rainbow's). Learned on the way: the expression language has no
  exponents, `1e-7` reads as 1·e - 7, so the tool writes small numbers in full and refuses
  an exponent; a guard must not compare a NaN (a comparison with NaN is NaN, and `if` of a
  NaN is NaN), so a square root is taken of `max(..., 0)` and the miss decided apart; a
  figure that is dragged by its body must have a free point among its roots (a lens whose
  middle was a point on a fixed axis could not be dragged at all: its roots were points by
  coordinates); and with a gradient paper a caption has no plate, so drawings with rays
  that run under the caption have a plain paper.
- **The dissections** (2026-10-26, two drawings at the end of the gallery, both written by
  `dotnet tools/dissections.cs -- <the Drawings folder>` and a `--rewrite`; `--probe
  <folder>` writes copies with the slider at 0, 0.5, 1, 1.5 and 2 for `--check`): The
  Haberdasher's Puzzle (`HaberdashersPuzzle.lgf`: Dudeney's four pieces of 1902 as
  polygons on `RotatedPoint`s, hinged at the midpoints D of AB and E of BC and at K on
  AC, each a half turn through the outside of the triangle so that none passes over
  another, all three from one slider through hidden `Number`s with formulas (a
  `RotatedPoint` takes a Number's value as degrees); the whole turns about E by 11° more
  so that the square comes out level - without that it ends tilted, as in every book) and
  Triangle to Square in Five Pieces (`TriangleToSquare.lgf`: the top folds down about D and
  E into the corners, a rectangle, then the P-slide: the piece above the slant slides
  along it, which opens the notch at the top right, and then the small triangle glides
  along the slant into it, listed last so that it passes over the base; the three moves
  take turns on one slider, since the small triangle would pass through the sliding piece
  if they went together). Both the tool checks: the pieces at the slider's end are in the
  square and reach its four corners. The sweep of the pieces is the `<Scene>` (the fit is
  taken at open time, so a piece swinging past the figure's box at rest would leave the
  screen), and the slider goes at the left, as low as the sweep leaves its track and
  caption clear at every step, else under the sweep. Dudeney's own five-piece answer of
  April 1902 is not described anywhere found (his column gave it; the four-piece one came
  from C. W. McElroy two weeks later), so the five-piece drawing is the natural one, not
  his.
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
