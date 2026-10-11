# Gallery ideas

A brainstorm from 2026-10-10 about what the gallery lacks and what to add. Kept here so
that the ideas aren't lost; strike an item when it is built or rejected.

## What the gallery is for

The gallery whets the appetite. A drawing goes in when a kid would say "wait, what?"
before knowing the word for what they see: motion they cause, a curve they didn't expect,
a pattern that fills itself, a puzzle with a trick ending, a measurement that says the eye
is lying, 3D out of a flat page.

The curriculum is not the gallery's job. Angle sum, parallel lines and a transversal,
similar triangles, area formulas, vectors, the unit circle, Euclid's constructions, the
Euler line: those belong in a separate structured page (`/library` or `/lessons`, below),
in order, chapter by chapter. Don't push them into the gallery.

## How the gallery looks today (64 drawings)

Two strong clusters: toys for the eye (slime, jelly, balloon, fireworks, castle, the
emoji drawings) and named theorems from olympiad folklore (Morley, Simson, Ceva,
Desargues, Pappus, Van Aubel, Napoleon, Steiner). Observations:

- **Motion and mechanisms are under-used.** Loci are the app's best trick and only
  Spiral, Continuous Deformations and the ellipse's evolute use them. Nothing rolls,
  swings, meshes or links.
- **Only one transformation shows.** Two Reflections and the inversion. Rotation,
  translation, dilation and the tessellations they make are absent, though every tool
  for them exists.
- **The thumbnails blur together.** About fifteen tiles are a pale blue triangle with
  thin gray lines. Varying the hue per drawing on purpose (the palette has eight) would
  make the grid scannable.
- **Named theorems mean nothing to a 13-year-old** as titles (Van Aubel, Ceva, Pappus).
  A title that says what you see might serve better; the name can live in the caption.
- **At 80 drawings the flat grid will want sections or tags** (Play, Curves, Theorems,
  Patterns, 3D, Puzzles). The tour order can keep interleaving; the page could group.
- **No drawing uses** the Vector tool, the Perimeter measurement (2026-10-09), Line at
  Angle, Rotate or Dilate.

## Gallery candidates

**Built on 2026-10-10** (19 drawings, in the gallery now): cycloid, spirograph, two pins
and a string, Peaucellier linkage, gears, clock hands, billiards, tessellation, snowflake,
times tables on a circle, flower of life, tangram, optical illusions, spiral of Theodorus,
pizza slices, spin the die, fold the cube, two-point perspective, what is π. Still open
below: pendulum, rotation and dilation playgrounds, Islamic star, Dudeney's dissection,
proofs without words, the missing square, cross sections, the shadow of a stick.

Roughly by spark per hour of work. "Generated" means a `tools/<name>.cs` writes the
`.lgf`, as the slime and the Platonic solids are made.

### Motion, the loci the app is best at

- **Cycloid.** A wheel on a road, a point on the rim, the locus as the wheel rolls. Drag
  the wheel along the road. The center is a point on a line; the rim point is rotated
  about it by an angle tied to a label `[distance / radius]`. Second scene or a slider:
  the point inside or outside the rim (curtate and prolate cycloids).
- **Spirograph.** Epicycloids and hypocycloids: a small circle rolling inside or outside
  a big one, a slider for the ratio of the radii and one for how far the pen sits from
  the small circle's center. A kid magnet. Generated, or hand-built with expressions.
- **Ellipse with two pins and a string.** A point held by a fixed sum of distances to
  two pins; drag the pins, the string length by a slider. Fix length makes it honest.
- **Peaucellier linkage.** Seven bars, a crank, and the free end draws a straight line.
  It is inversion in action, so it links to the inversion drawing.
- **Gears.** Two circles with teeth, a slider turns one, the other follows by the
  ratio. A third gear reverses the direction again.
- **Clock hands.** A slider for the time, the angle between the hands measured. The
  classic puzzle ("when are the hands at right angles?") in the caption.
- **Billiards.** A ball in a rectangle bouncing by reflections; aim at a pocket by
  dragging the cue direction. The trick: unfold the table by reflecting it.
- **Pendulum / swinging.** A point on an arc with a slider for the angle, the bob's
  height measured: the potential-energy picture without the words.

### Transformations and patterns

- **Escher tile.** A Bezier path whose edges you drag, with translated copies tiling the
  plane; the opposite edges are images of each other (the path images already follow
  their source). A square or hexagonal grid of copies.
- **Design a snowflake.** Kaleidoscope's idea with a Bezier path instead of emoji: draw
  one twelfth, six-fold rotation and a reflection fill the rest.
- **Times tables on a circle.** n points on a circle, a segment from point i to point
  k·i mod n, a slider for k. A cardioid appears at 2, a nephroid at 3. Generated.
- **Rotation and dilation playgrounds.** A figure, a center and an angle or a factor you
  drag; the image follows. Also a glide reflection.
- **Islamic star pattern.** A rosette from one regular polygon and its rotated copies.
- **Flower of life.** Seven circles, each through the centers of its neighbors; the
  Define figure tool could make the step a tool.

### Puzzles and illusions

- **Tangram.** Seven pieces, each polygon on a pivot and a direction point, so a drag on
  the piece moves it and a drag on the second point rotates it. A shape to match shown
  faintly behind; a show/hide box for the solution.
- **Dudeney's hinged dissection.** A triangle folds into a square with one slider for the
  hinge angle.
- **Optical illusions.** Müller-Lyer, Ebbinghaus, Ponzo, the café wall, with a Distance
  measurement to prove the eye wrong. Cheap to make and perfect for the audience. One
  drawing with several, or one each.
- **Spiral of Theodorus.** Sixteen right triangles making every square root, each on
  the previous hypotenuse.
- **Proofs without words.** The sum of the first n odd numbers is a square; 1/2 + 1/4 +
  1/8 ... fills the square; (a + b)² as four rectangles with sliders for a and b.
- **The missing square.** The 64 = 65 chessboard paradox, with an Area measurement.
- **Pizza theorem** and the area of a circle as sectors laid head to tail into a
  parallelogram (sectors exist since 2026-10-09).

### 3D, where only three drawings live

- **Spin the cube.** A slider or two for the rotation, done with expressions over hidden
  variable points as the Platonic solids are. Generated.
- **Net of a cube folding up** with a slider for the fold angle. Generated.
- **Two-point perspective.** A box whose vanishing points you drag. Gives Desargues a
  reason to exist nearby.
- **Cross sections of a cube.** A plane by a slider, the section polygon changes from
  triangle to hexagon.
- **Shadow of a stick** / sundial: a sun you drag, the shadow's length.

### Borderline

- **What is π.** Drag any circle, the Perimeter over the diameter never moves. As
  "the number never changes" it is a spark; as a formula it is a lesson. Also the
  first use of the Perimeter measurement in the gallery.
- **Euler line, nine-point circle, centroid 2:1.** The same kind as Morley and Simson,
  so they fit the theorem cluster, but they are lessons first.

### The first batch

times tables on a circle, cycloid, spirograph, optical illusions, tangram, Escher tile.
The first three generated, the last three by hand.

## The lessons page (`/library` or `/lessons`)

Not built. The decision to make early: what a lesson *is* on this site, since the player
exists. The cheapest shape is a page per topic with prose and several embedded players,
each a small drawing with one thing to drag, in a clear order within a chapter, as
`/embed` already demonstrates. The gallery drawings become the "see it in the wild"
link at the end of a lesson. The alternative, `.lgf` files with captions and a second
catalog, is less flexible.

Content that belongs there, in curriculum order:

- Euclid's first construction (the splash animates it and nothing on the site shows
  it); bisecting an angle, a perpendicular bisector, copying an angle, a perpendicular
  through a point.
- Angle sum of a triangle: three measurements and a live total that stays 180; the
  corners rotated to sit on a line. Exterior angle. Angle sum of a polygon.
- Parallel lines and a transversal: eight angles, which pairs stay equal.
- Triangle congruence and the SSA ambiguity (two triangles from the same data).
- Similar triangles; how tall is the tree (a stick, its shadow, the tree's shadow);
  Eratosthenes and the size of the Earth; how far is the horizon.
- Area: a triangle is half a parallelogram, a trapezoid by the same shear, Heron.
  Cavalieri next to these.
- Circles: π by perimeter over diameter; Thales; the inscribed angle (in the gallery);
  cyclic quadrilaterals; intersecting chords and the power of a point; tangents (in the
  gallery).
- Triangle centers: centroid (medians, 2:1, six equal areas), orthocenter,
  circumcenter and incenter (in the gallery), the Euler line, the nine-point circle.
- The unit circle: a point on a circle, cos and sin as its coordinates, a locus tracing
  the sine wave beside it; tangent as a length on the tangent line.
- Vectors: head to tail, the parallelogram, a boat crossing a river, forces on a slope.
- Transformations one at a time: reflection, rotation, translation, dilation; which
  compositions give which; symmetry groups of a square.
- The geometric mean in a semicircle, the golden ratio, the golden rectangle, square
  roots by the spiral of Theodorus.
- Quadrilaterals: the family tree (the app classifies them live), rhombus diagonals,
  kite, parallelogram properties.
- Coordinates: slope, distance, midpoint, the equation of a line and of a circle.
