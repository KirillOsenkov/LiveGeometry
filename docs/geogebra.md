# GeoGebra files (.ggb)

`Serialization/GeoGebraReader.cs` opens GeoGebra worksheets (Open dialog, command line,
`--check`): a `.ggb` is a zip with `geogebra.xml` inside. The construction is free elements
(`<element type="point">` with `<coords x y z>`, homogeneous: divide by z) and
`<command name="Segment"><input a0="A" a1="B"/><output a0="f"/>` with the outputs' elements
(style, show, coords) following the command. Each known command becomes the figure that
stands for it here, with hidden helpers where the shapes differ (a circle through three
points is the circle around the crossing of two bisectors, an ellipse by foci is center and
axis ends as points by coordinates, a regular polygon is a plain polygon whose vertices are
points by coordinates over its side as the Polygon on a side tool makes them, keeping the
file's names, tangents from a point go through the Thales circle);
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
numeric inputs to Rotate are radians, while imported angle Numbers already hold degrees, and
so do the hidden labels made for angles (`[deg(...)]`, read in expressions as `rad(a.Value)`).
GeoGebra's `log` is e's and ours 10's: `log`, `lg`, `ld`, `sgn`, `nroot` are renamed
(`ln`, `log`, `log_2`, `sign`, `nthroot`); its `If` and comparisons are ours as they are.
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
