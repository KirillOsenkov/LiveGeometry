# The Geometry Snacks library

`snacks/` (2026-11-05; linked into `wwwroot/snacks/`): puzzles from *Geometry Snacks* (2018) and
*More Geometry Snacks* (2019) by Ed Southall and Vincent Pantaloni, published by Tarquin, redrawn
as live drawings at /snacks (the list) and /snacks/<slug> (the puzzle full screen, its question,
which book and number it is). Built like the Catriona Agg library (see "The Catriona Agg
library" in catriona-agg.md, whose pages these are copies of, cut down): `index.html`,
`puzzle.html`, `snacks.css`, `snacks.js`, `icons.svg`, the catalog `puzzles.json` and the
drawings in `drawings/`. web.config routes the addresses, the csproj links the folder and
`WriteGalleryIndex` writes a page per puzzle with its title, question and preview picture in
the head (`tools/previews.cs` renders the pictures for both libraries). The dev server is
`.claude/launch.json`'s "snacks-dev" (port 5012, over the same publish folder as "catriona-dev",
`C:\temp\LiveGeometry\aggsite`); `window.snacksPlayer` is the puzzle's player.
- **Terms with the authors** (Vincent Pantaloni said yes by email on 2026-11-06, Ed Southall is on
  board): about half the puzzles at most, no answers or solutions anywhere on the pages, both
  authors credited, links to the books (Amazon and Tarquin), and the prototype shown to them for
  review before it is published. Until then nothing links to /snacks from the start page.
- **Which puzzles**: only those that gain something from dragging, so a puzzle whose figure the
  statement fixes completely is left out. Doubtful ones are listed for the maintainer to decide.
  The aim is 30 to 50 percent of the puzzles, not more, so the books keep their point.
- **The look** is the books': the figures are a solid red (`#EF303F`, sampled from the pages)
  with black outlines of 3 px, on white; under the dark theme the red is a touch deeper
  (`#D9333F`) and the ink light. The second book draws in a lavender blue (`#B9B9FF`) instead;
  its drawings will say so. The pages take the Catriona Agg pages' layout with a red accent;
  no "Newest" card (the right half of the hero is left open), no search or arrangement controls,
  no date, no author links, no technique or solution links and no tags on a puzzle. The hero
  shows the two covers under the title (`book1.jpg`, `book2.jpg`, the publisher's pictures at
  600 px) with a Tarquin and an Amazon link under each, then the two paragraphs. The catalog's `books` carry the links, each puzzle its `book`, `number`, `chapter` (the book's
  running header), `title` (ours; the books number their puzzles without naming them),
  `statement` (the book's wording) and `drawing`.
- **The donut** of the books' cover turns slowly behind the hero at 30% opacity, centered on
  the hero's open second column, on a canvas as wide as the window so it isn't cut at the
  column's edge (`donut.js`, no dependencies: a jittered torus grid
  with a random diagonal per cell, flat shading from a fixed light, facets sorted far to near
  on a 2D canvas; the mouse drags it and a fling slows down, a finger scrolls the page instead;
  it always turns, reduced motion or not; drawn only while on screen, and drawn again inside
  the resize handler so a resize of the window shows no blank frame). Hidden below 980 px,
  where the hero is one column. Chosen on 2026-11-05 over a card in the right half; the lab
  with both placements and the options table is `C:\temp\snacks\donut` (outside the repo).
- **The drawings are ruler-and-compass constructions**, as every new drawing is (see "Ruler
  and compass from now on" in catriona-agg.md): a hand-written `.lgf` per puzzle, with a comment
  at the top saying how it is built, the `Handle` point style for what a reader drags, every
  helper hidden, the scene over the handle's whole range, no `PointByCoordinates` formulas. The
  first is `a-parallelogram-in-a-square.lgf` (book 1, #1): a square from two points by
  perpendiculars and circles, the handle on the top side, the opposite vertex where the line
  through the center meets the bottom side.
- **The 30th drawing**, `a-star-in-an-octagon` (book 1, #46), is the one fixed figure, chosen for
  its look: no handle, `live` false in the catalog so its tile carries no drag badge. The regular
  octagon is built about its center by reflections through it, a perpendicular and the lines
  through the midpoints of two chords; the star's points are the triangles cut off each side
  by the two diagonals that skip two vertices.
- **The 29 drawings of 2026-11-05** (16 from book 1, 13 from book 2) were made in parallel by
  agents from a brief (`C:\temp\Snacks\BRIEF.md`, outside the repo) and checked on the real
  puzzle page by a copy of the Catriona checker (`C:\temp\Snacks\snackcheck.cs`, built to
  `C:\temp\Snacks\bin\snackcheck\snackcheck.exe <slug> --out <dir>`: the page in both themes and
  on a phone, two frames with the handles moved, a sweep that reports every visible figure that
  stops existing). Two are rigid and only turn: `three-squares` (the tick marks and the 45°
  squares fix the ratio) and `two-concentric-pentagons` (parallel sides and concentricity put
  the small vertex at one point of the diagonal); `three-equal-segments-in-a-circle` changes
  its given angle with the handle; `a-path-across-a-cube` is the unfolded net. The maintainer
  decides whether those stay. What the agents learned, to keep: a base `FreePoint` must be
  `Locked="true"`, or the tiles' hover drift and the checker move it; an `IntersectionPoint`
  with a `Segment` exists only on the segment, so a point that may pass an end is cut on a
  `LineTwoPoints`; line-and-circle algorithm 1 is behind along the line's direction and 2
  ahead, circle-and-circle 1 is to the right of the first center seen from the second;
  an `AngleArc` in a filled `ShapeStyle` fills its sector (the books' marked angles) and draws
  a square at 90°; a `PointInRegion` is clamped inside a circle as well as a polygon; a label
  on a hidden point is hung on a shown point in a style that draws nothing (`Anchor`); the
  player fills a self-crossing polygon even-odd, so a pentagram is drawn as its ten-cornered
  outline. Book 2's shaded parts are the lavender `#B9B9FF` with blue `#0000FF` outlines, the
  rest black ink.
- **Reading the books**: the page images the drawings are made from are the maintainer's own
  (C:\temp\Snacks, outside the repo, from the Kindle editions), one PNG per page of the first book
  and per two-page spread of the second.
