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
  no date, no author links, no technique or solution links and no tags on a puzzle. The
  catalog's `books` carry the links, each puzzle its `book`, `number`, `chapter` (the book's
  running header), `title` (ours; the books number their puzzles without naming them),
  `statement` (the book's wording) and `drawing`.
- **The drawings are ruler-and-compass constructions**, as every new drawing is (see "Ruler
  and compass from now on" in catriona-agg.md): a hand-written `.lgf` per puzzle, with a comment
  at the top saying how it is built, the `Handle` point style for what a reader drags, every
  helper hidden, the scene over the handle's whole range, no `PointByCoordinates` formulas. The
  first is `a-parallelogram-in-a-square.lgf` (book 1, #1): a square from two points by
  perpendiculars and circles, the handle on the top side, the opposite vertex where the line
  through the center meets the bottom side.
- **Reading the books**: the page images the drawings are made from are the maintainer's own
  (C:\temp\Snacks, outside the repo, from the Kindle editions), one PNG per page of the first book
  and per two-page spread of the second.
