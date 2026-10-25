# The Catriona Agg library

`CatrionaAgg/` (2026-10-25; linked into `wwwroot/CatrionaAgg/`): Catriona Agg's geometry puzzles
redrawn as live drawings, at /CatrionaAgg (the gallery: search, arrange by date, year, topic,
shape or question, the Invariance filter) and /CatrionaAgg/<slug> (the puzzle full screen, her
question, links to Andrew Stacey's puzzle, techniques and solution pages on Mathforge Notes).
One page, `index.html`, played by the player (`/player/1/player.js`), no .NET: it reads the
address like /web. The puzzles are hers and the catalogue his: both are credited on every view,
the solution is never on the page, only linked, and her original drawing is linked (Mathforge's
page shows it), never copied.
- **The catalog** is `puzzles.json`, written by `dotnet run tools/catriona.cs -- [--recent N]
  [--images <folder>]` from Mathforge's wiki sources (`/notes/source/<page>`): every puzzle of
  the year pages, the topics from "Puzzles by Topic" (or, for one newer than that page, its own
  techniques page), and for the drawn and the --recent ones the title, the question and the
  BlueSky thread (Andrew Stacey links no tweets; neither do we). Pages are cached in the temp
  folder, so a rerun asks for nothing it has. Fields it doesn't write (`play`, the line under
  the question saying what to drag, by hand) are kept. Rerun it after adding drawings: it
  marks the drawn ones (`drawing`), and the gallery shows only those.
- **The drawings** are `drawings/<slug>.lgf`, written by `dotnet run
  tools/catrionadrawings.cs -- [slug ...]`, one function per puzzle: her diagram as a
  construction over one or two handles (points on hidden segments or arcs that keep them in
  range) so that what the puzzle says holds while it is dragged - for an invariance puzzle the
  answer stays put. Hidden points by coordinates over the handles carry the geometry; a puzzle
  the conditions fix has no handle. Her colors (the averages of the originals' areas) in
  pencil or felt tip as she used them (`Lgf.Texture` per drawing, see "Textured fills"; null
  for a gradient),
  outlines in ink with a light Dark override, text on a filled shape the same in both themes.
  Traps met: a point on an arc keeps its angle from the center as its parameter (absolute, not
  from the arc's start); a text attached to a point hides with the point, so a text hangs on a
  shown point in a style that draws nothing (`Anchor`); and nothing built on a handle may be
  `Locked` - a point counts as locked when anything built on it is, and the handle froze. An
  angle mark (`AngleArc`) draws a square at 90°, which tells the answer when it is 90°: such a
  mark is a plain arc (`ArcMark`). Check them all at once in the page's console (load, then
  sweep every handle and look for visible figures that don't exist).
- **Tiles** are pictures: each drawing is played once off screen without its handles and kept
  as a PNG (an object URL) that every tile of it shows, only near the view, the nearest first;
  a theme change draws them again. Hovering a tile plays it and drifts its handles. A
  rearrangement glides the tiles to their new places (FLIP, tiles reused per puzzle).
- **Link previews**: the `WriteGalleryIndex` task writes `CatrionaAgg/<slug>.html` (the page
  with the puzzle's title, question, canonical address and `/CatrionaAgg/previews/<slug>.png`
  in the head, the question in the text a search engine reads) and the sitemap entries;
  `tools/previews.cs` renders those pictures and `previews/index.png`, a mosaic, for the
  gallery. web.config routes /CatrionaAgg and /CatrionaAgg/<slug>.
- **Working on it**: `.claude/launch.json` has "catriona-dev": `tools/serve.cs` over a folder
  holding `player/1/player.js` (copied from a build's `obj/player/1/`) with
  `--overlay CatrionaAgg=CatrionaAgg --reload`; `window.catrionaPlayer` is the puzzle's player.
