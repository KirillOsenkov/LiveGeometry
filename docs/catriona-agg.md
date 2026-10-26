# The Catriona Agg library

`CatrionaAgg/` (2026-10-25; linked into `wwwroot/CatrionaAgg/`): Catriona Agg's geometry puzzles
redrawn as live drawings, at /CatrionaAgg (the gallery: search, arrange by date, year, topic,
shape or question) and /CatrionaAgg/<slug> (the puzzle full screen, her
question, links to Andrew Stacey's puzzle, techniques and solution pages on Mathforge Notes).
Two pages, played by the player (`/player/1/player.js`), no .NET: the gallery `index.html` and
`puzzle.html`, which reads the puzzle from its address (and says when it has no drawing of it).
They share `catriona.css`, `catriona.js` (each page runs its part) and `icons.svg` (the
icons, each drawn by `<use href="/CatrionaAgg/icons.svg#icon-back">`; the mark is the site's
`/logo.svg`, whose gradients an icon referenced from another file may not find), linked by name
(revalidated, not cached for good: web.config's mutable set), so a page carries its markup
alone. Until 2026-10-25 both were one page with two views, the gallery kept behind the puzzle;
now a step from one to the other is a page load, and what carried over is kept for the tab
(`sessionStorage`): the gallery's arrangement (grouping, search, order), which the puzzle's
arrows step through (date order when there is none, or it leaves the puzzle out), and the
gallery's scroll, put back when "All puzzles" or Escape (a flag) or Back brings the reader to
it. The address of the gallery says the arrangement too, and wins when it says one: a puzzle's
tag is `/CatrionaAgg?by=topic#group-circle-theorems`. Browser Back usually restores the
gallery from memory, tiles and all; otherwise the tiles draw again. The puzzles are hers and the catalogue his: both are credited on every view,
the solution is never on the page, only linked, and her original drawing is linked (Mathforge's
page shows it), never copied.
- **The catalog** is `puzzles.json`, written by `dotnet run tools/catriona.cs -- [--recent N]
  [--images <folder>]` from Mathforge's wiki sources (`/notes/source/<page>`): every puzzle of
  the year pages, the topics from "Puzzles by Topic" (or, for one newer than that page, its own
  techniques page), and for the drawn and the --recent ones the title, the question and the
  BlueSky thread (Andrew Stacey links no tweets; neither do we). Pages are cached in the temp
  folder, so a rerun asks for nothing it has. Fields it doesn't write are kept. Rerun it after
  adding drawings: it marks the drawn ones (`drawing`, and `live` when the drawing has
  handles: the drag badge on the tile), and the gallery shows only those. (A "play" line per
  puzzle saying what to drag, under the question, and a "Drag the white points" pop-up over
  the drawing were tried and dropped, 2026-10-25: they added nothing.)
- **The drawings** are `drawings/<slug>.lgf`, written by `dotnet run
  tools/catrionadrawings.cs -- [slug ...]`, one function per puzzle: her diagram as a
  construction over one or two handles (points on hidden segments or arcs that keep them in
  range) so that what the puzzle says holds while it is dragged - for an invariance puzzle the
  answer stays put. Hidden points by coordinates over the handles carry the geometry; a puzzle
  the conditions fix has no handle. Her colors (the averages of the originals' areas) in
  pencil or felt tip as she used them (`Lgf.Texture` per drawing, see "Textured fills" in
  styles.md; null for a gradient), outlines in ink with a light Dark override, text on a
  filled shape the same in both themes. Under the dark theme the fills are not hers: a scan
  of ink on white can't be darkened faithfully (its pale spots would be the dark paper), so
  each fill's `<Dark>` says `Texture="None"` and a gentle diagonal gradient in a tone of the
  same hue (`DarkTone`: lightness pulled toward the middle, chroma - not HSL saturation, which
  a pastel has as much of as a vivid color - raised for pastels and capped), an artistic
  liberty decided 2026-10-25. Colored lines, marks and labels are a light tint of their hue
  there (`DarkLine`), but those that mostly cross fills the dark theme leaves bright (yellow,
  orange, green) are given a deep color (`dark:` on `Line`/`Text`), since a light line
  vanishes on those; a colored text on a fill is deep in the same hue (`DarkOnShape`). Look
  at every drawing under the dark theme after changing one.
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
- **Link previews**: the `WriteGalleryIndex` task writes `CatrionaAgg/<slug>.html` (`puzzle.html`
  with the puzzle's title, question, canonical address and `/CatrionaAgg/previews/<slug>.png`
  in the head, the question in the text a search engine reads) and the sitemap entries;
  `tools/previews.cs` renders those pictures and `previews/index.png`, a mosaic, for the
  gallery. web.config routes /CatrionaAgg to the gallery and /CatrionaAgg/<slug> to the
  puzzle's page, else to `puzzle.html`.
- **Working on it**: `.claude/launch.json` has "catriona-dev": `tools/serve.cs` over a folder
  holding `player/1/player.js` (copied from a build's `obj/player/1/`) with
  `--overlay CatrionaAgg=CatrionaAgg --route CatrionaAgg=CatrionaAgg/puzzle.html --reload`
  (every puzzle's address gets `puzzle.html` itself, the build's copies not being there);
  `window.catrionaPlayer` is the puzzle's player.
