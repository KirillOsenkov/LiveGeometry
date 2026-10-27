# The JavaScript player

`Main/Player/` (2026-10-07; the plan it follows is `docs/player-plan.md`): a drawing on any
web page, in plain JavaScript with no dependencies, a canvas with the Drag tool and nothing
else - no ribbon, grid, list or gallery, no undo, no selection.

- **A port of the library, file by file**: `src/` mirrors `LiveGeometry/` (`math.js`,
  `drawing.js`, `figures/points/midPoint.js`, `expressions/parser/scanner.js`,
  `serialization/drawingDeserializer.js`, `styles/styleManager.js`, `behaviors/dragger.js`),
  the C# names in camelCase, so that a reader of `MidPoint.cs` finds its twin and a change
  to one wants the same change in the other. **The rule: a new figure kind, a new function
  of the expression language, a new file attribute, a changed `Recalculate`, `HitTest`,
  `ReadXml` or sampler goes into the player in the same change** (a kind the player can't
  have is said so in its deserializer's error, "this version has no figure of the kind").
  What has no C# twin sits apart: `src/render/` (the canvas renderer, `TextMeasurer`) and
  `src/host/` (`Player`, `PlayerCanvas`, the `LiveGeometry` global). Left out on purpose:
  tools, undo, selection, snapping and joining, Convert verbs, layout changes of a Bezier
  path (anchors in or out, holes cut), transformations' images, Hyperlink, user tools.
- **JavaScript conventions**: an interface is a getter that answers true (`isPoint`,
  `isLine`, `isLinearFigure`, `isCircle`, `isLengthProvider`, `isAngleProvider`,
  `isNumber`, `isMovable`, `isMovableParts`, `isFigurePart`, `isBezierPathPiece`...), and
  a figure's own flag must not take such a name (an `isLine` field on the bisector broke
  every line: it is `wholeLine`). An expression reads a figure's fields as well as its
  getters (`ExpressionReflection.propertyNames`), so a field must not take the name of a
  C# property it doesn't hold (a Bezier path's typed tension was `tension`, and
  `p.Tension` read it instead of a tied slider's value: it is `typedTension`). An
  interface flag that is read must also be defined: `isBezierPathInterior` never was,
  and a path on its default style drew no fill. Classes that would shadow a browser
  global are renamed: `GeometryMath` (Math.cs), `NumberFigure` (Number.cs; the element is still `<Number>`),
  `SyntaxNode` (Node.cs). Every figure file ends with `FigureTypes.register("<element
  name>", Class)`. A figure has `render(renderer)` besides `updateVisual`; `resolvedStyle`
  is the style resolved for the theme on screen (what the Avalonia shape had applied);
  `stroke`/`fill` are what the renderer takes. Classic scripts, no modules, no bundler:
  `files.txt` is the load order (bases before derived classes), and a file referenced only
  at run time (`Dragger` from a figure's `hitTest`) may come later in it. The Write tool
  emits LF: new files are converted to CRLF afterwards, like the C#.
- **The bundle**: `Player.targets`, imported by the Browser and the Desktop projects,
  concatenates `files.txt` into `obj/player/<version>/player.js` before the build (an
  inline task, dotnet only; comment lines, block comments at the start of a line and the
  indentation are left out, 2026-10-27, which took the file from 244 KB brotli to 198: no
  minifier, the code reads as written, and the dev page loads the sources whole), wrapped
  in a function so that the host page sees one global,
  `LiveGeometry` (`play(element, options)`, `playAll`, `players`, and the classes for a
  page that builds on it). The Browser serves it as `/player/1/player.js`, every
  gallery drawing as `/gallery/<file>.lgf` and every Catriona Agg puzzle as
  `/CatrionaAgg/drawings/<slug>.lgf` (all `no-cache` and with
  `Access-Control-Allow-Origin: *` in `web.config`, the `shared-file` precondition, since a
  page on another site fetches the `.lgf` and the browser blocks that otherwise; a script
  tag needs no CORS. A new folder of drawings to embed goes into that precondition: the
  puzzles were left out until 2026-10-26, and an embed of one played nothing); the
  Desktop copies it to `player/player.js` beside the exe. **Versioning**: the folder is the
  player's major version. A snippet pasted into a post pins `/player/1/`; when the format
  breaks that player, `/player/2/` starts and the last `/player/1/player.js` is committed
  as a static file and served forever. Within a version the file revalidates, so fixes
  reach old embeds (since 2026-10-27 in the background, `stale-while-revalidate`: an embed
  plays the copy it has and gets the fix on its next load).
- **Embedding** (`PlayerEmbed.cs`; Export > "Copy embed code" and "Save as .html"): a
  script tag plus `<div class="livegeometry">` with the drawing's XML inside a `<script
  type="text/x-livegeometry">` (inline, works from a disk too; the XML can't contain
  `</script`, since it escapes `<`), or `data-src="<url of an .lgf>"` (the file's server
  must allow CORS - that goes for a page opened from `file://` too, whose origin is `null`:
  it plays the site's drawings but can't fetch a file beside itself on the disk). A page's
  link works as `data-src` where `Player.drawingAddress` knows the drawing behind it: a
  share link (`/s/<id>` is `/s/<id>.lgf`) and a Catriona Agg puzzle (`/CatrionaAgg/<slug>`
  is `/CatrionaAgg/drawings/<slug>.lgf`). Options: `data-theme`
  (light, dark, auto), `data-font`, `data-wheel="zoom"` (else the wheel zooms only with
  Ctrl, so an embed doesn't take a blog's scrolling), `data-fit="content"`. "Save as .html"
  writes the player into the page. Fonts: the host page's, and the browser's own emoji
  font (none shipped; a character the system lacks is a box). The page that explains all
  of this to a reader is `/embed` (see "Deployment and caching"): a change to the options
  or the snippet wants the same change there.
- **Behavior that differs from the app, on purpose**: no selection, so a Bezier path's
  handles show from a press on an anchor until the next press elsewhere
  (`BezierPath.showHandlesWhileDragging` from `Dragger.mouseDown`), and the Tab choice of
  a handle under its anchor doesn't exist. A drawing with a caption (Title and Description
  labels: the gallery's, an export of one) is laid out for the element as the app lays it
  out for its window, by a port of the app's `GalleryDrawing.Fit` (`src/host/galleryDrawing.js`,
  the one twin of a file outside the library), now and on every zoom to fit; an export
  carries the pin and offsets it had on screen, which a small element would clip. Any
  other drawing gets the file's viewport (or its scene) fitted into the element.
- **Working on it**: `dotnet run tools/serve.cs -- Main 5005` (`.claude/launch.json` has it
  as "player-dev") serves the repo's `Main` folder, and
  `http://localhost:5005/Player/dev/index.html#<Drawing>` loads `files.txt` one script at a
  time (`?v=` cache-busted; to reload after an edit change the query, `?r=2`, a changed hash
  alone reloads nothing) with a list of the gallery drawings, Fit, dark and wheel switches,
  and a status line with the figure count, the load time and the load errors.
  `Main/Player/dev/kinds.lgf` has the kinds no gallery drawing uses (a regular hexagon of
  points by coordinates, a polyline, lines and a circle by equation, a line at an angle, segment marks, a
  name label, a sector, an ellipse arc and segment, a horizontal angle). To check every
  drawing at once, in the page's console: for each file of the listing, `new
  Drawing(player.canvas).addFromXml(text)` and look at `loadErrors` and the visible figures
  that don't exist (Fireworks, Pump Up the Balloon and Poke the Blob have some by design).
  For pictures use headless Edge (`webauto start <url> 1000 700`, `nav`, `shot`, `drag`,
  `eval`): the app's browser pane can be too small to show a drawing, and `--check`'s PNGs
  of the desktop are the reference to put beside them. Headless Edge has a pixel ratio of
  1; `Player.pixelRatio` says what the canvas is scaled by. **A resize draws at once**
  (2026-10-25): setting a canvas's size clears it, and the ResizeObserver runs just before
  the frame is painted, so `Player.resize` (and a page's player that overrides it) calls
  `render()`, not `invalidate()` - a repaint on the next frame left every step of a window
  drag showing a blank canvas, a flicker. `Main/Player/dev/bundle.html`
  plays the bundle the Desktop build made, as a page on another site would: an embed by
  URL and an inline one (dark, wheel zooming); `hosted.html?host=http://localhost:5006`
  takes the player and a drawing from another origin (a publish served by `tools/serve.cs`
  on that port, or `https://livegeometry.com` after a deploy: the cross-site fetch for real).
- **The check**: `dotnet run tools/playerparity.cs` (a project reference to the desktop
  head, like `regression.cs`): every figure kind `DrawingDeserializer.FigureTypes` reads is
  registered in the player or in `Main/Player/excluded.txt` with a reason, every function
  of `Functions` has a twin in `functions.js`; then every gallery drawing is loaded by both
  (the library in-process, the player in headless Edge on the dev page, started if they
  aren't running), dumped the same way (`Drawing.dump`: the figures in the list's order
  with the numbers that define them, points, lines, ellipses, polygon vertices, label
  texts, numbers, rounded to 9 decimals) and compared at 1e-6, once as loaded and once
  after every free point is moved by (0.3, 0.2). Left out of the numbers: what is sized in
  pixels (an angle's mark) and what is sampled by the window (a locus, a graph; the points
  on them are compared). Run it after a change to a figure's numbers on either side.
