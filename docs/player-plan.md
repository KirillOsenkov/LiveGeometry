# JavaScript player: plan

Status: 2026-10-07, phases 1 to 6 built (the port of every kind the library reads but the
two excluded ones, the bundle, the export items, the hosting rules, `tools/playerparity.cs`,
the AGENTS.md section); phase 7 and the cleanup of the Silverlight-era defines remain.
AGENTS.md ("The JavaScript player") is the current description; this is the plan it came from.

## Goal

A drawing exported from the app plays on any web page that can hold HTML: an HTML canvas
that shows the drawing and lets the reader drag points and figures (and tick show/hide
boxes, pan and zoom). No ribbon, no property grid, no Figure List, no gallery. The page may
fetch files from https://livegeometry.com.

Explicit non-goals for now: tools, editing, undo, selection, a toolbar, a gallery, GeoGebra
and DG files. But nothing may be built in a way that keeps those from being added later,
up to a full editor in JavaScript one day.

The Silverlight generation had exactly this (`Main/SilverlightPlayer`: a canvas with the
Drag tool, the drawing given by a `LoadFile` URL, built from the same library with a
`PLAYER` define). It is obsolete, as is everything Silverlight: the `#if !PLAYER`,
`TABULA` and `TABULAPLAYER` blocks still in `DynamicGeometry` (about 40 of them) are
leftovers to delete, not to build on; the new player is JavaScript and needs no define.
Their removal is a cleanup item of phase 6.

## The decision: a hand-written JavaScript port, mirroring the C# file by file

Decided 2026-10-07: pure JavaScript, no .NET, no Avalonia, no dependency on anything we
don't build ourselves and host on livegeometry.com. (A trimmed .NET WebAssembly build, the
Silverlight way, was considered and dropped: the library is grown into Avalonia, so a
headless core would be a refactoring of the whole library, and the result would still be
a multi-megabyte download per page where an embed has to be as cheap as an image.)

The port covers the part of the library a player needs, written as parallel code: the
same folders, one JS file per C# file, the same class and method names, so that a reader
of `Figures/Points/MidPoint.cs` finds `figures/points/midPoint.js` beside it and a new
figure kind in C# visibly lacks its JS twin (a check enforces that, see "Keeping it in
sync").

The part to port is about 12-15k of the library's 68k lines: the geometry (`Math.cs`),
the figure classes' `Recalculate`/`HitTest`/`ReadXml`, the dependency engine, the
expression language, the deserializer, the default styles, the coordinate system and grid,
labels, and the plain path of the Drag tool. Roughly 8-10k lines of JavaScript, plus a
canvas renderer that has no C# counterpart (Avalonia draws the shapes there).

## What the player plays

Everything an `.lgf` can hold, in tiers. Tier 1 is what the gallery's 62 drawings use,
by frequency (counted over the folder): PointByCoordinates (1786 of them: the expression
language is tier 1), Segment, FreePoint, Polygon, MidPoint, IntersectionPoint (its four
algorithm pairs), Label (pinned captions with `[expressions]`, wrap width, backdrop),
PointOnFigure (and so the parameter on every figure kind), PointLabel, RotatedPoint,
PerpendicularLine, Circle, ReflectedPoint (point, line and circle mirrors), LineTwoPoints,
BezierPath with the smoother and holes, Locus (the adaptive sampler and the recomputed
chain), CircleByRadius, AngleBisector, Bezier, Ray, DilatedPoint, ShowHideControl,
CircleArc, DistanceMeasurement, Number, AngleMeasurement and AngleArc, Vector, Slider,
ParallelLine, AreaMeasurement, FunctionGraph, Ellipse; gradient fills and papers, Dark
overrides, Scenes, the grid and axes.

Tier 2, nothing in the gallery but the tools make them: SegmentBisector, LineAtAngle,
LineByEquation, CircleByEquation, TranslatedPoint in full (the slider's knob is one, so its
core is tier 1), RegularPolygon with its parts, Polyline, the ellipse arcs, sectors and
segments, FigureLabel, segment decorations, right angle marks, AxisLine, points on the
axes, HorizontalAngleMeasurement.

Left out, by design: Hyperlink (fetches drawings; no tool makes one), PolygonIntersection
(a stub, `[Ignore]`d), undo, selection halos, snapping and joining, the context menu, every
tool but dragging, user-defined tools.

## Architecture

- **Folder**: `Main/Player/`, next to `Main/Avalonia/`. Inside, `src/` mirrors
  `DynamicGeometry/`: `math.js`, `drawing.js`, `figures/figureBase.js`,
  `figures/points/freePoint.js`, `figures/lines/segment.js`, `expressions/parser/scanner.js`,
  `serialization/drawingDeserializer.js`, `styles/styleManager.js`, `behaviors/dragger.js`...
  Names stay the C# names in camelCase: `recalculate`, `updateExistence`, `hitTest`,
  `readXml`, `getPointFromParameter`, `getNearestParameterFromPoint`, `moveTo`, `allowMove`,
  `Math.getIntersectionOfCircleAndLine`. New files with no C# twin (the renderer, the host
  element) sit apart in `src/render/` and `src/host/`.
- **Language**: plain JavaScript (ES2022 classes), `// @ts-check` with JSDoc types so the
  editor checks it without a compiler. Not TypeScript for now: it needs a toolchain, and
  the one-file-per-class layout survives a move to it later if wanted.
- **Build**: no bundler. The files are classic scripts on one namespace, listed in load
  order (bases before derived classes, as the C# inheritance goes) in `Main/Player/files.txt`;
  a dev page (`Main/Player/dev/index.html`, served by `tools/serve.cs`) loads them one by
  one, and an MSBuild target of the Browser project concatenates them into
  `wwwroot/player/<version>/player.js` at build and publish. Dotnet only, like the rest of
  the repo. Minification with esbuild (`npx`, node is installed here and on the CI runners)
  can be added later if the size matters; unminified the player should be 200-300 KB, 50-70
  KB brotli.
- **The engine stays the engine.** `Drawing`, the flat figure list in dependency order,
  `Dependencies`/`Dependents` registration, `recalculate` down the list,
  `recalculateAllDependents` through the topological sort, `exists` propagation,
  composites with parts, `Figures[name]` looking inside composites: all ported as they
  are, not flattened into an evaluator for a fixed drawing. That is what lets tools add
  figures live later. Every change of a figure's place goes through an `actions.js`
  (`move(movables, offset, toRecalculate)`), which does the move directly now and is where
  undo attaches later.
- **Rendering**: a `Renderer` interface (`drawLine`, `drawEllipse`, `drawArc`, `drawPolygon`,
  `drawPath`, `drawText`, `measureText`, `beginFrame`...) with one implementation, Canvas 2D,
  redrawing the whole scene per animation frame (figures sorted by `zIndex`, the grid
  first, pinned labels last), devicePixelRatio aware, sizes in CSS pixels as the app's are
  in DIPs. Each figure keeps its `updateVisual` (it computes on-screen geometry: clipped
  lines, arc parameters, label places) and gets a `render(renderer)` the C# doesn't have.
  An SVG-DOM renderer could replace it one day without touching the figures.
- **Hit testing** is already geometry in C# (`hitTest(point)` in logical units, cursor
  tolerance plus half the stroke), so it ports as is. The two Avalonia pixel hit tests (a
  filled arc's inside, a Bézier's inside) become point-in-polygon on the sampled outline;
  labels use the measured text box.
- **Text**: `fillText` with word wrapping by `measureText` at `WrapWidth`; `NameDisplay`
  already turns `A_1` into Unicode subscripts, so names need nothing special; a backdrop is
  a filled rectangle of the paper. Font: the host page's `font-family` by default, or the
  one the style names if the page has it, so the embed reads like the post around it;
  `data-font` to force one (Inter from livegeometry.com, if we serve it with CORS).
- **Emoji points**: the browser's own color emoji font (`fillText` draws it on every
  modern platform), so the player brings no font. Decided: no Twemoji in the player.
- **Interaction**, the Drag tool's plain path only: press, hit test, then one movable (a
  point jumps under the cursor, a label keeps its grab offset, a slider gives its knob,
  anchor or whole), or the free roots of a dependent figure (`findRoots` minus Numbers and
  points by coordinates; a locked figure anywhere downstream locks the drag), else the view
  pans; `DragThreshold` 3 px before a press is a drag; wheel zooms 1.2 around the cursor;
  one finger drags, two fingers pan and zoom (`CoordinateSystem.panAndZoom`); a click on a
  show/hide box ticks it; a double click zooms to fit; Shift snaps to the grid step; the
  hand cursor over what a press would move. `Behavior` with its pointer adapter is ported
  as the base, `Dragger` derives from it, so tools derive from the same base later.
- **Fit**: the file's Viewport fitted into the element, as the app opens a file
  (`SetViewport` = `Fit(rect)`), or a scene when the file has scenes. `data-fit="content"`
  for `ZoomExtend`. The gallery's caption layout (`GalleryDrawing.Fit`) is UI of the app,
  not of the library, and is not ported: the exported file carries its pin and offsets.
- **Styles**: the default styles are ported from `StyleManager.AddDefaultStyles` (a file
  names them without carrying them) with their Light values and Dark overrides, so the
  player can play any `.lgf`, the gallery's included. `data-theme="light|dark|auto"`: the
  data is in the files already, so dark is cheap.
- **Host API** (`src/host/`): `LiveGeometry.play(element, options)` returns a player with
  `drawing`, `load(lgfText)`, `fit()`, `resize()` (a ResizeObserver follows the element),
  `dump()` (for the parity check), and a `dispose()`. An element with class `livegeometry`
  is played on page load from an inner `<script type="text/x-livegeometry">` holding the
  XML, or from `data-src` (a URL; the response must allow CORS).

## Embedding

Decided: no iframe. Two forms, both one script tag plus one element.

- **Inline** (Export > "Copy embed code"):
  ```html
  <script src="https://livegeometry.com/player/1/player.js" defer></script>
  <div class="livegeometry" style="width: 100%; height: 480px;">
  <script type="text/x-livegeometry"><Drawing Version="1" ...>...</Drawing></script>
  </div>
  ```
  The drawing goes in as is: the `.lgf` XML is the one format, written by the serializer
  there is (XML escapes `<` in attributes, so `</script>` can't occur in it). Pasteable
  into any HTML page, GitHub Pages, a static blog, a wiki that allows scripts; works in a
  local `.html` file opened from disk, since nothing is fetched but the player.
- **By URL**:
  ```html
  <div class="livegeometry" data-src="https://livegeometry.com/gallery/Pythagoras.lgf"></div>
  ```
  The `.lgf` can live anywhere: next to the `.html` on the author's own site, or on
  livegeometry.com, which serves every gallery drawing as `/gallery/<file>.lgf` (the
  Drawings folder linked into `wwwroot/gallery/`, beside the `/gallery/<slug>` routes; a
  path with an extension is a file and never falls through to the app).
- **"Save as .html"**: a page with the inline snippet; and a self-contained variant with
  `player.js` inlined (the desktop reads it from beside the exe as it reads the emoji font,
  the browser fetches `/player/<version>/player.js`). With the browser's emoji font there
  is nothing else to inline: one file that works from a USB stick.
- **GitHub README, issues and comments** strip scripts: the only thing that works there
  is a picture (Save as .png/.svg) linking to a page that holds the embed. Said as much
  in the export menu's tooltip, not solved.

### How cross-site loading works, for the record

- A `<script src="https://livegeometry.com/...">` on any site is allowed by every
  browser without conditions: scripts are exempt from the same-origin rule. So the
  player itself needs nothing from the server but to be served.
- The player then runs *as part of the host page*, with no access to livegeometry.com's
  storage or cookies and no need for it. Its only contact with our server is downloading
  files. It must not assume anything about the host page: it draws only inside its
  element, touches no global style, and prefixes the one global it defines
  (`LiveGeometry`).
- `fetch` of a file from another origin (the `.lgf` by URL) *is* subject to the
  same-origin rule: the browser blocks the response unless the server that holds the file
  answers with `Access-Control-Allow-Origin: *` (CORS). livegeometry.com will send that
  header for `/gallery/*.lgf` and `/player/*` (a `web.config` outbound rule;
  `tools/serve.cs` the same). An `.lgf` on the author's own site next to the `.html` is
  same-origin and needs nothing. An `.lgf` on a third site plays only if that site sends
  the header, which GitHub Pages and most static hosts do; a `.lgf` in a GitHub
  repository can be fetched through `raw.githubusercontent.com`, which sends it too.
- A page opened from disk (`file://`) can't fetch files at all in Chrome and Edge, so
  there only the inline form works. "Save as .html" writes the inline form for that
  reason.
- Nothing the player loads is executable but `player.js` itself, so a drawing from an
  untrusted URL can't run code: the player parses it as XML with `DOMParser` and never
  puts drawing text into `innerHTML`. Label text is drawn with `fillText`.
- The host page's Content Security Policy, where it has one, must allow scripts from
  livegeometry.com and `connect-src` for the `.lgf`'s origin; that is the host author's
  setting and is mentioned in the export's help text.
- **Versioning and caching**: a snippet pasted into a post must play in five years, while
  the `.lgf` format may change freely (see AGENTS.md). So the snippet pins a player major
  version (`/player/1/`); when the format breaks, `/player/2/` starts and the last build of
  `/player/1/player.js` is committed as a static file and served forever. Within a version
  the file is served `no-cache` with an ETag (it is small), so fixes reach old embeds. The
  gallery drawings are rewritten with the format, as the rule says; old embeds keep their
  old player.
- **web.config**: `Access-Control-Allow-Origin: *` on `/player/` and `/gallery/*.lgf`
  (scripts themselves need no CORS; fetched files do); `.lgf` as `application/xml`; cache
  rules as above. `tools/serve.cs` gets the same headers and MIME type.

## Keeping it in sync with the C# code

- **Parallel files**: see "Architecture". The rule, to go into AGENTS.md: a new figure
  kind, a new expression function, a new file attribute, a change to a `Recalculate`, a
  `HitTest` or a sampler is made in both places in the same change, or the player's
  exclusion list says why not.
- **A structural check**, `tools/playerparity.cs`, run by `tools/regression.cs`: every
  concrete `IFigure` type the deserializer can read (`DrawingDeserializer.FigureTypes`)
  has a `src/figures/**/<name>.js` whose class is registered under the same element name,
  or is on `Main/Player/excluded.txt` with a reason; every public static method of
  `Functions` and every `Math.cs` method the port lists exist in JS by name. It fails the
  suite when a C# figure appears without its twin.
- **A numeric check**: `LiveGeometry.Desktop.exe --dump <folder> <out>` writes, per
  drawing, a JSON of every figure (name, kind, exists, the numbers that define it: point
  coordinates, a line's two points, a circle's center and radius, a label's text, a curve's
  sample count and a few samples) at the file's viewport and a fixed canvas size, then
  again after a scripted drag (every free point and slider moved by a fixed offset). The
  player's `dump()` produces the same JSON; `tools/playerparity.cs` runs the player over
  the gallery and the curated GeoGebra imports in headless Edge (the CDP code of
  `webauto.cs`) and compares with a tolerance. Where the two must differ (label sizes come
  from different fonts, so pinned captions and label orbits differ by pixels) the dump
  leaves the number out or the tolerance says so.
- **Pictures**: `--check`'s PNGs beside the player's screenshots at the same size on one
  contact sheet, for the eye.

## Known parity traps

- `Math.Round` goes through `decimal` and rounds halves away from zero as written (2.675
  is 2.68): the JS port rounds on the 15-significant-digit decimal string, and the parity
  check covers labels.
- `ToLogical` rounds to 10 decimals (`RoundToEpsilon`); the port does the same, or
  intersections and `exists` flip differently.
- Dash patterns are in stroke widths in Avalonia and in pixels in Canvas: the port
  multiplies as `LineStyle.OnApplied` does.
- `Flipped` circles: `GetNearestParameterFromPoint` negates the angle and
  `GetPointFromParameter` doesn't (a C# oddity the survey found; port as is, fix in both
  if it is a bug).
- The locus and the function graph sample by the window (pixels), so their points differ
  between canvas sizes by design; compare at one size.
- Tabular figures (`tnum`) have no Canvas equivalent; numbers in captions may jitter by a
  pixel while dragging.

## Phases

1. **Skeleton**: folder, file list, dev page, the host element, `CoordinateSystem`,
   the XML reader with styles and defaults, FreePoint, Segment, Circle, MidPoint, Label
   (plain text, pinned, backdrop), the renderer, dragging free points, wheel zoom, pan.
   Plays `Bezier.lgf` (minus the curve) and Pythagoras.
2. **Numbers**: the expression language whole (scanner, parser, tree builder, bound tree,
   functions, binder), PointByCoordinates, labels with `[...]`, Number, Slider,
   ShowHideControl, gradients and Dark, Scenes. Plays Fireworks, Pick's Theorem, the
   Castle.
3. **Figures**: every line kind, the four intersections, PointOnFigure on every figure
   kind, the transformed points, circles, ellipses, arcs, polygons, regular polygons with
   parts, vectors, measurements, angle marks, right angle marks, decorations, the grid and
   axes. Plays most of the gallery.
4. **Curves**: Bezier, FunctionGraph, Locus (the recomputed chain), BezierPath with the
   smoother and holes. Plays all of it.
5. **Export and deploy**: the Export menu items, the embed page, web.config, the Browser
   csproj build step, the version folder, the smoke test.
6. **Sync tooling**: `--dump`, `tools/playerparity.cs`, the AGENTS.md section, the
   regression hook; and the cleanup of the Silverlight-era `#if !PLAYER` / `TABULA`
   blocks in `DynamicGeometry`, so that "player" in the codebase means this one.
7. **Polish**: touch, cursors, `data-theme`, resize, Twemoji option, a Mac check.

Phases 1-4 are the port and the bulk of the work; 5 and 6 are small but are what makes it a
feature rather than a demo. Each phase ends with the gallery drawings it claims playing in
the dev page next to the desktop app.

## Decided (2026-10-07)

- A JavaScript port; no .NET, no Avalonia, no external dependencies.
- No iframe. Inline `.lgf` or a URL to an `.lgf`; the gallery's drawings are served from
  livegeometry.com as `.lgf`.
- The browser's own emoji font; no fonts shipped.
- Classic scripts concatenated by MSBuild, no node in the build.
- Folder `Main/Player/`.

## Still open

1. The version-folder scheme for embeds (`/player/1/player.js`, old versions kept
   forever, `no-cache` within a version) - or a single `/player/player.js` whose reader
   keeps every old format, against the repo's no-compatibility rule. Recommended: the
   version folder.
2. Dark theme as an embed option (`data-theme`), recommended yes: the data is in the
   files already.
3. The mouse wheel: in a blog post an embed that zooms on wheel hijacks the page's
   scrolling. Recommended: wheel zooms only with Ctrl held (as maps do), pinch and
   double-click fit always; a `data-wheel="zoom"` opt-in for a page that is the drawing.
4. A small "Live Geometry" mark in a corner linking to livegeometry.com, perhaps "open in
   Live Geometry" with the drawing handed over. Later, if at all.
