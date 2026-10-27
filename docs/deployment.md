# Deployment and caching

- A push to main takes about 8 minutes to go live (GitHub Actions: build ~5, deploy ~3); the
  files flip at the very end. Until then every reload, cached or not, gets the previous version.
  The commit in the tooltip of the Octocat at the right end of the toolbar (`BuildVersion`, also
  logged to the console at startup) tells which build is on screen.
- Entry files are `no-cache` with an ETag that changes on every deploy, so a browser holding the
  previous version picks up the new one on a plain reload. If something looks stale, check the
  commit stamp and the Actions run before suspecting the cache.
- **Link previews** (2026-10-24): a site that shows a link's preview (X, Facebook, Discord,
  Slack, a search engine) reads the page's HTML and never runs its script, so every
  `/gallery/<slug>` was the same "Live Geometry" card. The `WriteGalleryIndex` task of the
  Browser csproj also writes `wwwroot/gallery/<slug>.html` for each catalog item: the app's
  `index.html` with the drawing's title, the first sentences of its explanation that have no
  `[...]` part, `og:image`, `twitter:card` and a canonical link in the head, and the title and
  sentences on the splash (what a search engine indexes). It fails the build when
  `index.html` no longer has one place for each tag it replaces. `web.config` serves the page
  for `/gallery/<slug>` (`tools/serve.cs` does by its `<path>.html` rule); it writes
  `/sitemap.xml` too, which `wwwroot/robots.txt` names. The pictures are made after the
  publish, in CI: `dotnet run tools/previews.cs -- <publish>/wwwroot` plays each drawing in
  the player in headless Chrome (Edge on Windows) as its tile shows it, on its paper, into
  `gallery/previews/<File>.png` (1200 x 630), revalidated like the `.lgf` files. A new
  drawing needs nothing more than its catalog line. A local publish has the pages but no
  pictures until the tool is run on it. To see what a site sees after a deploy: the
  page's source (`curl -s https://livegeometry.com/gallery/<slug>`) and each site's card
  validator. The homepage's own picture (2026-10-27) is `gallery/previews/index.png`,
  drawn by the same tool: the gallery page as it opens under the light theme, the brand
  over the first eight tiles (their plates, pictures and captions as `web/index.html`
  draws them). `wwwroot/index.html` names it, with the whole set of tags the Catriona Agg
  pages have (`og:site_name`, the picture and its size, `twitter:card`, a canonical link),
  and the per-drawing pages are written by replacing the picture and the canonical link in
  it. `/embed` and `/web` name the same picture, and `/web/<slug>` gets a page of its own
  like `/gallery/<slug>` (`web/index.html` with the drawing's tags, `obj/web/pages`,
  linked into `wwwroot/web/`, a `web.config` rule before the one that answers every
  `/web` path with the page); the `/web` pages are not in the sitemap, `/gallery/<slug>`
  being the drawing's address. `/history` has only the title and description. To draw the pictures
  locally the tool wants a folder with `gallery/index.json` and the `.lgf` files,
  `player/1/player.js`, `fonts/Twemoji.Mozilla.ttf` and `logo-512.png`: a publish's
  `wwwroot`, or one put together from `obj/` and `wwwroot/` of the Browser project.
- **`/history`** is a static page of its own (`history/index.html` at the repo root, no build
  step, everything inline but the Google Fonts; a relative URL in it would resolve against
  `/`, since `/history` has no trailing slash): the Browser csproj links the folder into
  `wwwroot/history/`, and a `web.config` rule answers `/history` with its `index.html` (a folder
  is not a file, so the SPA fallback would otherwise serve the app). `tools/serve.cs` has no
  such rule: locally open `/history/index.html`, or serve the `history` folder itself.
  `/history/tabula` (2026-10-09) is a wing of it, `history/tabula.html`, with its pictures
  in `history/tabula/` (a `web.config` rule answers the path with the page, and `serve.cs`
  serves `<path>.html` for an extensionless path that has one); it references them by
  absolute path, since a relative one would resolve against `/history/`. Everything on it
  came from the Wayback Machine's captures of numeracyworks.com and from YouTube.
- **`/embed`** is the same kind of page (`embed/index.html` at the repo root, linked into
  `wwwroot/embed/`, a `web.config` rule): the hub for putting a drawing on another page,
  itself played by the player it describes (`/player/1/player.js`, the gallery's `.lgf`
  files). Its picker lists the gallery from `/gallery/index.json`, which the Browser csproj
  writes at build from `GalleryCatalog.cs` (`WriteGalleryIndex`: the `Item("slug",
  "Title"[, "File"])` lines, in order), so the list never has to be kept in step by hand.
  Its sun and moon read and write the app's `LiveGeometry.Theme` entry, so the choice is
  one for the site, and set every player on the page (`applyTheme`). **The player plays
  before DOMContentLoaded** (a deferred script runs once the parser is done, and the
  player plays every element right then), so nothing of the page may be left to a
  DOMContentLoaded listener that must come first: the drawings' `data-theme` is set by the
  page's inline script as it runs, and the hero's load waits for the player when the
  catalog's fetch comes back first (2026-10-26: the two finish within a millisecond of
  each other on the site, and the hero was blank on about one reload in a few). To work on it it
  needs the player and the drawings at those absolute paths: serve a publish's `wwwroot`
  with the repo's folder overlaid and live reload,
  `dotnet run tools/serve.cs -- <publish>\wwwroot 5006 --overlay embed=embed --reload`,
  and open `/embed/index.html`; every save of the page reloads the browser. A change to
  the player or the catalog still wants a new publish (or its bundle copied into the
  publish's `player/1/`).
- **`/web`** (2026-10-08) is the gallery and its drawings without .NET: `web/index.html` at
  the repo root (linked into `wwwroot/web/`), played by the player. `/web` is the gallery
  page as the app draws it (the New and Open tiles, which lead to the editor at
  `/drawing`, the brand, the tiles: live drawings in a `TilePlayer`, a `Player` whose
  canvas is 560 x 380 scaled down through the pixel ratio, the caption left out of the
  XML as `DrawingThumbnail.LeaveOutCaption` does, drifting on hover, loaded by distance
  from the view a few per frame as `GalleryView.LoadNextTiles` does); `/web/<slug>` is
  the app's toolbar over the drawing - the mark (inert), Gallery, New and Open (links to
  `/drawing`), the tour group (Page Up/Down), the sun and the Octocat - and nothing the
  player can't do (no ribbon, Save, Export, Undo, Settings, status bar). Routes are the
  page's own (`pushState`); `web.config` answers every extensionless path under `/web`
  with the file, and `tools/serve.cs` answers an extensionless path under any folder with
  an `index.html` with that file (a publish has `web/<slug>.html` per drawing for the link
  previews, see "Link previews", which both serve first). The catalog comes from `/gallery/index.json`, which now
  also carries `stackedFigureShare`, passed to the player as an option. The player exports
  `AppTheme`, `GalleryDrawing`, `Label`, `ControlBase` and `LabelPin` for it. The page
  logs its timings to the console (`web: ...` lines: a drawing's fetch and load, each
  tile, and when the tiles in view are in), which is what it is for: a feel for the
  speed without .NET before deciding whether the gallery should move. To work on it:
  `dotnet run tools/serve.cs -- <publish>\wwwroot 5007 --overlay web=web`, with the
  current bundle and `index.json` copied into the publish (`dotnet msbuild
  LiveGeometry.Browser.csproj "-t:BuildPlayer;WriteGalleryIndex"` makes them in `obj/`),
  then `webauto start http://localhost:5007/web/`.
- **web.config**: `LiveGeometry.Browser/web.config` is hand-written (serves the precompressed
  `.br` files, sets immutable caching on fingerprinted assets, `no-cache` on entry files, and
  `max-age=0, stale-while-revalidate=86400` on the player, the logo and the Catriona Agg
  library's files, which a browser then uses from its cache while it revalidates them in the
  background: a page of the library renders without a round trip per file, and a deployment
  reaches a returning reader one page load late). The
  wasm SDK drops a project web.config from publish, so the csproj copies it with an explicit
  `AfterTargets="Publish"` target. It sits at the publish root and rewrites into `wwwroot\`.
  There is no IIS locally: verify changes after deploy with
  `curl -s -o /dev/null -D - -H "Accept-Encoding: gzip, br" https://livegeometry.com/_framework/<file>`
  and expect `Content-Encoding: br` and `Cache-Control: public, max-age=31536000, immutable`.
  A malformed web.config takes the whole site down (HTTP 500).
