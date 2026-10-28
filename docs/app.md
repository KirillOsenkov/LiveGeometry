# The app around the drawing

The Figure List, Save, Export, share links, the toolbar, settings between runs, the ribbon's look, the splash.

- **The Figure List** (`UI/FigureExplorer.cs`, left column of `DrawingHost` with a splitter):
  top-level figures but the grid and point labels, by `Title`, with the ribbon icon of the
  tool that makes each (`UI/FigureIcons.cs`: figure type -> tool, walking base types; a new
  figure kind wants an entry). Hidden and `Auxiliary` figures are faded. Its selection *is*
  the drawing's (`Selected` + one `RaiseSelectionChanged`); arrows in the left margin go from
  the keyboard's row to its dependencies (blue), and with `FigureExplorer.RecursiveArrows`
  (off; no UI for it) on up from those (gray), one trunk per figure, a step left of the
  trunks above it that it runs beside, its branches stopping short of the next trunk: so no
  two arrows cross (`LayOutTrunks`). It rebuilds on `ActionManager.CollectionChanged`
  (posted, coalesced, and through `Throttle` at most every 300 ms: a drag of a figure or the
  view records a move per step, and it rebuilt every row at every frame), never between `ConstructionStepStarted` and a complete step: temporary
  figures of a tool are never recorded and its real steps sit in its transaction. Undoing a
  deletion puts each figure back at its old index (`RemoveFigureAction.Indices`), so it
  doesn't jump to the end of the list. A hidden figure that is selected is shown ghosted
  (`ShapeBase.IsGhost`, half opacity, shape not hit-testable) until unselected; `Visible`
  stays false, so no hit test, snap or drag sees it (they all ask `Visible`;
  `HitTestShape` refuses hidden figures too). Only figures directly in the drawing: a
  composite's parts (a vector's hidden segment) never ghost, so a hidden composite shows
  nothing. `UpdateVisual` overrides that skip hidden figures test `IsShown` instead, or
  the ghost sits at a stale place. A figure hidden when it comes onto the canvas keeps
  its shape off the canvas until it first shows, Visible or a ghost
  (`ShapeBase.OnAddingToCanvas`, `PutShapeOnCanvas`, 2026-10-07; shapes proper only, a
  label is measured as shown while hidden and wants the tree's font): the hidden helpers
  of a construction and the variables of a generated drawing (193 of Stretchy Slime's 264
  figures) cost the canvas nothing at the load and at every frame. The shape is styled all
  the same, since code reads a point's size or a line's thickness off it hidden or not
  (`Arrow`, `CoordinateSystem.LimitZoomByPointReach`). `Actions.ReplacePoint` puts the replacement in the
  old point's place (`MoveBefore`: with what it is built on that came later, such as its
  Number, to keep dependency order; `Figures.Move` doesn't touch the canvas). Several figures selected show as a `FigureSelection`
  in the property grid (common properties + Delete).
- **Save writes the drawing's own file again, without a dialog** (`MainView.SaveDrawing`),
  when it has one: `OwnFile`, the .lgf it was read from (the Open dialog, the command line)
  or last saved to. A new drawing, a drawing of the gallery and one read from a GeoGebra or
  DG file have none and get Save as (`SaveDrawingAs`, also the first item of the Export
  menu), whose file is the drawing's from then on. A file that failed to load, or loaded
  only in part, is not kept (`Drawing.Name` is set only by a load that went through, see
  "A file is read as far as it goes"), or Save would write the ruins over it. A
  construction under way is put away before saving (its point following the cursor and
  its preview are figures while it lasts, and went into the file). In the browser the first Save of an opened file makes the browser ask for
  permission to write; a browser without the File System Access API gives files to read
  only, and Save falls back to Save as there (`IsReadOnlyFile`; not tried in a real
  Firefox). A file the desktop can't write (read-only, open in another program) is said so
  in the hint and Save goes on to Save as (`MainView.TryWriteFile`, Export's too): it was
  an error report, as for a bug. There is no unsaved-changes prompt and no mark of a
  changed drawing.
- **Export** (`MainView.Export.cs`, the button after Save; its menu is a `MenuFlyout`): Save
  as .lgf (see above), then Save
  as .png, Save as .svg, Copy image - the canvas as it is on screen at that moment, same view,
  same size, without the side panel and the status bar (siblings over the canvas, not in it;
  a selection's highlight is in it). `ViewportImage` makes both: the PNG has the pixels of the
  screen (layout size × the screen's scaling; the clipboard gets the same picture), the SVG is
  in layout units and is what Skia writes (`SKSvgCanvas`) when Avalonia draws the canvas onto
  it (`DrawingContextHelper.RenderAsync`), so anything that draws on screen is in it with no
  code per figure. Skia writes text as `<text>` in a font, and a viewer without that font -
  every viewer, for Inter and the emoji - draws its own: `SvgTextOutlines` replaces each
  `<text>` by the outlines of its glyphs and a color emoji by its COLR layers in their CPAL
  colors (version 0 tables, which Twemoji has), so the file needs no font. The glyphs are
  looked up by character (cmap), not taken from the shaped run: a ligature or contextual
  alternate would come out as its plain letters. A text whose font isn't found stays
  `<text>`. Skia's `font-weight` is one step too light from 500 on ("600" is a bold 700):
  `ReadWeight` undoes that, look at it again when SkiaSharp changes. Fonts are found by the
  family name Skia wrote: the emoji font, Avalonia's `fonts:Inter` collection, then the
  installed ones; a new embedded font wants an entry in `FontFaces.OpenEmbedded`. To check
  an SVG, open it in headless Edge (`webauto start file:///...svg`) next to the PNG.
  The last item of the menu is How to embed (2026-10-26): /embed is opened in the browser
  (`Launcher.LaunchUriAsync`: a new tab in the browser head). The hint after Copy embed
  code has a link to the same page: `DrawingHost.ShowHint(text, linkText, link)` is shown
  with a link in it (`StatusBar.Show`: a `HyperlinkButton` in an `InlineUIContainer`; the
  press is handled there, otherwise the bar would be hidden by its own click-to-dismiss
  handler before the release).
- **Share links** (2026-10-14; `MainView.Share.cs`, Export > Copy share link): the drawing
  as Save writes it is posted to the site's API, which stores it under an id made from its
  content and answers it; the link is `https://livegeometry.com/s/<id>`. The API is a
  service of its own. What the app relies on:
  - `POST {ShareService}/api/v1/share`, the `.lgf` as the body, `Content-Type: text/plain`
    (a browser sends that without a preflight) -> `200 {"id": "abcd2345"}`: 8 or more
    characters of `a-z2-7`, the same for the same bytes. `413` over 512 KB, `429` too many
    calls for now (with `Retry-After`), `400` not XML, `403` from a page of another site.
    The status says each in words; a failed request or a refused clipboard is no error
    report (`IsSharingFailure`, in `IsBenign` while sharing).
  - `ShareService` is a build property (`-p:ShareService=https://...`, the public
    workflow's Actions variable `SHARE_SERVICE`) that reaches the app as an
    `AssemblyMetadata` (`LiveGeometry.App.csproj`): a build without it has no menu item and
    never calls the API.
  - `web.config`: `/s/<id>` is the `/web` page, `/s/<id>.lgf` a 302 to where the drawing is
    stored (`no-cache`, so that the storage can move). Links and embeds name only these.
    `tools/serve.cs --shares <the storage, ending in />` answers them the same way locally.
  - The same drawing again (the same text) gets the same link without a request. The text
    includes the view (`<Viewport>`), which the player and the editor fit into their window:
    a pan makes another link.
  - Safari writes to the clipboard only during the click, and the link comes after a
    request: in the browser the write starts in the click with a `ClipboardItem` whose text
    is a promise (`copyTextWhenReady` in `main.js`, `CopyTextWhenReady`), resolved with the
    link. The desktop puts the link on its clipboard when it comes.
  - The page at the link (`web/index.html`, route `/s/<id>`) plays the drawing, with New,
    Open and Edit (the pencil of the My Drawing tile) in its toolbar. They hand over to
    the editor (`MainView.Handoff.cs`): the page puts `{ action, name, data }` (data in
    base64) into `sessionStorage` (`LiveGeometry.Handoff`) and goes to `/drawing`; `main.js`
    takes it out before the runtime starts, so a reload doesn't do it again, and
    `OpenStartupFile` does it: `new` is New, `open` is Open with the file picked on the page
    (the app can't open a picker without a click of its own), `edit` loads the drawing as
    the user's own without a file (Save is Save as). Each replaces the kept drawing, as New
    does. `/web`'s New and Open tiles and buttons hand over the same way.
- **No menu.** New, Open, Save, Export | Undo, Redo are one row (`LiveGeometry.App/MainToolbar.cs`) above the
  ribbon, with the tour group (◀ n/N ▶ + title) between them and the Octocat at the right,
  whose tooltip is the build. The first button is the app's mark and folds the ribbon (Ctrl+F1,
  `MainView.UpdateRibbon`): folded by default when a gallery drawing opens on a small screen
  (under 700x500), open otherwise; once pressed, the user's choice holds for the session.
  Everything else is keys and the context menus (Lock among them). Not on the Selection tab (obscure for the audience; the commands and
  settings are still there in `DrawingHost`): Ortho, Polar, Snap to grid, Snap to point, Snap to
  center. Shift while dragging or clicking still snaps to the grid, and a click near the middle
  of a segment or a vector still makes a midpoint (`PointPlacement.HasMidpoint`: a vector
  is a figure of its own around a hidden segment, and a check for `Segment` missed it). In a narrow window (a phone, under about 410 px)
  the buttons close up (`MainToolbar.SetCompact`: square, no spacing, tighter
  separators, about 330 px), or the last of them, the settings gear, was cut off; what is
  at the right (theme, Octocat) is dropped first when there is no room.
- **Settings between runs** (`LiveGeometry.App/SettingsStore.cs`): `Get`/`Set` by key, the desktop
  head keeping them as `key=value` lines in `%LocalAppData%\LiveGeometry\Settings.txt`
  (`~/Library/Application Support/LiveGeometry/Settings.txt` on a Mac;
  `FileSettingsStore`), the browser in `localStorage` under `LiveGeometry.<key>`
  (`BrowserSettingsStore`, two imports in `main.js`; `index.html` reads the `Theme` entry
  before the runtime starts, so the splash and the page are already dark). `AppSettings`
  (`[PropertyGridName("Settings")]`, the gear) is the page over it: `Theme` is `System` or a
  theme's name, applied in `App.Initialize` before the first frame. The window placement is
  the `WindowPlacement` line of the same file (`WindowBounds` on a Mac). Colors tweaked on the "Theme colors" page are
  kept per theme (`ThemeColors.Dark`), only those that differ from `AppTheme.cs`
  (`AppTheme.EditsToText`/`ApplyEdits`, one line, read without throwing), so an untouched
  color follows the code; they are put on before `AppTheme.Register`, and "Built-in colors"
  drops them. Writes wait for a half-second pause (a picker drag sets a color per move) or
  `SettingsStore.Leaving` (the window closing; the page hidden, from `main.js`'s
  `visibilitychange`). In the browser only (`MainView.KeepsOwnDrawing`), the user's own
  drawing is kept too (`MainView.KeptDrawing.cs`: keys `Drawing`, `DrawingName`): saved
  a second after the last change of the undo history (`Throttle`, which only posts the
  save to the UI thread) and on Leaving, never during a construction, removed when empty,
  and not written again when the text is what the store holds (saved at the next idle
  moment, it was saved at every move of a figure's drag, whose moves are merged steps, and
  the drag went a few frames a second);
  read at startup but loaded only when the user goes to it (/drawing, My Drawing). Until
  then it is their drawing and nothing is written over it; New or an opened file replaces it.
- **Ribbon look**: `ButtonGrid` draws the
  hover/pressed/checked plate; `Ribbon`/`TabPanel` replace the Fluent templates in code. To
  bring the group headers closer together, change `ButtonGrid.HeaderOverlap`, not the padding
  inside the tab. An on/off `Command` exposes `IsChecked` (a `Func<bool>`), which its button
  re-reads after any toggle (`CommandToolButton.UpdateToggles`, called by anything that toggles
  from a key) - don't go back to `CheckBox` icons.
- **The splash** (`wwwroot/index.html` + `app.css`) stays up until the app says it has drawn
  what the page opened at (`SplashScreen`, 2026-10-07): the gallery with its tiles in view
  loaded (`GalleryView.LoadNextTiles`) or a drawing laid out (`MainView.HideSplashWhenDrawn`);
  `main.js` adds `app-ready` then and the splash fades. Avalonia's own `splash-close`, added
  at its first frame, hides nothing any more: on the gallery page that frame came seconds
  before the tiles, which then popped in batch by batch. Two deadlines in `main.js` take the
  splash down should the app never say so: 10 s after Avalonia's first frame, 20 s after Main
  returns. It animates Euclid's first construction in HTML, every move a transform or an
  opacity (a ring is drawn by turning a half-colored ring behind a half-window, a line is ink
  sliding into a slot), which the browser plays on its compositor thread while the main
  thread is held by the app's startup; the SVG version's `stroke-dashoffset` and
  `offset-distance` were painted on the main thread and froze with it. The bar under it is
  the downloads for its first 60% (a `withResourceLoader` wrapper counting fetches: the
  loader's `onDownloadResourceProgress` is on the module config, which the .NET 10 host
  builder doesn't expose) and what the app reports for the rest (`SplashScreen.ReportProgress`:
  tiles in view loaded, out of all in view); it is a cover sliding off the gradient, not a
  width, for the same reason. The splash takes the pointer events while it is up (the app
  under it got the clicks). The wrapper hands back a Response only for assemblies, the wasm
  and ICU data; the runtime's own JavaScript modules and its config must be left to the
  default loading (a Response for `dotnetjs` leaves the site spinning on the splash forever).
  A change to `main.js` needs the publish smoke test below before deploying. `?splash` on
  the url shows the splash without starting the app: serve the source `wwwroot` with
  `tools/serve.cs` (`.claude/launch.json` has "splash-dev") and open
  `http://localhost:<port>/index.html?splash`. The intro's last paragraph links
  /CatrionaAgg, /web and /embed (2026-10-27): the app is a canvas, so these are the only
  links from the homepage a crawler can follow, and the per-drawing pages carry them too;
  /history is left out on purpose. It animates
  regardless of `prefers-reduced-motion` on purpose: the query follows the Windows "Animation
  effects" setting, which many machines have off, and a still splash looks stuck.
