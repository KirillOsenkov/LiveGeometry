# Notes for agents working in this repo

Persistent working notes. Keep them factual and short; update or delete entries that stop
being true. Things derivable from the code or git history don't belong here. This file holds
what every task needs; the rest is in `docs/`, a file per topic (see "Topics").

## Layout

- `Main/Avalonia/` - the live codebase. Solution `LiveGeometry.Avalonia.slnx`:
  - `LiveGeometry/` - the geometry library (assembly and namespace `LiveGeometry`, the namespace
    shared with the app; it was `DynamicGeometry` until 2026-10-15), a fork of
    the WPF sources rewritten onto Avalonia, plus `WpfCompat/` shims.
  - `LiveGeometry.App/` - shared UI (`MainView`), used by both heads (assembly
    `LiveGeometry.App`, namespace `LiveGeometry`: its `App` class is `LiveGeometry.App`).
  - `LiveGeometry.Desktop/` - desktop head (net10.0).
  - `LiveGeometry.Browser/` - WASM head (net10.0-browser), deployed to https://livegeometry.com.
- `Main/DynamicGeometryLibrary`, `Main/WPFClient`, `Main/Silverlight*` - the older WPF/Silverlight
  generation. Reference only.
- `Reference/VB6/Source/` - source of the original VB6 app ("DG"). A built copy may exist outside
  the repo (process name `Geometry`).
- `tools/` - file-based C# tools (`dotnet run tools/x.cs -- args`), see `docs/testing.md`.
- `CatrionaAgg/` - the library of Catriona Agg's puzzles at /CatrionaAgg (see
  `docs/catriona-agg.md`).

## Working rules

- The maintainer reviews and commits everything: never commit, push, or otherwise change git
  state. Leave changes in the working tree, no need to report them.
- Scripting is C# file-based apps (`dotnet run file.cs`). Don't assume Python is installed.

## Code style

- **Apply these rules to new code only.** Keep diffs minimal - don't reformat or rename existing
  code as a drive-by. Most of `LiveGeometry/` is decades-old WPF/Silverlight-era code; it
  stays that way unless a cleanup is explicitly asked for (2026-10-15: every file is
  file-scoped, done in bulk through Helix's `apply_code_action` with `scope: project`).
- **New files: CRLF line endings, file-scoped namespaces.** Existing files keep what they have.
  The Write tool emits LF: after creating a file, or rewriting an existing CRLF file with Write
  rather than Edit, convert its line endings with an editor tool. Never grep for `\r`.
- **Never edit workspace files from the shell** (sed/perl/heredocs). Use the editing tools.
- **Backwards compatibility is not a concern**: every caller is in-tree, so change all call
  sites rather than keep a worse shape. That goes for the `.lgf` file format too: as of
  2026-09-29 there are no drawings to stay compatible with other than the gallery's, so the
  format may change freely as long as the gallery drawings are updated to it (edit them, or
  `--rewrite` the folder) and still load. Don't add upgrade paths for old files; the ones the
  loader has are from before this rule.
- **Always brace** single-statement `if`s, even one-liners.
- **No two consecutive blank lines.** One blank between members, none at the top of a block,
  none before `}`.
- **Blank line after a block** - control-flow statements and member/type declarations alike -
  unless it is the last thing in its enclosing block.
- **Expression-bodied getters for trivial properties**: `public int Count => items.Count;`,
  not a `get { return ...; }` block; a getter whose expression wouldn't fit on one line keeps
  its block. Same for a setter next to such a getter (`set => field = value;`).
- **Collection expressions over array initializers** where the target is an array: `int[] x =
  [1, 2, 3]`, not `new[] { 1, 2, 3 }`, and `[]` for an empty one. Where the target isn't an
  array - `foreach (var s in new[] { ... })`, an `IEnumerable<T>` parameter, one arm of a
  conditional - `new[]` stays: a collection expression there has no type, or gets a
  read-only wrapper instead of an array. (`Main/Avalonia/.editorconfig` says both of these
  to the IDE, `when_on_single_line` and `when_types_exactly_match`; the code was converted
  on 2026-10-15 through Roslyn's fixes, IDE0025/27 and IDE0300/01.)
- **Before finishing a file**, sort usings and remove unused ones.
- **Avoid `sealed`. Avoid `internal`. Prefer `public`.** Only with a concrete reason.
- **No abbreviations in identifiers** (`operation`, not `op`). Exceptions: `sb`, `ex`, `sw`.
- **No `Async` suffix on async methods.**
- **More than 4 parameters: one per line**, in declarations and at call sites.
- **Name literal arguments at call sites**: `overwrite: true`, `parent: null`, `radius: 0.32` -
  a bare `true`/`null`/number in an argument list is opaque. Self-describing expressions don't
  need it. (Runs of x/y coordinates, as in `IconBuilder` calls, are fine bare.)
- **MSBuild conditions: no quotes** around property names or `true`/`false`:
  `Condition="$(Configuration) == Debug"`. Quote only values that can be empty or contain spaces.
- **No global.json.** The repo must build with whatever new-enough .NET SDK is installed.
- Build with `dotnet build` / `dotnet publish` (there is no WPF markup compilation in the
  Avalonia solution). Pass `-bl` and read the binlog with the binlog MCP
  tools when a build misbehaves, instead of parsing console output.

## Build and run

- Desktop: `dotnet build Main/Avalonia/LiveGeometry.Desktop/LiveGeometry.Desktop.csproj`, then
  run `bin/Debug/net10.0/LiveGeometry.Desktop.exe [drawing.lgf|.dgf]`. Fastest loop; use it for
  feature work. A drawing can be hand-written as XML (`docs/*.lgf` are examples: styles, then
  figures by type name with `<Dependency Name=...>` children; a point's name shows through a
  separate `PointLabel` figure that depends on it) and opened from the command line to check it.
  On a Mac the executable is `bin/Debug/net10.0/LiveGeometry.Desktop` (no `.exe`), and the
  same build runs as is: see `docs/macos.md`.
- Browser: needs the `wasm-tools` workload (`WasmBuildNative=true` links Skia/HarfBuzz; without
  it Skia throws DllNotFoundException). Do a full bin/obj clean when native assets change.
- Browser publish: `dotnet publish Main/Avalonia/LiveGeometry.Browser/LiveGeometry.Browser.csproj -c Release -o <dir>`.
  Output is `<dir>/web.config` + `<dir>/wwwroot/`. CI (`.github/workflows/main_livegeometry.yml`)
  does exactly this and deploys to Azure App Service (IIS). A publish compiles the .NET code
  ahead of time to wasm (`RunAOTCompilation` in the csproj, a few minutes more; see
  `docs/performance.md` for what it buys and costs, and how to switch it off).

## Topics

The rest of the notes are in `docs/`, a file per topic: read the ones a task touches before
starting on it. Entries cite each other by their bold title in quotes (see "Undo"): search
`docs/` for the title to find one. What is learned goes into the file of its topic; a new
topic gets a file of its own, listed here.

- `docs/tools.md` - the ribbon, tab by tab: what every tool does and how it takes its clicks.
- `docs/canvas.md` - the view (zoom, pan, fit, grid, axes), what a click makes of a point,
  Tab among overlapping figures, the reach of a click, cursors, snapping, joining and
  releasing points, dragging, tool letters, Z order, selection halos.
- `docs/figures.md` - the dependency graph and its order; crossings with curves, points inside
  shapes, regions, tolerances, squares and regular polygons, lengths, ellipses, arcs, sectors
  and their sweep, curves with gaps, vectors, segment and right angle marks, angles, parts of
  a figure.
- `docs/bezier-paths.md` - Bezier paths: anchors, handles, automatic smoothing, holes, images.
- `docs/values.md` - setting and fixing lengths and directions, Numbers, sliders, tied values,
  translated points.
- `docs/expressions.md` - the expression language, renames, cycle checks, exact forms, how
  expressions are bound and evaluated, building a formula by clicking, coloring, completion.
- `docs/labels.md` - default names and subscripts, labels that say numbers, wrapping, names of
  lines and circles, attached texts, names keeping clear, pinned labels, editing in place.
- `docs/property-grid.md` - the side panel: rows, editors, verbs, the order of a page,
  quantities, where the panel goes; tool panels with their OK and Enter; typed text and its
  errors; context menus.
- `docs/styles.md` - default styles and the palette, coincident lines, point shapes and emoji,
  dashes, textured fills, the color picker, the paper, the app's theme and how drawings
  follow it.
- `docs/undo.md` - what undo must hold, every trap met so far, and how to audit it.
- `docs/app.md` - the Figure List, Save, Export, share links, the toolbar, settings between
  runs, the ribbon's look, the splash.
- `docs/avalonia.md` - Avalonia and browser traps: trimming, inherited attributes, input and
  touch, collections, infinite coordinates, layout at no size, render-pass freezes, fonts,
  culture, files in the browser, focus and keyboard shortcuts, how every exception is shown
  and reported.
- `docs/file-format.md` - the `.lgf` format: colors, versions, styles in files, reading a
  damaged file, label text, show/hide boxes, scenes.
- `docs/geogebra.md` - the GeoGebra (.ggb) reader; `docs/dgf.md` - the DG 1.0 (.dgf) reader.
- `docs/gallery.md` - the gallery: routes, captions, tiles, which drawings are generated and by
  which tool, the batches of drawings and what was learned making them.
- `docs/player.md` - the JavaScript player (`Main/Player/`). A new figure kind, a new function
  of the expression language, a new file attribute, a changed `Recalculate`, `HitTest`,
  `ReadXml` or sampler goes into the player in the same change: read it before any of those.
- `docs/catriona-agg.md` - the Catriona Agg library at /CatrionaAgg.
- `docs/deployment.md` - deploying, caching, web.config, link previews, the static pages
  (/history, /embed, /web).
- `docs/macos.md` - what is the Mac's own.
- `docs/testing.md` - `--check` batch runs, `tools/regression.cs`, winauto, macauto, webauto,
  and the web smoke test before deploying.
- `docs/performance.md` - measuring speed on the desktop and in the browser, AOT.
- `docs/backlog.md` - the VB6 parity backlog, and what is not yet verified in the browser.
- Plans and ideas: `docs/player-plan.md`, `docs/completion-plan.md`, `docs/gallery-ideas.md`.

## Running instances

Always build into the normal `bin/` - no scratch output folders. The maintainer rarely has the app
running and doesn't keep state in it that matters, so a `LiveGeometry.Desktop` process that locks
`bin/` is almost certainly a leftover test instance, and killing it is fine in most cases.
Only if a build actually fails on a lock, enumerate and clean up:

`Get-Process LiveGeometry.Desktop | Select Id, StartTime, MainWindowTitle, Path` (pwsh), then
`(Get-Process -Id N).CloseMainWindow()` and, if it is still there a few seconds later,
`Stop-Process -Id N`. Alt+F4 through winauto can miss (it goes to whatever window of the app has
focus, e.g. a popup), so check that the process is really gone. Target test windows by `pid:`/`hwnd:`
rather than by process name when more than one could exist.
