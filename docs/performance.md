# Measuring speed

- The browser ran .NET interpreted until 2026-10-07 (see AOT below; the interpreter is one
  property away): the same code took 3-4 times as long there as in the Debug desktop build
  (2026-10-03, reading the gallery's drawings), and a loop that costs nothing on the desktop
  can take a second there. Measure in the browser before deciding what is slow. Reflection is the worst of it: `GetAttribute` (`Utilities`) and
  the property lists of `PropertyDiscoveryStrategy` are cached per member and type, and a
  save compares the drawing's styles with default styles made once per `AppTheme.Version`
  (`StyleManager.IsUnchangedDefault`). Before that (2026-10-04) a save of a triangle took
  5 ms on the desktop, nine tenths of it in the styles.
- CPU samples of the desktop app need neither admin rights nor a tool: start it with
  `DOTNET_EnableEventPipe=1`, `DOTNET_EventPipeOutputPath=<file>.nettrace` and
  `DOTNET_EventPipeConfig=Microsoft-DotNETCore-SampleProfiler:0:5,Microsoft-Windows-DotNETRuntime:0x4c14fccbd:5`
  in its environment, close it normally, and read the stacks with the TraceEvent package
  (`TraceLog.CreateFromEventPipeDataFile`, then `CallStack()` of each `Thread/Sample` event).
- Better, with an elevated shell (ETW needs admin on Windows; not on a Mac): the Ultra
  profiler (`dotnet tool install -g Ultra`) and the UltraMcp server (`dotnet tool install -g
  ultramcp`, an MCP server that reads its traces). Build Release, then
  `ultra profile --duration 15 -o <name> -- <...>/bin/Release/net10.0/LiveGeometry.Desktop.exe [--gallery <slug>]`
  in a scratch folder: it samples every thread 8190 times a second (kernel, native and
  managed stacks, JIT and GC markers) and writes `<name>.json.gz` in the Firefox Profiler
  format; the app stays open when the time is up (close it). The MCP tools read the `.gz`
  as it is: `list_threads` (thread 0 is the UI thread), `call_tree` with `focus` and
  `startMs`/`endMs`, `call_tree_inverted` for who calls a hot function, `find_hotspots` to
  compare named functions between two traces of the same scenario. A sample is 0.122 ms of
  CPU. `NtTraceEvent` at the top of a leaf list is the profiler's own cost. The site's
  speed is what counts, and the desktop is only a proxy for it: read a trace for the work
  our code does, not for the JIT (the browser interprets), ReadyToRun, or the start of the
  GPU (ANGLE, Direct3D: the browser has WebGL), and check a fix in the browser.
  2026-10-03, startup into the gallery at 1700x1100: the window's first frame at about
  1.7 s (half of it the JIT: neither our assemblies nor Avalonia's are ReadyToRun), then
  the tiles in view, 0.85 s of the UI thread's CPU (1.73 s before that day's fixes).
- In the browser: a `Stopwatch` around what is suspected and `Console.WriteLine`, in a
  Release publish served by `tools/serve.cs`, read with `webauto console`. Each `webauto`
  call is a process that takes CPU from the page: compare runs made the same way.
- How a page loads, as F12 shows it: `dotnet run tools/loadperf.cs -- <url> <out folder>
  [--seconds 20] [--size 1700x1000] [--trace] [--phone] [--wifi]`. A first visit (a new Edge
  profile, nothing cached) and a returning visitor's (a new browser on the same profile):
  the downloads, when .NET runs (the build line in the console), Avalonia's first frame and
  when the app took the splash down, and the long tasks of the page's thread - each load of
  a gallery tile is one. `--trace` records the first visit for F12's Performance panel (Load
  profile) and saves what was on screen every half second (`contactsheet` puts it on one
  image); `--phone` is a phone's screen, a 4 times slower CPU and Fast 4G (`--wifi` leaves
  the network alone: a local publish has no brotli, so its 4G download says nothing about
  the site's; the CPU part does). Nothing else may run meanwhile: a browser busy on the same
  machine (a page refreshed by hand, a headless one left behind) made every visit 1.5 to 2
  times slower. livegeometry.com on 2026-10-03, 1700x1000: on a first visit the splash
  closed at 2.7-3.0 s and the tiles in view were drawn by 6.2-6.6 s; for a returning
  visitor, 1.3-1.4 s and 4.6-4.8 s.
- How the gallery scrolls: `dotnet run tools/scrollperf.cs -- <url> <out folder> [--phone]
  [--wheel] [--passes 3] [--wait 15] [--dpr 1]`: the page loaded and left alone until its
  tiles are in, then scrolled down (loading the tiles it reaches) and back up (over loaded
  tiles: the scrolling alone), each with its frame intervals, the main thread's time in
  timer and animation-frame callbacks (Avalonia's dispatcher and render pass run in those)
  and its long tasks. `--phone` is 390x645 and a 4x CPU throttle, at a pixel ratio of 1:
  under the emulation Avalonia measured its canvas in physical pixels and divided by the
  emulated ratio, and laid the page out at a third of its width.
- Where the time went on 2026-10-07 (a local publish, headless Edge, returning visitor,
  1700x1000; the runtime is up at 0.2 s and everything below is after that): Avalonia's
  platform init 0.07 s, FluentTheme 0.03 s, the gallery page built 0.1 s, its first layout
  and frame 0.2 s (the frame itself 0.12 s of text shaping and Skia, native: the same under
  AOT), then the tiles in view 2.6 s in all, Stretchy Slime 0.47 s of it (the XML 0.03 s,
  reading the figures 0.13 s - the expressions compile there - adding them to the drawing
  0.19 s). The emoji font's parse is 1 ms. Under the 4x throttle every managed part is 4x:
  the first frame 3.1 s after the runtime, Stretchy Slime 2.6-3.2 s. Scrolling over loaded
  tiles at the 4x throttle: a p95 frame of 14 ms, where it was 40 ms before the tile
  distances were made arithmetic (above).
- AOT (`RunAOTCompilation` in the Browser csproj, on since 2026-10-07; set it to false, or
  publish with `-p:RunAOTCompilation=false`, to go back to the interpreter): the download is
  9.9 MB brotli instead of 5.3 (`dotnet.native.wasm` 35 MB raw, 7.2 MB brotli, which the
  browser also has to compile), the tiles load 2-3.5x faster (Stretchy Slime 0.24 s), the
  startup 1.3-1.7x, the gallery under the phone throttle is in at 6 s after the runtime
  instead of 12, and scrolling over loaded tiles takes 15% of the main thread instead of
  27%. A first visit on 4G pays about 4 s more for the download; a returning visitor has it
  cached. Only a publish compiles ahead of time (about 3 minutes more on this machine);
  `dotnet build` and `dotnet run` never do. Switching the property between two publishes
  on one machine wants the Browser project's `bin` and `obj` deleted in between: published
  the other way into the same `obj`, the runtime's native image and the assemblies didn't
  match, and the page died while loading the runtime ("MONO interpreter: NIY encountered",
  `mono_wasm_load_runtime () failed`). CI starts clean. A profile-guided partial AOT (only
  the methods a gallery load runs) would be the next thing to try for most of the speed at
  less of the size.
