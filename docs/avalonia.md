# Avalonia and framework traps

- **Trimming only happens on Release publish**, never in `dotnet run`. The geometry library is
  discovered via reflection (`GetTypes()` for behaviors, serializers, figures), hence
  `TrimmerRootAssembly LiveGeometry` in the Browser csproj; `LiveGeometry.App` is a root
  too (the property grid reads `ExceptionReport` by reflection, and `[DynamicallyAccessedMembers]`
  on the class did not keep its rows). A trimming break shows up as a white screen and a
  `CRASH: ...` console line (`Program.cs` prints those). Anything new that is reached only by
  reflection outside those assemblies needs its own root. The functions of the expression
  language are found by name (`ExpressionReflection`, methods of `Functions` by the names of
  `FunctionCatalog`, which the trimmer can't follow): `[DynamicDependency]` on its static
  constructor keeps them all. Without it (when they were System.Math's) the browser had
  only those the app also calls itself, and `asinh`, `sinh`, `cosh` were "Could not find
  method" - the Catenary drew no curve there, while the desktop (never trimmed) drew it.
- **Inherited attributes come twice in the browser**: Mono's `inherit: true` lookup on a
  property declared in a base class (`Circle.Length` from `CircleBase`) adds the base
  property's attributes again. With `AllowMultiple = true` they aren't deduplicated and
  `Attribute.GetCustomAttribute` throws AmbiguousMatchException; CoreCLR (desktop) returns
  one. An attribute read singly must say `AllowMultiple = false`.
- **StyleKeyOverride**: a subclass of a templated Avalonia control (TabControl, TabItem, ListBox,
  Button, UserControl, ColorPicker...) renders invisible unless it overrides `StyleKeyOverride`
  to return the base type.
- **Input**: all WPF-style mouse events funnel through `LiveGeometry/Behaviors/Behavior.cs`
  (pointer adapters live there). Avalonia has no static Keyboard; modifiers come from event args.
  Fingers (`PointerType.Touch`, "Touch" region there) don't go to the tool as they come: a
  finger's press is held back until it is a tap (lifted where it came down) or a drag (gone
  `TouchSlop` pixels: the tool gets the press where it came down, then the moves). Each
  finger is for a figure or for the paper, decided where it comes down
  (`Behavior.TakesFingerOnFigure`, 2026-10-11: the Drag tool says a figure wherever a press
  would drag or select one and the paper wherever it would move the view - empty paper, an
  axis, a caption, a graph with nothing to move: `Dragger.PrepareDrag`, what a press
  holds, worked out without a press; a tool that constructs says the paper for every
  finger). A figure finger is the tool's for as long as it is down and
  drags its own figure, whatever the other fingers do; its press goes to the tool at once
  (unless a tap there would have to ask what it means), since held back until the slop
  it found the figure no longer under the finger once the paper fingers had moved the
  view, and the Drag tool waits for `TouchSlop` itself before a finger's press is a
  drag. The Drag tool holds a `Drag` per
  pointer (the mouse is one), so two fingers drag the two ends of a segment at once, in one
  undo transaction for all the fingers down together, from the first drag that moves
  something to the last release (transactions nest and can't be interleaved; and the
  moves of two fingers on one figure come interleaved and can't merge into one action, so
  every drag opens the transaction, not only a point's). The paper fingers are the tool's
  only while there is one finger on the canvas (a tap selects, a drag pans through the
  tool, recorded); with any other finger they move the view themselves
  (`CoordinateSystem.PanAndZoom`: one pans, two pan and zoom, the paper under them stays
  under them, a third does nothing), the tool hears nothing of them, and the held figures
  get their last move again after each step, so that a held point stays under its finger
  while the paper moves beneath it. Passed on as they came, a second finger was a second
  press: the Drag tool's view jumped between the fingers, any other tool left points
  behind; and until 2026-10-11 a second finger always made a pinch of the two and ended
  the first finger's drag. While a tool handles a finger, `Math.CursorTolerance` is
  `TouchTolerance` (10 px, a finger reaches further than a mouse); whatever sizes things in
  pixels must not take the cursor's tolerance for it (`PointLabel.Margin` does now: a label
  placed by touch sat further out). No long press yet: what the right button does (the
  context menu) has no touch equivalent. `webauto tap/touch/pinch` drive it headless.
- **Mutating a collection in place does not redraw in Avalonia** the way a WPF Freezable did.
  `Polygon/Polyline.Points` only rebuild geometry when the property gets a different list;
  figures call `Shape.PointsChanged()` (`WpfCompat.cs`) instead. Suspect the same thing whenever
  something "renders once and then never updates". Likewise a `Path` whose figure list became
  empty is not repainted: keep the first `PathFigure` and empty its segments (`AngleArc` does).
- **`PathFigure.IsClosed` defaults to true in Avalonia** (false in WPF), and `IsFilled` to true.
  Every hand-built `PathFigure` must set both; a forgotten one draws a line from the end of the
  path back to its start.
- **A point with an infinite coordinate does not exist** (`Math.Exists(Point)` rejects NaN
  and infinity; `IsValidValue` for a double). Avalonia's layout throws "Invalid Arrange
  rectangle" for a canvas child placed at infinity, unhandled: the desktop app dies and the
  browser app freezes on the spot. `CenterAt`/`MoveTo` in `Utilities` also refuse a non-finite
  place. Anything new that positions a control from figure coordinates must not hand it NaN
  or infinity. A size neither: `Width` throws on infinity right in the setter, and a
  figure's `UpdateVisual` runs also while it doesn't exist (a circle around a center that
  is nowhere has an infinite radius: the exception ended a drag, which then could not be
  undone; `CircleBase`/`EllipseBase.UpdateVisual` return early).
- **The page can have no size**: a browser lays it out at 0x0 (a tab opened in the
  background, DevTools taking the whole height), and a desktop window shrunk to its title
  bar leaves a canvas of a few pixels. Arithmetic on a layout size (`finalSize.Height -
  top`, a room beside a caption) goes negative there, and Avalonia's `Arrange` throws
  "Invalid Arrange rectangle" (the toolbar's tab did, 2026-10-06); a scene chosen for a room
  of no area is null. Code that lays out or fits by sizes must survive 0 and a few pixels,
  and a fit must not depend on the view it came from (`GalleryDrawing.Fit` measured the
  figure's labels at the previous zoom: shrunk and grown back, the figure came back at half
  its size). Checked by laying `MainView` out at a grid of sizes from 0 to 800 px on every
  page and side-panel state, with every first-chance exception logged (a temporary
  in-process probe, as for undo); Windows and headless Edge can't make windows that small.
- **A visual invalidated during the render pass freezes the window**: `InvalidateVisual`
  from inside a `Render` (also through something it calls that invalidates, such as a
  presenter making its text layout) throws "Visual was invalidated during the render pass"
  in the compositing renderer, and the renderer stops for good. The window keeps its last
  frame, ignores nothing but shows nothing new, every thread idle; no error report is seen,
  since it is never drawn. Invalidate from where the state changes, not where it is read
  (`ColoredTextPresenter.InvalidateTextLayout`).
- **`Shape.Render` is sealed** (Avalonia 12): a Shape can't draw anything but its geometry.
  `PointMarker` draws a character through an `EmojiGlyph` visual child it measures and arranges
  itself; with no geometry, `Shape.ArrangeOverride` returns size 0 and the shape collapses to
  the middle of its place, so the override returns the final size then.
- **A `FontFamily` in a collection that isn't registered yet** ("fonts:Emoji#...") makes text
  layout throw *while rendering*, which ends the desktop app. `EmojiFont.Family` is the default
  family until the font has loaded.
- **`RotateTransform`**: no CenterX/CenterY - Avalonia rotates about `RenderTransformOrigin`
  (the middle by default), and WPF-style centering shifts a tilted ellipse off its center.
- **A TextBlock drops lines that don't fit its arranged height**: it lays its text out again
  at arrange time, in the arranged size, and leaves out every line past its height (they are
  not clipped: dragging the label doesn't bring them back). On an iPhone only (3x pixels;
  never seen on Windows or in headless Edge) a wrapped caption lost its last line, also when
  laid out at least as high as it measured, so the arranged layout wrapped into a line more.
  A label's text is a `LabelTextBlock`, laid out again at the width it measured at and with
  no height limit; that fixed it (2026-10-03). The exact cause wasn't found. Its arrange
  keeps the measured layout, and Avalonia 12's TextBlock keeps its layout across a text
  change unless its measure was valid (`OnMeasureInvalidated` drops it; only its own
  arrange rebuilt it every time), so `LabelTextBlock.MeasureOverride` drops it first: an
  angle's number dragged through length 0 and back showed "0°" for "120°" (2026-10-14).
- **In the browser a posted job can keep the page from painting**: Avalonia's dispatcher
  runs what is due (Background jobs, a short timer) in the same JavaScript task as the
  animation frame, before it gives the page back, and the browser paints only between
  tasks. The gallery's first batch of tiles, posted by the first layout, kept the splash
  (which Avalonia closes at the first frame) up a second longer. To let a frame be seen
  first, start the work at the frame after it (`TopLevel.RequestAnimationFrame` twice:
  `GalleryView.OnAttachedToVisualTree`).
- **TextChanged arrives late**: Avalonia raises a TextBox's TextChanged through the dispatcher,
  after a programmatic-set guard is gone. Text editors of the property grid remember the text
  they put in the box (`StringEditor.ShownText`) and ignore a TextChanged carrying it;
  otherwise showing a figure records a property set per text row. Only until the user's first
  edit: after that the same text is theirs (typed and deleted back to empty must reach the
  property). A commit (Enter, leaving the box) applies the text itself rather than trust the
  pending TextChanged, since a panel's Enter handler (Add point) reads the property next.
- **Stale Bounds after a load**: zoom to fit measures labels itself (`Measure`) because their
  Bounds are stale right after a load. `Label.MeasureSize` invalidates the Border before
  measuring: its measure stays valid when only the TextBlock inside changed, and answers with
  the old size.
- **The Fluent theme paints hover/selected states on the template's ContentPresenter**, so
  style overrides must target that part (`PropertyGrid/PropertyGridTheme.cs` does).
- **The library's own types shadow framework ones**: `Math`, `Ellipse`, `Polygon`, `Path`...
  In a file-scoped-namespace file a `using X = ...;` alias does NOT win over a type of the
  enclosing namespace - write `System.Math.Max`, `Avalonia.Controls.Shapes.Ellipse` in full.
- **`Style` and `Setter` are ambiguous in the library**: `LiveGeometry.Style`/`Setter` are the
  WPF shims and shadow Avalonia's. For real Avalonia styles alias them
  (`using AvaloniaStyle = Avalonia.Styling.Style;`).
- **Browser has no system fonts.** Text renders only because `Avalonia.Fonts.Inter` is embedded
  (`.WithInterFont()`). Inter reaches text by *inheritance* from the window (the theme sets it
  there); an explicit family that doesn't exist - and `FontFamily.Default` and
  `FontManager.DefaultFontFamily` too - renders as Noto Mono in the browser. So `TextStyle` sets
  the font a drawing names (Arial, Segoe UI) only when it really resolves, and naming "Inter"
  in a drawing doesn't work (embedded, not installed).
- **The app runs under the invariant culture** (`App.UseInvariantCulture`, first thing in
  both `Main`s; the Browser project also builds with `InvariantGlobalization`, so no ICU
  data is downloaded). Numbers use a point everywhere: the expression language, `.lgf` and
  `.dgf` files, Avalonia path markup, labels. Anything that formats or parses numbers for a
  parser still says `InvariantCulture` explicitly; a `$"M{x},13.5"` under a decimal-comma
  culture makes `Geometry.Parse` throw. To test a culture: `webauto stop`, then
  `webauto start <url> 1280 800 --lang de-DE`.
- **Files in the browser** (`MainView.SaveDrawingAs`/`OpenDrawingFromFile`): the File
  System Access API takes a file type only as a MIME type with its extensions, and Avalonia
  drops a `FilePickerFileType` without `MimeTypes`. The stream from `OpenWriteAsync` has only
  `WriteAsync`; a `StreamWriter` flushes synchronously on dispose and throws, so the text is
  written as bytes with `WriteAsync`/`FlushAsync`. To test saving headless (no native dialog):
  `webauto eval` a fake `globalThis.showSaveFilePicker` returning a handle with `kind`, `name`,
  `getFile`, `queryPermission`, `requestPermission`, `createWritable`
  (`write`/`close`/`seek`/`truncate`) *before the first picker use* - Avalonia's picker
  polyfill captures the global when its storage module is first imported, later replacements
  are ignored - then click Save and read back the bytes; throw a
  `DOMException(..., "AbortError")` for cancel. Opening headless works the same way: copy the
  file into the served `wwwroot`, `eval` a `globalThis.showOpenFilePicker` returning one
  handle whose `getFile` gives a `File` built from `fetch('/name.ggb')`, then `key O ctrl`.
- **A press moves the keyboard before its PointerPressed is raised**: Avalonia focuses the
  nearest focusable control under the pointer (the drawing control, the Figure List) as
  the press comes in, so a text box's LostFocus fires before the canvas handlers see the
  press. The expression box keeps its role as the drawing's insert target until the press
  has been handled (a posted clear, `ExpressionTextBox`) and takes the keyboard back after
  it put a reference in; cleared on LostFocus at once, every click on a figure went to the
  tool (2026-10-15).
- **Keyboard focus drifts into tool panels.** A tool's PropertyBag panel (e.g. "Point by
  coordinates") takes focus into its TextBox after every construction step, so neither the canvas
  KeyDown nor `MainView_KeyUp` (which skips TextBox focus) sees keys then. Anything that must
  always work (Escape) belongs in the `MainView_KeyDown` tunnel handler. When the focused
  control leaves the tree (a panel rebuilt around its button), Avalonia focuses nothing and
  keys go to the window alone, past MainView; `MainView.TopLevel_KeyDown` catches that,
  refocuses the canvas and forwards the key. Ctrl shortcuts are
  handled on key *down* (`MainView.HandleControlShortcut`): on key up Ctrl may already be
  released and a bare S is the Segment tool. In a text box Ctrl+C, V, X, Z, Y, A are the
  box's own; S, O, N and F1 are not (they did nothing there, and the browser took them),
  and the canvas takes the keyboard first so that the typed text is committed. The letter
  of a Ctrl shortcut is skipped on its way up (`shortcutKeyDown`) - until it is pressed
  again on its own, since after Ctrl+S or Ctrl+O its release goes to the file dialog.
  Cmd does what Ctrl does (`MainView.IsCommandModifier`, `Behavior.IsCtrlPressed`): in the
  browser on a Mac every shortcut was dead, Cmd+S saved the web page, and Ctrl+click there
  is a right click. Ctrl+Shift+Z redoes as Ctrl+Y does, and Backspace (a Mac's "delete"
  key) deletes the selection as Delete does. Text that names a key names it as the
  keyboard does (`KeyNames.Alt`: Option on a Mac); `KeyNames.IsMac` is the desktop's OS,
  and in the browser the page's platform (`isMac` in `main.js`), since a browser says it
  runs on "browser". The browser sets it before Avalonia is up, which is why it is a class
  of its own: touching `Behavior` that early ran its static constructor, which makes
  cursors, and the page died with "Unable to locate ICursorFactory".
  Plain keys: `MainView.HandlePlainKey`; tool letters: `UI/BehaviorShortcuts.cs` (also feeds
  the tooltips). The keys that move the view (arrows, Home, Page Up/Down, +/-) do nothing
  while the side panel has the keyboard: its lists, sliders and combos use them, and an
  arrow that picked the next style also panned the canvas.
- **Every exception is shown** (`MainView.CurrentDomain_FirstChanceException`): status bar plus
  an `ExceptionReport` page in the side panel, also printed to the console. Add to `IsBenign`
  when a framework exception turns out to be noise (the browser's file picker throws
  `JSException` "AbortError..." on cancel; `GeoGebraReader.LeftOutException` is how that
  reader leaves out an element it can't read, and says so in the status). So never use a
  caught exception as a test (decode
  as UTF-8 and catch, as `DecodeLegacyText` once did): check instead. `--check` runs print
  every one, and should print none. One that escapes an event handler or a posted job is
  then swallowed (`Dispatcher.UIThread.UnhandledException` in `App.Initialize`), so a bug
  shows as an error report and the app keeps going. Not tried on the layout exceptions of
  "A point with an infinite coordinate", which may come back on every layout pass: still
  never hand layout NaN or infinity. A `JSException` has no .NET stack, so the handler
  captures `Environment.StackTrace` at the throw. Text boxes of the property grid wrap at
  480 px (`PropertyGridTheme`), so a stack trace doesn't stretch the panel across the window.
  What is shown is also sent to the site (2026-10-15; `ExceptionReports.cs`), desktop and
  browser alike, in a build given `ShareService` (see "Share links"):
  `POST {ShareService}/api/v1/exceptions`, `Content-Type: text/plain` with
  `{ build, platform, reports: [ { type, message, frames } ] }` (204; 400, 403, 413, 429 as
  for shares), where `frames` is the stack at the throw as `Namespace.Type.Method`, captured
  in the handler (the exception's own trace has only the throwing frame at first chance) and
  cut after `CurrentDomain_FirstChanceException` and the runtime's `System.Runtime.EH`
  frames, and the message is cut at 1000 characters. In the browser build (compiled ahead of
  time) that walk finds no frames at all: the report keeps the exception and reads its own
  trace when it is sent, by when it has unwound - the throw site out to where it was caught
  (`TraceFrames`, which reads both runtimes' text). There the message is .NET's resource
  key and its arguments ("Xml_MessageWithErrorPosition, Xml_UnexpectedEOF, 1, 33"): the
  trimmed build leaves the framework's message texts out. Each goes once among the last 10 (a loop sends one), at most 50 a session,
  queued and sent by a background task in batches of up to 20, 5 seconds after the first of
  a burst. Nothing of the reporting is reported (`ExceptionReports.IsOwn`, checked first in
  the handler, so it isn't shown either): an exception while a report is made, anything the
  sender's code throws (an `AsyncLocal` flag flows into its awaits), and a network failure
  while a batch is on its way (`IsNetworkFailure`, which sharing uses too). A build without
  `ShareService` sends nothing.
