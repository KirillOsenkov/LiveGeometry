# macOS

The desktop head builds and runs on a Mac (Apple silicon, .NET 10 SDK, no workloads) from the
same project; the regression suite, `--check` and `--rewrite` work there too (the gallery
rewrites byte for byte). What is the Mac's own (2026-10-02):
- **Ctrl+click is a right click** (`Behavior.IsSecondaryClick`): AppKit hands it over as a left
  press with Control, and Avalonia passes it on as one. Its moves and release are swallowed.
  Ctrl's other uses are Cmd's there, as in the browser on a Mac.
- **The menu bar**: `App.Initialize` sets the application's `Name` (it said "Avalonia
  Application") and its own `NativeMenu` ("Live Geometry on GitHub"; Avalonia's default has
  "About Avalonia"); Avalonia adds Services, Hide and Quit after it. Cmd+Q goes through the
  window's Closing, so settings and the window bounds are saved.
- **The Dock icon** (`LiveGeometry.Desktop/MacDockIcon.cs`): Avalonia ignores `Window.Icon`
  on a Mac and a bare executable shows a blank document, so the vector `AppIcon` is rendered
  and given to `NSApplication.applicationIconImage` through the Objective-C runtime.
- **Window bounds** (`WindowBoundsPersistence`, the `WindowBounds` setting): Avalonia's own
  position and client size, those of the last moment the window was normal. A zoom animates
  through a few sizes while the window still says Normal, so bounds count once they have held
  for a second. The green button (Option+click: Zoom) maximizes, but in Avalonia 12.1 it
  doesn't un-zoom a maximized window, also one zoomed by hand: drag its edge instead.
- `.lgf` files are written with CRLF on every system (`DrawingSerializer.XmlSettings`), as
  the gallery's are and the batch rewrite does.
- Testing traps: Cmd+A in the save panel's name box selects the name without its extension,
  so typing `x.lgf` gives `x.lgf.lgf` (the panel, not the app). `winauto.cs` and
  `contactsheet.cs` are Windows only (WinForms); `macauto.cs` is the Mac's (below).
