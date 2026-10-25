# Backlog

What is still missing, and what has not been tried in the browser.

## VB6 parity backlog

Still missing compared to VB6, roughly by value: unsaved-changes prompt; live cursor
coordinates in the status bar; undo/redo
captions naming the action; recent files; print; message, sound and launch buttons (the `.dgf`
reader leaves button types 1-3 out); "Create locus" on a point (the Locus tool does the
tracing); measurement label dragging constraints; rulers; "tool select once" option (every
construction returns to Drag); languages (en/ru/uk/de); step-by-step construction playback;
Calculator. Not a gap but a choice: double-click
zooms to fit, where VB6 opened properties (here selecting a figure shows them).

Done since the list was made: the point symmetric about a point and the inverted point are
Reflect with a point or a circle for the mirror; point shapes and sizes, name colors and dashes
are per style; show/hide buttons are `ShowHideControl`, made with the Show/hide box tool
(2026-10-07; VB6 dragged its buttons with the right button, here a drag is a drag and a
click ticks); settings persist (`SettingsStore`);
the axes are lines to build on (`AxisLine`); a click among overlapping figures is chosen
with Tab, a tap's in a menu, a selection in the context menu (`ClickChoice`).

## Not yet verified in the browser

The `Hyperlink` figure (only in old files, no tool makes one),
which fetches the drawing at its URL with `HttpClient` when clicked - not when read, which
put a failure's whole stack trace into its text, saved with it.
Saving through the browser storage provider works (drawings, and PNG and SVG pictures);
opening has only been checked as far as the picker (see "Files in the browser"). Copy image
finishes without an error in headless Edge, but reading the clipboard back is denied there:
what a paste gives was only checked on the desktop.
