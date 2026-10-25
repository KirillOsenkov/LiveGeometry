# The side panel and context menus

The property grid (rows, editors, verbs, typed text), where the side panel goes, tool panels, context menus.

- **Numbers are shown with `Settings.DisplayDecimals`** (2) by every editor of the property grid
  and by labels by default; values keep their digits and a typed number is taken as typed.
  Rounding for display goes through `Math.Round(value, digits)` of the library, which adds
  0.0: -0.001 rounds to a negative zero, and that is written "-0" (a point at (3, -0)).
  A half rounds up, as at school (3.125 is 3.13): through `decimal`, which takes 2.675 as
  written where the double is a hair under, and away from zero; `System.Math.Round` goes
  to the even digit, and an area of 3.125 said 3.12. A point's coordinates are written
  (3, 4), with a comma.
- **Several figures in the grid** (`FigureSelection`): a row whose values differ has no
  value (`CompositeValueProvider` answers null), and every editor has to show that as
  nothing - an empty box, no item selected (`SelectorValueEditor.ShowSelected`), a check
  box in its third state (a click then ticks them all) - and not as a value none of them
  has (0, the first item, unticked). The enum editor threw on the null: select all in a
  drawing with two kinds of measurement units. Up/down buttons do nothing without a value
  to step from. What defines a figure is no row of a selection (decided 2026-10-12): the
  X and Y of points by coordinates and the parts of a circle's equation (expression rows,
  left out by type in `CompositePropertyProvider`, since the row is the expression and
  carries no attributes), a function, a label's text, a line's equation
  (`[PropertyGridDisallowMultiEdit]`, as Name has). One text made every selected figure
  the same, and each needs its own checks; the X row threw, the function row saved
  "sin(" into both graphs. How figures look, Visible, Locked, Z and Delete stay.
- **The order of a page is declared** (2026-10-18; `PropertyGridSectionAttribute`): every row
  and button is in a section - Header, Definition, Display, Style, Verbs, Arrange, Delete,
  the order of the context menus - at a place in it (smaller first, in tens so a class can
  slot a row between its base's; ties in declaration order). Unmarked, a row is Definition,
  a button Verbs, a destructive one Delete; a group box goes where its first member goes.
  So the shared rows are marked once in the base classes (`FigureBase`: Name, Visible,
  Locked, Style, Bring to front...) and every figure's own rows come first. The Header
  section is not on the page but on the title's line: what the figure is
  (`FigureBase.KindCaption`, faded), the name as a box as wide as the name, at least 60
  (`HeaderNameEditor`: no box until hovered or focused, A₁ shown until edited as A_1), and
  on/off rows as icon toggles right after it (`HeaderToggle`: an eye for Visible, a
  padlock for Locked; the usual state faint).
  A page without such rows (a style, a tool panel) keeps the plain title.
- **What can be measured of a figure is on its page** (2026-10-18; `FigureQuantities`, shown
  by `QuantitiesEditor` after the Definition rows): the lengths the Length tool takes of it
  (`LengthMeasurement.QuantitiesOf`: a circle's radius, diameter, circumference; a polygon's
  perimeter; an arc's length), its area where the Area tool takes it, the central angle of
  an arc, sector or segment, a straight figure's slope, and its angle where no Direction row
  says it. Read-only and selectable (double click takes the whole number, Ctrl+C copies it:
  `MainView` leaves Ctrl+C to a `SelectableTextBlock` with a selection; the row's context
  menu copies it too). A "+" puts the measurement on the drawing, one undo step, shown only
  while the drawing has none of that kind for the figure (for a segment's length, a distance
  between its ends counts). A quantity an editable row says (a segment's Length, a circle's
  radius: `RowName`) has no row of its own; its "+" goes after that row's box
  (`LabeledValueEditor.AddTrailing`). A slope and a line's angle have no measurement figure,
  so no "+".
- **Property grid layout is declarative** (`PropertyGrid/`): `[PropertyGridGroup]` boxes rows
  and their buttons together, `[PropertyGridDestructive]` puts a button last under a divider,
  `[PropertyGridIcon]` puts a drawn icon (`PropertyGridIcons`) in front of a caption (every verb
  button has one - give a new verb one too), `[PropertyGridPreferredEditor("UpDown")]` picks the
  editor, `[Domain(min, max)]` on a double gives a `SliderEditor`. String boxes are one line
  (Enter is the panel's: Plot, Add point) unless the property says `[PropertyGridMultiline]`
  (label text). Rows that are editable only
  sometimes: `[PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]` on the
  property and `IConditionalProperties` on the figure (`CanEdit`, `Caption`); the same interface
  vetoes buttons by method name. Rows that show only sometimes: `[PropertyGridCondition(name)]`
  on the property (as on a verb, a bool method of the object), asked again on every
  `PropertyChanged` with a name - the row is hidden, not taken away, so a box being typed in
  keeps the keyboard (a label's Decimals show only while its text has a `[...]` part,
  `Label.ShowsDecimals`, and come as the first one is typed; the Text setter raises the
  change). Where rows and buttons go: see "The order of a page is declared". The grid refreshes a row's value on `PropertyChanged` with
  its name, and rebuilds itself (posted) on a `PropertyChanged` with a null name: for a set
  after which the rows change shape - what is editable, a caption, which buttons (a
  translated point's Free direction). A property setter that changes the figure list (the length
  panel's Show, the point's Free toggles) does so directly: the grid records the property set as
  the undo step and the undo library refuses an action recorded from inside another. A tool's
  panel holds settings, not drawing state, so it says `[PropertyGridNoUndo]` (`ToolPanel` does,
  for all that derive from it; a panel that doesn't derive says it itself): otherwise typing in
  it is an undo step that undoes nothing visible. `[PropertyGridFocus]`
  (the editor takes the keyboard whenever it appears) is for tool panels only; a figure's
  row that should take it just once, when the figure is created (a new label's text), is
  named by `RaiseDisplayProperties(figure, focusProperty:)`, otherwise selecting the figure
  anywhere (canvas, Figure List) steals the keys. A tool panel with a focused row appears
  only at the step that asks for it (Rotate, Dilate, Translate: `PropertyBag` is null
  otherwise - test `Transaction != null` too, since "construction complete" re-reads it
  before the found figures are cleared); one shown all along holding a setting (Line at
  Angle) doesn't take focus.
- **The side panel** (`DrawingHost.CreatePropertyGrid`) is one rounded surface: a fixed
  header row (the grid's title, put there through `PropertyGrid.HeaderHost`, and the ×)
  over a scroll viewer with the rows. Nested grids (complex types, method parameters) have
  no `HeaderHost` and keep their title as the first row. A page starts at its top
  (the scroll offset is reset whenever the grid is given something to show), and a list in it that brings
  its selected item into view (the styles, the emoji) scrolls only itself: the request
  is stopped on its way up to the panel's scroll viewer (`CreatePropertyGrid`), or in a
  small window the panel opened scrolled down to the figure's style, its first rows cut
  off. After undo or redo (`DrawingControl.UndoneOrRedone`, `DrawingHost`) a panel that was
  up shows a figure's page as the selection is now, read again, and a tool's own panel
  again; anything else goes away (the panel after a construction, a style, the settings:
  they may be about a figure that left), and a panel put away stays away. Every panel went
  away until 2026-10-12, a selected figure's too.
  **Where it goes** (`DrawingHost.SidePanel.cs`, 2026-10-17): it must not cover what it is
  about. In a wide window it is at the right, or at the left: the button by the ×
  (`MovePanelToOtherSide`, an arrow to the edge it would go to) moves it and makes that its
  side (`PanelAtLeft`, kept between runs by the app: `SidePanel=Left` in the settings
  store, `MainView.SidePanelKey`). Where the drawing would keep less than
  `MinimumRoomBeside` beside it (a phone) it goes across, at most half the height, the rows
  wrapping and scrolling, at the top, and no button. A figure the user just picked
  (`revealSelection`: a selection changed, not undo's page shown again) that is under the
  panel is brought out by moving the view, gliding (`RevealInPlane`): the least move that
  shows it whole, `RevealGap` from the panel, clear of text pinned to the screen. Not an
  undo step (the app's move, as a fit after a resize; the history keeps a view's moves as
  offsets, so one left out upsets nothing), and the view goes back when the panel goes,
  unless something moved it since. What the view can't move - a caption, any label pinned to
  the screen - or what has no room to go to, gets the panel on the other side for that page
  (on a phone the bottom half); right after the button the panel stays where the user
  put it. The panel is placed once the page has settled (posted at Background, unseen
  until then: the value plates and style swatches fill in after the first layout, and the
  rows' shared caption column settles on a second pass - measured at once, the panel came
  out 50 px narrower than it is), with a real layout pass (`UpdateLayout`), so the grid
  raises `VisibilityChanged` after its rows are built. Tried and left (2026-10-17, by
  screenshots): sliding a caption out from under the panel (its plate then covered the
  figure, and the text being edited moved), and always taking the side that covers less
  (the panel jumped from side to side between clicks).
- **Context menus** (`UI/ContextMenus.cs`, 2026-10-18; the Drag tool's right click,
  `Dragger.MouseRightClick`): fixed sections in a fixed order, a separator only between two
  that have items (`ContextMenus.Section`). A figure: Choose figure (where figures overlap),
  Show (check items), Verbs, Arrange (Bring to front, Send to back), State (Hide, Lock),
  Edit (Copy, then Delete, always last). The paper: Zoom to fit, Axes, Grid lines; Paste,
  Select all; Background…. Show and Verbs are the property grid's own rows and buttons
  marked `[ContextMenu]`, found and vetoed as the grid finds them
  (`PropertyGrid.GetCallableMethods`, the discovery strategy, `IConditionalProperties`,
  `[PropertyGridCondition]`) with the grid's captions and icons, so a verb never needs
  code in the menu; "Convert to …" verbs go into one submenu when there are two or more.
  Several figures selected: `FigureSelection`'s rows (the ones all of them have) and verbs.
  A name label's menu is its figure's Show items alone. Only a menu can offer: Snap to among
  several figures, Edit text. What is marked is what is handy at the cursor (Show name,
  Free point, Fix length, Horizontal, Convert to, Detach...); style, values and typing stay
  in the grid.
- **A figure inherits the rows and buttons of its base class**, and has to take away the
  ones that are not its own: a button through `IConditionalProperties.CanEdit` (by method
  name), a row by overriding the property with `[PropertyGridVisible(false)]`, or by making
  it read-only (`ConditionalPropertyValue`). Each of these was a bug: Convert to ray /
  segment on a parallel, a perpendicular or a bisector (built on a line and a point: it
  threw; `LineTwoPoints.IsThroughTwoPoints`), Convert to segment / sector on an angle's
  arc (`AngleArc`), the Text box of a measurement (`Measurement`, whose text is
  worked out), the style rows and buttons, Visible and Locked of a Number (nothing on
  the paper; Select all leaves numbers out, or a selection with one would lose those
  rows for every figure), a vector's Direction and an
  angle's Arcs when nothing can change them, "Delete this style" on a default style. A
  row or button that does nothing is an undo step that undoes nothing. A row whose setter
  only stores (a point on a figure's `Parameter`, which a locus samples through) gets a
  property of its own for the grid that also moves things (`ParameterDisplay`).
- **A Convert button is not offered when the new kind lacks what something takes from the
  figure**: a segment's length (`IFigureExtensions.IsUsedForLength`: a length measurement
  of it, a circle with it for a radius, a translation or dilation by it) for Convert to
  line / ray, a sector's or circular segment's area (`IsUsedForArea`) for Convert to arc,
  and a number an expression reads off it that the new kind has no property for
  (`IsReadByExpressionsBeyond`: [AB.Length] for Convert to ray; read off the parsed
  expressions, so that `` `base side`.Length `` counts - a regular expression over the text
  missed it once the segment was renamed so). An expression that no
  longer compiles drops what it was compiled to (`DrawingExpression.Recalculate`), which
  held the figure that left: the point at (AB.Length, -1) stayed at its place for the
  session and was nowhere once the file was read again.
  `ReplaceWithNew` hands every dependent over to the new figure whatever it is: a distance
  measurement of a segment that became a ray threw on every redraw (it now doesn't exist
  when it has nothing to measure). Convert to polyline (not offered since 2026-10-01: no
  tool makes polylines; a method is a button when it has `[PropertyGridVisible]` at all,
  `(false)` included, so it has none) deletes the polygon, and an area
  measurement of it with it.
- **Closing the side panel** (its ×, or a press on empty chrome: the ribbon's empty
  strip, the toolbar beside its buttons, the Figure List below its rows) goes through
  `DrawingHost.CloseSidePanel`. A tool's own panel (same type as the tool's `PropertyBag`)
  can be the only way on with that tool, so closing it puts the tool down (Drag, as
  Escape); anything else is only hidden. "Empty" is decided in `MainView.IsEmptyChrome`:
  from the hit visual up to the ribbon/toolbar/list through plain layout types only (exact
  types: a toolbar button is a `Border` subclass).
- **The button that accepts a panel says OK**, unless a verb says more (Plot, Add point,
  Close figure, Create tool), and **no panel has a Cancel**: the × of the side panel is
  that (`DrawingHost.CloseSidePanel` puts the tool down when the panel is the tool's own,
  which goes by `PropertyBag` - so a tool with a panel per step, like Define figure, must
  answer with the panel of the step at hand).
- **Enter in a tool panel's box is the panel's button** (OK, Plot, Add point). Each panel
  wires that itself: `[PropertyGridEvent("KeyDown", ...)]` on the property and a handler
  that calls the method. Nothing does it for a new panel, which is how Rotate and Dilate
  went without. The editor commits the text first (its own handler is on the text box, the
  panel's on the editor around it). A number box that says "Type a number." has kept the old
  value, so the handler looks at `StringEditor.ErrorText` before it goes on. Left alone on
  purpose: the length panel, where Enter applies the length and OK only closes. A press
  on the button of the tool that is on already (opening a tab picks its tool) takes the
  keyboard from the panel's box; `BehaviorToolButton.Click` shows the panel again, which
  gives it back. `winauto keys` swallows parentheses: test with `x*x`, not `sin(x)`.
- **Typed text that is wrong is said under its box**, never in the status bar
  (`StringEditor.ErrorText`: a pink plate attached to the box, widening the row up to the
  480 px a box may take). Only on commit - Enter or leaving the box - never while typing, and
  it goes when the text is edited. Two sources: an editor's own `Validate` (`ExpressionEditor`,
  `FunctionEditor` on `[PropertyGridPreferredEditor("Function")]`, `DoubleEditor`), and a tool
  panel's command: panels derive from `ToolPanel` and check each row with `Compile`/`Evaluate`
  (`ReportError` for anything else), which the grid routes to that row's editor. An empty box
  compiles to nothing without an error: `CompileResult.GetErrorText` supplies one. A whole
  number (`IntEditor`) is checked against the property's `[Domain]` ("Type a whole number
  from 3 to 500."). The two number boxes without an error plate put the box right on commit
  instead: `SliderEditor` takes a number beyond its range to the nearest end, `UpDownEditor`
  shows the value again when the text is no number. No editor takes "Infinity" or 1e999.
  **What an expression is worth is said under its box too** (2026-10-15;
  `StringEditor.ValueText`, a plate in the box's own colors): "= 18.85" while the text
  compiles, as VB6's calculator said it, for the rows of a point by coordinates, a
  number's formula and the expression rows of the tool panels; the error plate takes its
  place. Every box an expression is typed into is an `ExpressionTextBox` (see "Building a
  formula by clicking").
- **A commit sets the typed text once.** Enter commits, and leaving the box commits again:
  the editors remember the text that has gone into the property (`appliedText` in
  `StringEditor` and `UpDownEditor`) rather than compare it with the value read back, which
  need not equal it (a vector's direction and a segment's length are worked out from
  points: 27.57 comes back as 27.570000000000004) - each further set was an undo step that
  undid nothing. And the text an editor put in the box itself (`shownText`, kept so that
  its own late TextChanged is not taken for typing) is forgotten at the user's first edit
  in the number boxes too (`UpDownEditor`, `SliderEditor`, as in `StringEditor`): a box
  that said 12, Backspace (the length is 1), the 2 typed back - the 12 was taken for the
  editor's own text and dropped, and the segment stayed 1 long under a box saying 12.
