# Undo

- **Undo** (GuiLabs.Undo; `Actions/`). What must hold: every change to what a file saves is
  one undo step, undo puts back exactly what was there, redo what was made. The traps, each
  of which was a bug (2026-09-30):
  - *A grid button calls its method directly* (`MethodCallerButton`): the method records
    itself (`Actions.*`, a `CallMethodAction`, a transaction when it is several actions -
    Create new style), or it is a hole.
  - *`SetPropertyAction`* keeps an old value per object of a multiple selection (they may
    have differed). A value whose getter resolves - the theme's paper for a drawing without
    one, the base value of a style without an override - is an `IRestorableValue`
    (`ThemedValue`, `Drawing.PaperValue`): undo of the first set takes the override or the
    paper away again instead of writing the resolved value.
  - *Merging* (a run of sets as one step) is asked for by the caller (`coalesce:`): editors
    of typed or dragged values do (`LabeledValueEditor.CoalescesEdits`), a check box or a
    choice doesn't (merged, hide-then-show was a step that undid nothing). Same object,
    same property, same theme only; and never while there is something to redo: the
    library merges without ending the redo chain, so `SetPropertyAction` and `MoveAction`
    refuse then. And only within one run of edits (`SetPropertyAction.Run`, a token of the
    editor): Enter, leaving the box and the grid showing the object anew each start
    another (`EndEditRun`), or a length typed now joined the one typed five minutes ago.
    Undo of a set on several objects restores them last to first.
  - *`MoveAction`* moves by an offset once; undo and redo restore places
    (`IRestorablePlace`: coordinates, a label's pixel offset or pin offset, a translated
    point's distance and direction). Moving back by the offset left a point on a circle, a
    slider knob behind its stop and a label after a zoom somewhere else. The coordinate
    system has no place, only offsets. Keyboard pans use one list of movables so that a
    run of arrow keys merges like a drag.
  - *A setter that adds or removes a figure* (a point's label, a line's name, the length
    panel's measurement) brings the same object back, never a new one: the history may
    hold a drag of it. `PointBase` keeps its label (`keptLabel`), also across undo and redo
    of the point itself; a point is auto-labeled only the first time it enters a drawing.
    And in the same place of the figure list (`RootFigureList.Retire`, then `Return` or
    `ReturnBefore`): the list's order is the file's.
  - *A check box that stands for more than yes and no* (a translated point's Free
    distance: free, a Number, or tied to a vector) is an `IRestorableValue` too
    (`TranslatedPoint.FreedomValue`): undo puts the source back, where unticking the box
    would give it a Number - and a Number's value with it: freed, the Number is kept aside,
    and unticking the box after a drag writes the new value into that same Number. A
    label's pin too (`Label.PinValue`): pinned, the label goes
    where the screen carries it, so undo of the pin puts it back where it was in the
    plane, not just the pin (pin, zoom, undo left it somewhere else).
  - *Replacing a figure by another kind* (`Actions.ReplaceWithNew`: Convert to line, ray,
    segment, arc, sector; Reverse) is a transaction that is not delayed, so that each
    step sees the drawing as the one before left it: the new figure goes to the old one's
    place in the list (`MoveBefore`), takes over its name label, stays hidden or locked
    if the old one was, takes its selection (outside the transaction, as a paste's: the
    verb was clicked in its grid, and the halo stayed on the figure that had left), and
    is named AB by its place in the list (`SettleDefaultNames`).
    While both are in the drawing they trade the names AB and AB2, so expressions that
    name the figure (`[AB.Length]`) sit the replacement out and are compiled again at the
    end, undo included, as in `ReplacePoint` (`SuppressRenameInExpressions`, then
    `RebindExpressions`): following the renames, the text ended as `AB2.Length`, the name
    of the figure that had just left.
    Convert to polyline keeps the polygon's vertices (ABCA).
  - *Rows of a nested grid* (`ComplexTypeEditor`: a line's Equation) get the action
    manager from their parent editor, which passes it on when it is set: an editor that
    starts expanded makes its rows before that, and m and b were edited past undo.
  - *An undo or redo that throws* leaves the library's "currently running" mark on, and
    it then refuses every action: `DrawingControl.ReleaseStuckAction` takes it off.
  - *A click is not a drag* until the cursor has gone `Dragger.DragThreshold` pixels:
    a press with a wobble moved the point or the view by a pixel, selected nothing and
    left an undo step that seemed to undo nothing.
  - *While a transaction is open* (`Drawing.IsRecordingTransaction`: a construction or a
    point drag under way) anything recorded joins it and is rolled back with it. So Redo,
    Delete and Paste do nothing then, nor does a click on a show/hide box (shown without being
    recorded, its figures stayed hidden after Escape and undo), a keyboard pan and the grid
    toggle are not recorded, and a click in the Figure List abandons the construction first.
  - *Loading is not undoable*: `DrawingControl.ForgetLoading` clears the history in a
    `finally`, also when the file threw half way. A construction left half way goes with
    its drawing (`DrawingControl.Drawing` resets `ConstructionInProgress`): still "in
    progress" for the next drawing, Undo only restarted the tool and Redo did nothing.
  - *Paste* (`Actions.Paste`) is one step: the copies, then a move of their roots
    `Drawing.PasteStep` pixels down and right per paste of the same copy, and they are
    selected. On top of the originals and not selected, a paste looked like nothing, and
    each try left another copy under them. Copy with nothing selected leaves the clipboard
    as it was. The copies are read under new names where the old ones are taken, so
    whatever a figure's `ReadXml` finds by name must look among its own `Dependency`
    elements, not the names its dependencies have now (a translated point's sources: the
    copy of a fixed segment's end came back free, on its pivot). Expressions of the copies
    are compiled as they are read, when the names they say are the originals'; once the
    copies are in, their texts are rewritten to the copies' names and compiled again
    (`PasteAction.RebindExpressions`, an `ExpressionRenamer` that looks among the copies
    first): the copy of a label `[AB]` measured the original segment. The clipboard
    carries the styles of their own the copies name (`DrawingSerializer.WriteFiguresWithStyles`):
    pasted into another drawing, a figure looked its style up there by name, and took the
    default look, or another drawing's style "1" (a text style on a point). The paste brings
    a style the drawing lacks, under a free name where the name is taken by another look
    (`StyleManager.FreeName`), and takes it away again on undo. The default styles the
    copies are on come along too, named by them (`StyleManager.NamesDefaults`,
    `GoesWithoutSaying`): a file leaves a figure's default style unsaid, and a copy took
    the other drawing's default, changed in either drawing (Edit this style); one the
    same in both is taken as it is. A tool being defined carries only the defaults its
    drawing changed (`DefaultStyleNames.Changed`): kept between runs, a copy of an
    unchanged one would come out as a second style once the theme's colors changed. The new names are
    decided for all the styles first, none of them the name of another style coming in,
    and the copies renamed once (red "1" became "12", the blue style's name, and renaming
    blue to "122" took the red copies along). Text that is no figures
    (plain text, a page) is said so in the status, not read as XML (that threw).
  - *Parts of a composite* that are among a figure's dependents (a Bezier path's handles)
    are not removed or put back by `RemoveFigureAction`: they go with their composite.
    Parts removed by a change of the composite are retained for reuse: redo may restore a
    later construction that still holds those exact objects.
  - *Length edits* restore the endpoint's place and any Number holding its distance
    (`LengthPropertyValue`, also used by the length panel), not just the previous length:
    zero loses the direction and a negative length reverses it. A point on a figure restores
    its exact parameter, not a new projection of its coordinates.
  To check a change: a temporary `DispatcherTimer` in `MainView` that appends the undo and
  redo counts and a hash of `Drawing.SaveAsText()` to a file whenever they change, driven
  with `winauto`. A hash that changes without a new step is a hole; one that doesn't come
  back on Ctrl+Z is a wrong undo. What found the most (2026-09-30, second pass) was a
  sweep in the same temporary code: for each figure of a drawing, show it in the real
  grid, and for each row set another value through the row's own value provider and
  action manager, for each button invoke it; after each, compare the saved text before,
  after, after undo, after redo, load the saved text into a second drawing and save that
  (a round trip), run `Figures.CheckConsistency`, and log every first-chance exception
  with the row at hand. Run over the gallery and over one drawing made with every tool.
  The third pass added simulated input to the same checks (pointer and key events raised
  on the canvas, `PointerPressedEventArgs` and friends built by hand, so the tools' real
  handlers run): every tool walked at random - clicks on paper, on figures of the kind it
  expects, drags, double clicks, Enter, its panels filled in and their buttons pressed,
  then given up by Escape, a right click, another tool; every figure dragged, with Shift
  and with Alt onto every kind of figure; deleted; copied and pasted; the grid's real
  editors driven (text set and Enter raised on the box, sliders, check boxes); and random
  sequences of all of these with undo and redo in between, then everything undone and
  redone, each state compared with what it was when first reached. Seeded, so a failure
  replays. The temporary code need not touch the repo: a partial `MainView` outside it,
  compiled in with `-p:CustomBeforeMicrosoftCommonTargets=<a .targets file adding the
  Compile items for the LiveGeometry project>` and started by a `[ModuleInitializer]`. What
  the checks can't see: anything that depends on the view (a point on an angle's mark sat
  elsewhere in the round trip's canvas) - which is how that one was found.
