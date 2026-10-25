# Expressions

The expression language: renames, cycles, the language itself, exact forms, binding and evaluation, building a formula by clicking, coloring, completion.

- **Expressions follow renames**. Compiled expressions hold the figures themselves, but their
  *text* (saved, edited) holds names. `FigureBase.Name` collects a rename wave (the figure,
  the default names that follow it, one that had to give up the name) and at the end of the
  outermost set rewrites every `IRenamableExpressions` in the dependents closure: label
  `[...]` parts, function graphs, `DrawingExpression`s (point by coordinates, line and circle
  by equation). `ExpressionRenamer` reads names as `ExpressionTreeBuilder` binds them, under
  the old names (A.X, AB as two points, the points of angle/distance/area, a Number) and swaps them
  all at once; two points whose new names run together would read differently (PB next to a
  point named PB), or one of which needs backticks, become `distance(P, B)`. Only the text changes, and not through undo: undoing the rename renames back.
  **Every name can be said** (2026-10-13): as it is when it is a name of the language (a
  letter or `_`, then letters, digits and primes: A', P_1; `Scanner.IsName`), else in
  backticks, a backtick in it doubled: `` `Drag me!`.X ``, `` distance(`my point`, B) ``.
  `Scanner.NameInExpression` writes a name so - the renamer and whatever puts a figure's name
  into an expression use it - and also puts pi, π, e (any case) and x in backticks, which
  bare are the constants and a function's variable. A name in backticks is the top-level
  figure called exactly that (`Binder.ResolveFigureExactly`): not one differing by case, not
  a part inside a composite (a slider's knob), never a constant, two points or a function.
  So the Name box takes any name for any figure. It refused, for a figure an expression
  named, a name no expression could say ("my point" left the label `[A.X]` an error), and
  pi, e and x for a number, whose texts the file read again took for the constant. A
  label's `[...]` part steps over a name in backticks (`LabelBase.FindExpressionParts`, a
  scan, not a regular expression, which backtracked without end on a run of backticks:
  `` [`a]b`.X] `` is one part). The player reads backticks too; it renames nothing.
  `pi`, `π` and `e` in lowercase are the constants whatever the drawing has; in any other case
  (`PI`) the two-points reading comes first, and with points P and I that is their
  distance. (Lowercase too, until 2026-10-01: `sin(pi * x)` came out wrong without a word
  in a drawing with points P and I. Generated expressions still say `rad(45)` or the digits
  of π.) A number (slider, Number) called exactly what the text
  says, capitals and all, comes before the two-points reading (`Binder.ResolveExactNumber`,
  mirrored in the renamer): two points match in any case, so with points A and B a slider
  named `ab` was the distance AB, and undo of renaming it could not find it in the text.
  A figure given by expressions (`IExpressionOwner`: point by coordinates, line and circle
  by equation) depends on what all of its expressions name, listed in the expressions'
  order - not on what the last one compiled named, appended (X = A.X, Y = A.Y, X edited:
  the point no longer followed A).
- **Expression cycle checks include every binding form**: bare distances, point-function
  arguments (`distance`, `angle`, `area`), numbers and property access. Labels and function graphs
  reject themselves and their descendants, just as coordinate expressions do. A function
  referencing a point on itself otherwise creates a dependency cycle.
- **The expression language calculates as it is written in class** (`Expressions/Parser`):
  a minus in front takes everything up to the next + - * /, powers included, so `-x^2` is
  -(x²) (it was (-x)²: the parabola y = -x^2 opened upward); `^` is right-associative. A
  number right before a name or a parenthesis multiplies it (`2x`, `3sin(x)`, `2(x + 1)`,
  `2AB`; 2026-10-18, as Desmos and GeoGebra read it), never two names (AB is a distance)
  nor a name before a parenthesis (a call); `1/2x` is (1/2)x. `30°` is the number in
  radians (a postfix, GeoGebra's), `|x|` the absolute value, and the comparisons `<` `<=`
  `≤` `>` `>=` `≥` `=` `==` `≠` `!=` come last and give 1 or 0, numbers equal but for
  rounding being equal (`Math.TangencyTolerance`), for `if(condition, a, b)` -
  `if(condition, a)` has no value where it doesn't hold. `−`, `×`, `·`, `÷` are the
  operators they look like. **The functions are a closed list** (2026-10-18:
  `FunctionCatalog`, a method of `Functions` each, by name in any case and its other names;
  until then every method of System.Math with doubles compiled, BitIncrement and
  Log(a, newBase) too), named as the American classroom and Desmos name them: `log` is base
  10 and `ln` e's (it was the European `log` = ln, `lg` = log₁₀), `log_2(x)` says its base in
  the name, `arcsin` (`asin` too), `ceil` (`ceiling` too), `distance`, `nthroot`, `mod` (never
  negative for a b above 0), `round(x, digits)`, `min`/`max`/`gcd`/`lcm` of any number of
  arguments (two at a time), `sec`/`csc`/`cot`. **Units**: trigonometry takes radians, as in
  Desmos and GeoGebra, and everything read off the drawing is in degrees, as the Angle tool
  shows it: `angle(A, B, C)` (the smaller), `directedangle(A, B, C)` (counterclockwise, 0 up
  to 360), `direction(A, B)` (as a segment's Direction), a slider, a Number, a label that
  stands for an angle (`Label.Angle`; it was radians, the Clock said `[rad(-30 * time)]`).
  `sin(angle(A, B, C))` is an error that says to write `angle(A, B, C)°`
  (`FunctionCatalog.TakesRadians`/`GivesDegrees`, a call written right in the argument).
  `ang`, `oang`, `xang`, `dist`, `sqr`, `int`, `sgn`, `lg`, `todeg` are no names any more:
  `LegacyExpressions` translates them (and VB6's `ANGLE`, in radians) keeping the values,
  for the `.dgf` reader; the gallery was rewritten with it once. Everything is a double: a whole number is converted (`sign(x)`, a
  polygon's `NumberOfSides` - unconverted, the first operator applied to it threw), and a
  property that is no number (`A.Name`) is an error said in words
  (`ExpressionTreeBuilder.AsNumber`). Nothing a user can type may throw, caught or not: a
  first-chance exception is an error report on screen. Checked with random texts built
  from the language's tokens and the drawing's names (20000 of them: compile as an
  expression and as a function, evaluate, run the renamer), next to a list of texts with
  known values and of wrong texts that must each give an error with words in it.
- **Exact forms** (`Expressions/LinearForm.cs`, 2026-10-13): `LinearForms` works an
  expression's bound tree (`DrawingExpression.Bound`, `CompileResult.Bound`) out as a
  constant plus coefficients times variables, all exact (`ExactNumber`: fractions
  (`Fraction`, a typed 0.1 being a tenth) times 1, sqrt(2), sqrt(3), sqrt(6), closed under
  + and *, `sqrt` of a fraction that is a square times 1, 2, 3 or 6). The X and Y of a point
  by coordinates are worked out through, each once; any other point's X and Y, a Number's
  value, any other property and a two-point distance are variables. A product of two
  variables, a function of one, an inexact constant, a division by anything but a fraction
  give no form, and nothing is equal to an expression without one: equal forms are equal
  for every position of the figures (the variables are taken as independent, which can
  only miss an equality). What decides that a regular polygon's vertex is a point the
  drawing has (see "Squares and regular polygons").
- **An expression is bound once and evaluated many times** (`Expressions/`, 2026-10-07):
  `ExpressionTreeBuilder` binds the parser's `Node` tree into a `BoundExpression` (names
  resolved to figures, numbers, two-point distances, properties and functions, every value
  a double), and `Compiler.Strategy` (`ExpressionStrategy`) says what the delegate run at
  every recalculation is made of: `OwnTree`, the bound tree evaluating itself (the
  default); `LightCompiler`, a System.Linq.Expressions tree run by that library's
  interpreter (what every expression was until then); `LinqInterpreter`, the same tree
  walked by `ExpressionTreeInterpreter`; `Compiled`, the tree compiled to IL (a dynamic
  method for the JIT on the desktop; the browser has no JIT and interprets it). All four
  give the same values: the regression suite's "Expression strategies agree" and the
  benchmark's mismatches column say so. Reflection is asked once (`ExpressionReflection`:
  a function by name and argument count, a method's shape, a method or a property getter as
  a delegate - a getter through an open delegate made by a generic method, whose
  instantiations are named so that Mono's AOT compiles them), and a `Binder` gathers the
  drawing's point names once per expression, not per identifier.
  `LiveGeometry.Desktop.exe --bench-expressions <file>` (the browser build at
  `/?bench=expressions`, read with `webauto console`) loads every gallery drawing under
  each strategy, after a warm-up pass, and prints the load, the compile share of it, the
  evaluation of every expression 200 times, the function graphs sampled, the drawings
  recalculated 20 times, and the mismatches (`ExpressionBenchmark`). 2026-10-07, the
  gallery's 3572 expressions, compile out of the load and the 200 evaluation rounds, in
  ms: desktop JIT - OwnTree 46 (8%) and 96, LightCompiler 60 and 144, LinqInterpreter 61
  and 443, Compiled 528 (68%) and 17; browser AOT - 38 (6%) and 188, 138 and 578, 54 and
  530, 277 (31%) and 96; browser interpreted - 168 (8%) and 246, 626 and 1412, 453 and
  3124, 1044 (23%) and 139 (there the later strategies also pay to warm the interpreter
  up on their own code). So expressions are a small part of a load either way (Stretchy
  Slime's 160 compile in under 2 ms under AOT) and an evaluation is a fraction of a
  microsecond; the bound tree is the cheapest to make and second only to IL at running,
  which is why it is the default. IL wins evaluation everywhere but costs 3 to 7 times the
  compile.
- **Building a formula by clicking** (2026-10-15; `Behaviors/ReferencePicker.cs`,
  `PropertyGrid/Editors/ExpressionTextBox.cs`): every box an expression is typed into - a
  number's formula, the X and Y of a point by coordinates, a function, the text of a label,
  the expression rows of the tool panels (`[PropertyGridPreferredEditor("Expression")]` on
  a string, `ExpressionStringEditor`; `"LabelText"` on a label's text) - is an
  `ExpressionTextBox` (the one class for them all, where completion, highlighting or a
  structured editor would go), and while it has the keyboard it is
  `Drawing.InsertTarget` (the box keeps the drawing it registered with: taken out of the
  tree, with a tool's panel, it can't find it through its ancestors, and stayed the target,
  so the next tool never got a click on a figure - until 2026-10-18): a click on a figure
  puts a reference to it in at the caret instead of going to the tool (`Behavior.TryInsertReference`, before `MouseDown`; the
  press is handled, so the canvas doesn't take the keyboard, `MainView`, and a drag of the
  paper keeps it), and so does a click on a row of the Figure List, the way to a hidden
  figure (`FigureExplorer.Row_PointerPressed`); the hover shows the halo and the hand,
  and the status says what a click puts in ("Insert d.ArcLength (the arc length of arc d)
  (1 of 3): Tab for the next.", `ClickChoice.Picking`), Tab from the box steps through
  the figure's quantities (`MainView.IsTabForPicking`). While the box has the keyboard a
  pencil sits at its right end (the box's InnerRightContent), since every click on a
  figure is the box's then; it is a button (2026-10-17) that switches that off and on
  while the keyboard stays in the box (`ExpressionTextBox.TogglePicking`: off, the box is
  no `InsertTarget`, the pencil is faint and a click on a figure selects it), and the
  status hint names it; the box taking the keyboard again starts with it on. In a label's
  text a [ ] button comes before it, in the brackets' color (`InsertBrackets`): the
  selection goes in brackets, or a pair goes in with the caret between - also what says
  that a text takes formulas at all, which nothing did. Enter in a
  figure's row is "done" and gives the keyboard to the canvas (`EndEditing`; a tool
  panel's Enter is its command, as ever), but not in a label's text, where it starts a
  line: there the pencil is the way out. The value under the box follows the figures it
  names (`PropertyGrid.RefreshNumbers` asks `StringEditor.RefreshLiveValue`). What goes in (`ReferenceText`):
  a figure's first quantity as the Length tool would pick it, then the others -
  `c.Radius`, `c.Diameter`, `c.Circumference`, `c.Area`; `d.ArcLength`...; `p.Area`,
  `p.Perimeter`; a slider or a number by name; a point by name (two in a row make a
  distance), `A.X` in an X row and `A.Y` in a Y row (the row's caption decides,
  `InsertContexts.FromRowName`); a segment as `AB` where the two names run together into
  that reading (checked by compiling), else `distance(A, B)`; a measurement as what it
  measures (so the formula reads like math and outlives the measurement), an angle one as
  its Sweep has it right now - `angle(A, B, C)`, `360 - angle(...)`, `directedangle(...)`,
  `360 - directedangle(...)`, in `rad()` when it shows radians - and an area one as
  `p.Area` or `area(A, B, C)`; into a label's text in brackets unless the caret is in a
  `[...]` part already (an unclosed `[` counts). A figure built on the box's owner gives
  nothing (a cycle). Lines, rays and graphs give nothing either. Clicking empty paper
  while a box has the keyboard is the tool's as ever (a click clears the selection and
  so the panel, a drag pans).
- **Expressions are colored as the compiler reads them** (2026-10-17;
  `Controls/ColoredTextBox.cs`, `Expressions/ExpressionClassification.cs`): every
  `ExpressionTextBox` and the in-place editor of a text label draw the pieces of what is
  typed in the theme's Expressions colors - the [ and ] of a label's part, figures,
  properties, functions, numbers, pi, e and x - and each [...] part of a label's text on
  a faint plate, a pink one where the part doesn't compile. The binder notes what each
  token was bound to (`CompileResult.Classification`, by the token's place in the text, so
  nothing keeps the whitespace), and only a text that compiles is colored: in a part with
  an error only the plate says so, and a [ never closed is text like the rest. A box
  follows its drawing's history (a name given to a figure, or taken away). A TextBox draws
  its text in one brush: `ColoredTextPresenter` overrides `TextPresenter.CreateTextLayout`
  and lays the text out again with runs of other brushes, in the room the base's layout
  was given (which it keeps private). The plates are a layer of their own under it
  (`PartPlates`, from the layout's `HitTestTextRange`): Avalonia sealed every control's
  `Render` in 2023, and a run's background would be the box of its glyphs, drawn over the
  selection. `ColoredTextBox`'s template is a code copy of Fluent's TextBox template of
  Avalonia 12.1 with those two in it: compare it with Fluent's TextBox.xaml when Avalonia
  is upgraded.
- **Completion** (2026-10-17; the design and its reasons, after the VS editor's, are
  `docs/completion-plan.md`): every `ColoredTextBox` with a host - the expression boxes,
  the in-place editor of a text, the theme page's sample - has a list of what can go at
  the caret (`CompletionList`, the popup and the keys; `ExpressionCompletion`, the model,
  which the regression suite tests). It opens on the first letter of a word, after a dot
  behind a figure's name (its quantities: `ReferenceText`'s, the same a click offers) and on
  Ctrl+Space (a word only one row matches is completed at once); the functions are
  `FunctionCatalog`'s, written by hand with a line of help each, a row matched by its other
  names too (`asin` selects `arcsin`, and the commit writes arcsin), and the word after a
  number is the name (`2si`). Every expression box has a ° button before the pencil (no
  keyboard has the key). Firm selection (the word
  has a letter): Tab, Enter, a click and any character that can't go on a name commit, the
  character typed after. Outlined (an empty word, after a dot): only Tab and a click, so
  that `A.` then `+` types a plus. The popup is in the window's overlay
  (`ShouldUseOverlayLayer`), so it never takes the keyboard (nothing in it is focusable: a
  click on a row leaves the keyboard in the box, which matters for the in-place editor,
  which commits when its box loses it) and shows in `winauto shot`. `MainView_KeyDown`, a
  tunnel handler, leaves Tab and Escape to an open list. A word that already reads as two
  points (AC) gets a row of its own, or `+` would put polygon ACDE in its place.
