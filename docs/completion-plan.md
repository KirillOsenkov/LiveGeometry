# Completion in expression boxes

A completion list for the boxes an expression is typed into, modeled on the Visual Studio
editor's async completion (`C:\Editor`, `src/Editor/Language/Impl/Language/AsyncCompletion`),
kept as small as it can be. VS rules are cited as file:line where they decide something here
(S = `AsyncCompletionSession.cs`, H = `CompletionCommandHandlers.cs`, M =
`DefaultCompletionItemManager.cs`, PA = `Text/Impl/WpfView/PopupAgent.cs`).

## Where it works

- Every `ExpressionTextBox`: a number's formula, the X and Y of a point by coordinates, a
  function graph's function, the expression rows of the tool panels, a label's text.
- The in-place editor of a text label (`InPlaceTextEditor`).
- In a label's text only inside a [...] part: after a [ that no ] has closed before the
  caret (`ExpressionTextBox.IsInBrackets`). Outside the brackets the text is prose.
- Never inside a name in backticks.

What a box needs is what its coloring needs already: the drawing, the figure whose
expression it is (nothing built on it may be offered: a cycle), and the kind of text
(expression, function in x, label text). Today `ExpressionTextBox` has that through its
`IExpressionHost`, and the in-place editor through a classifier lambda. Proposal: one
`ExpressionContext` (Drawing, Owner, Kind) on `ColoredTextBox`, used by the coloring and
the completion alike.

## The model (no UI)

`ExpressionCompletion.At(text, caret, context)` gives nothing (no completion here) or:

- the **span**: the word the list replaces,
- what kind of list it is: **names** or **members of a figure**,
- the **items**: display text, insert text, icon, help.

Pure, so the regression suite can test it: text and caret in, span and items out.

### The span

The word around the caret: from the caret left over name characters (letter, digit, `_`,
prime, as `Scanner` reads a name), and right over them too, so that a commit in the
middle of a word replaces the whole word. The word must start with a letter or `_`; a
digit starts a number, not a name (`2.5` offers nothing).

- The character before the word is a `.`, and before that is a name (or a name in
  backticks) that is a figure: **members** of that figure. The word may be empty (right
  after the dot).
- Otherwise: **names**.

VS gets the span from the language and tracks it edge-inclusively (S:321): typing at the
caret grows it, Backspace shrinks it. Same here: the span is worked out again from the
text and caret after every keystroke, which comes to the same thing for one word.

### The items

**Names**, in one list sorted alphabetically, ignoring case (VS sorts once, M:38-48, and
filtering keeps that order):

- **Figures** an expression can say: points, sliders and Numbers, text labels that say a
  number, and figures that have quantities to read (`ReferenceText.OptionsFor` gives
  some). Top-level, not `Auxiliary`, not the box's figure nor anything built on it.
  Inserted as `Scanner.NameInExpression` writes them (in backticks when needed).
- **Functions** from a catalog written by hand, with a signature and a line of help each
  (`FunctionCatalog`, see below), which is the whole language since 2026-10-18. A row is
  matched by its other names too (`asin` shows `arcsin`); the commit writes the main one.
- **pi** and **e**, and **x** in a function graph's function.

**Members** of figure F: the quantities `ReferenceText.OptionsFor(F)` gives as `F.Name`,
with its descriptions ("the radius of circle c"). These are the same quantities a click on
F offers, so typing and clicking agree, rather than every numeric property reflection
finds (`ZIndex`, `DecimalsToShow`...). A property the list doesn't offer can still be typed.
A figure with no members opens no list on the dot. One addition: a segment's `Length`.
A click puts a segment's length in as its two points (`AB`), but after a dot people type
`AB.Length`.

**The word already means something.** A word that already names something gets an item of
its own, even when it isn't in the list: two points run together (`AC` with points A and C),
or a figure under another case. Otherwise `AC` with a polygon `ACDE` in the drawing would
select ACDE, and the `+` typed next would commit it.

### The help for each item

| Item | Help |
|---|---|
| Function | `sin(x)` - the sine of x, x in radians |
| Figure | its title and construction, and its value now: "Circle c with center A through B", "slider a = 2.50" |
| Member | its description and value now: "the radius of circle c = 2.83" |

## Filtering and selection

- **Filter**: the items whose name starts with the word, ignoring case. If none do, the
  items that contain it. Narrowed from the first character. VS keeps the whole list for a
  one-character filter (M:271), but our lists are short.
- **Best item**:
  1. exact match, case and all;
  2. exact match ignoring case;
  3. the first prefix match in the list's order;
  4. the first substring match.
  The matched part of each row is bold.
- **Nothing matches**: the list closes. VS keeps the previous list, soft-selected
  (S:2554-2594), which is more state for little gain. Since only a word's first character
  opens a list, more letters typed into the same word don't bring it back. Ctrl+Space does.

### Hard and soft selection

Kept from VS: it is what makes a list safe to have open while typing math.

- **Hard** (a filled row): the word has a letter and matches. Tab, Enter, a click and the
  commit characters commit.
- **Soft** (an outlined row): the word is empty - right after a dot, or Ctrl+Space with
  nothing typed. Only Tab and a click commit. Enter and the other keys close the list and
  do what they always do (S:1062-1079). So `A.` and then `+` types `+`, and doesn't put
  `X` in.
- The first Up or Down in soft selection only makes the outlined item hard; it doesn't
  move it (S:1494-1501). Otherwise Down from the outlined first item would skip it.

## When it opens

| Event | List | Span | Selection |
|---|---|---|---|
| A letter or `_` typed as a word's first character, where completion works | names | that character | hard if something matches |
| `.` typed right after the name of a figure with members | its members | empty | soft |
| **Ctrl+Space** | what the caret's place gives | the word at the caret | empty word: soft; exactly one match for a typed word: commit at once |

For Ctrl+Space, VS's Ctrl+Space is "complete word, commit if unique" (S:2566-2571), and
its plain list is Ctrl+J. We have only the one key.

It does **not** open:

- on a digit;
- on Backspace or Delete (VS: only if the language asks, H:213-217);
- when the program sets the text (the grid showing a value, a renamed figure);
- when a click on a figure puts a reference in;
- on letters typed further into a word after the list was closed.

## While it is open

- **A name character typed**: the span grows; filter again.
- **Backspace in the span**: filter again. When the span becomes empty, show the whole list,
  soft. Backspace once more (into the dot or the word before) closes the list (S:2409-2443).
- **The caret leaves the span** (Left, Right, Home, End, a click elsewhere in the box):
  close. The span is inclusive at both ends (S:2291-2318).
- **Any other character typed** (`(`, `)`, `.`, `,`, `+`, `-`, `*`, `/`, `^`, space, `]`):
  a commit character.
  - Hard: commit, then type the character.
  - Soft: close, and type the character.

  The character is typed after the commit, so `.` and a letter can open the next list
  (H:630-673): `ci` + `.` gives `circle.` with circle's members open. One rule instead of
  a list of characters: whatever can't go on a name ends it.

### Keys

| Key | Hard | Soft |
|---|---|---|
| Tab | commit | commit |
| Enter | commit only, the box's Enter not done (VS swallows it, CommitBehavior.cs:15-31) | close; the box's Enter (done, new line, the panel's command) |
| Escape | close; nothing else | close; nothing else |
| Up / Down | move, stopping at the ends, no wrap | make the outlined item hard |
| Page Up / Page Down | move by the rows shown | move by the rows shown |
| Home / End, Left / Right | the box's (they move the caret, which closes the list once it leaves the span) | the box's |
| Ctrl+Space | filter again | filter again |

- Escape doesn't go on to the canvas the way it does when no list is open; the next
  Escape does.
- Tab is the list's while it is open. Otherwise Tab is what it is today: stepping through
  the quantities of the figure under the cursor.

### Closing

The list also closes when:

- the box loses the keyboard;
- the text changes by any other way than typing (a click puts a reference in, undo, a
  paste, the pencil or the `[ ]` button, the grid showing the value again);
- a commit happens.

## The popup

- **Kind**: an Avalonia `Popup` with `ShouldUseOverlayLayer = true`. It is drawn in the
  window's overlay, so it never takes the keyboard, is the same on the desktop and in the
  browser, and shows in `winauto shot`.
- **Place**: below the text at the span's start (`TextLayout.HitTestTextPosition(span.Start)`
  in the presenter, as `PlacementRect`), with `Placement = BottomEdgeAlignedLeft` and
  `PlacementConstraintAdjustment = FlipY | SlideX`. Above when there is no room below, and
  slid to stay in the window (PA:191-281). It stays put while the word is typed (PA:87-89).
- **Rows**:
  - the icon of what the item is: a figure's ribbon tool icon (`FigureIcons`), one for a
    function, one for a reserved name;
  - the name, its matched part bold;
  - a faint kind at the right ("point", "slider", "function").

  The chrome's font size, whatever the box's: a label's text is 20, the list stays 13.
- **Size**: 8 rows at most, then a scroll bar. As wide as its widest row, between 180 and
  320 px. It doesn't narrow while open, so it doesn't jump as the list filters (VS keeps
  the widest width seen, C:689-743).
- **Help**: a strip at the bottom of the same popup for the selected item, always shown.
  VS uses a separate tooltip beside the list after 400 ms. One popup and no timer is
  simpler, and the help is the point for a 13-year-old.
- **Selection look**: hard selection is a row filled in the chrome's checked color
  (`ButtonChecked`); soft selection is the same color as an outline only (X:147-189).
- **Mouse**: a click on a row commits it, and the wheel scrolls the list. VS selects on one
  click and commits on two.

## Keys and the rest of the app

- `MainView_KeyDown` is a tunnel handler, so it sees keys before the box. It takes Escape
  (the keyboard goes to the canvas) and Tab (the next quantity of the figure under the
  cursor). An open list must be asked first, as `Behavior.AbandonGesture` is asked first
  for Escape: `MainView` asks the focused box `HandlesCompletionKey(e)` before its own Tab
  and Escape.
- Up, Down, Page Up, Page Down and Enter: a tunnel handler on the box itself, ahead of the
  TextBox's own handling and of `StringEditor`'s Enter (the commit of the row, a panel's
  command).
- Ctrl+Space already reaches a text box: `MainView` handles only Ctrl+S, O, N and F1 from
  a text box.

## The function catalog

`Expressions/FunctionCatalog.cs`, one line each: name, signature, help, other names. The
regression suite checks that every entry and other name resolves through
`ExpressionReflection.ResolveMethod` with its number of arguments, and that every method of
`Functions` is an entry, so the catalog and the language are the same list.

| Group | Functions |
|---|---|
| Trigonometry | `sin(x)`, `cos(x)`, `tan(x)`, `sec`, `csc`, `cot` (x in radians: `sin(30°)`); `arcsin`, `arccos`, `arctan` (in radians); `atan2(y, x)`; `sinh`, `cosh`, `tanh`, `arcsinh`, `arccosh`, `arctanh` |
| Powers | `sqrt(x)`, `cbrt(x)`, `nthroot(x, n)`, `exp(x)`, `ln(x)`, `log(x)` (base 10), `log_2(x)` (the base in the name) |
| Rounding | `abs(x)`, `round(x, digits)` (a half up), `floor(x)`, `ceil(x)`, `sign(x)`, `mod(a, b)`, `min(a, b, ...)`, `max(a, b, ...)`, `clamp(x, low, high)`, `gcd(a, b, ...)`, `lcm(a, b, ...)` |
| Angles | `deg(radians)`, `rad(degrees)` |
| Logic | `if(condition, a, b)` |
| Points | `distance(A, B)`, `angle(A, B, C)` (degrees, as the Angle tool), `directedangle(A, B, C)` (counterclockwise, degrees), `direction(A, B)` (degrees), `area(A, B, C, ...)` |

## Left out of the first version

- Signature help after `(`.
- Most recently used items, preselection.
- Fuzzy and camel-case matching.
- Completion inside backticks.
- A list after Backspace.

## Questions to decide

1. Narrow from the first character (here) or keep the whole list for one character (VS)?
2. Commit characters: everything that can't go on a name, space included (here), or a
   short list?
3. A click commits (here), or a click selects and a double click commits (VS)?
4. A function inserts `sin` and `(` is typed (here, as VS), or `sin()` with the caret
   between?
5. Enter with hard selection: commit only (here, as VS), or commit and also do the box's
   Enter?
6. Members: `ReferenceText`'s quantities only (here), or every numeric property?
7. Help in a strip under the list (here), or beside it (VS)?
