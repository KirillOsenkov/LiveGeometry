# Styles and themes

Default styles and the palette, point shapes and emoji, dashes, textured fills, the color picker, the paper, the chrome's theme and how drawings follow it.

- **Default point styles are by kind** (`StyleManager.AssignDefaultStyle`, looked up by name):
  the draggable kinds (`FreePoint` yellow, `PointOnFigure` green) at size 10, constructed ones
  at 8. A drawing from a file brings its own styles, and loading adds the defaults it lacks
  by name (see "Styles in files"). Gallery point sizes follow the same standard:
  `dotnet tools/pointsizes.cs -- <folder> [--apply]` lists and raises undersized styles. The
  Rose's 90 control points stay at 5 px on purpose (at 10 they swallow the flower).
- **The default styles** (2026-10-03) are a palette of eight hues (`Styles/StyleHue.cs`:
  gray, red, orange, brown, green, cyan, blue, purple), a column each in every style picker,
  in full rows of eight: points (the kind defaults and red, blue, purple at 8; beads, 12 with
  a highlight), lines (thin 1.5, thick 2.5, dashed), shapes (outlined with a flat fill -
  on the light paper the gradient's lighter end, on the dark one the dark stroke 11% of the
  way from the paper at the same alpha (`StyleHue.DarkSolidFill`, 2026-10-17: the lighter
  end there was a slab, mid gray for `OutlinedShape`, which every sector and circular
  segment takes), so that each sits about as near its paper (CIELAB ΔE 2-5 in Light, 6-12
  in Dark) and mutes what is under it alike - outlined with the gradient, and gradients
  without an outline, but for the brown and green columns of that row: the fill of a new
  polygon, `Shape`, flat yellow from the theme's `ShapeFill` as it always was, and the
  classic flat green, `GreenShape`). Gray is the theme's where a default is (`Line`,
  `ThickLine`, the point fills). A hue is a stroke per paper; fills are diagonal gradients
  worked out from it (`Tint`, light to deeper on the light paper, bright to deep on the dark
  one, translucent; light tints at most 0.75 saturated, or cyan and orange glared). The
  brown column's fill hue is the ribbon's gold: of its own hue it was too near orange. The
  picker is a grid of fixed
  cells, eight to a row (`StylePickerEditor.CellSize`); a style kept for one purpose
  (`SliderTrack`, `GalleryLocus`) is offered only to its figure (`StyleManager.IsOffered`),
  or a row would be one short. A new kind of default wants a full row. "Create new style"
  on a figure that can be filled but has a line style (a circle takes either) makes a shape
  style, so that the fill is there to edit: the same stroke under every theme, a hint of
  its color for a fill, Filled unticked (`ShapeStyle.WithStrokeOf`). A polygon takes
  either too (2026-10-13; `[StyleFor(typeof(PolygonBase))]` on `LineStyle`): a line style
  is the one-click way to no fill, its outline only, and its default is still the yellow
  `Shape` (`GetDefaultStyle` prefers the line only for an `ILinearFigure`). An unfilled
  polygon is hit as a circle is, on its outline where that draws something and inside only
  where filled (`PolygonBase.HitTest` reads the shape as drawn; its shape has no fill of its
  own, or a line style applied first left the factory's yellow). Under side segments, an
  outline in a line style shows through them darkened: the segments are drawn over it.
  Text styles are three (2026-10-16): `Text` for labels, `Heading` and `Explanation` for a
  drawing's caption (the style alone makes a label part of it, see "The caption"). They
  were `GalleryTitle` and `GalleryText`, next to a 40 px `Heading` and a `Hyperlink` that
  nothing used, and the picker showed all five as the word "Text". It lists text styles,
  not swatches: each its name in its own look over what it is for
  (`StyleManager.DescribeTextStyle`).
- **A figure drawn over another as it is built shows through** (`Figures/Coincidence.cs`,
  2026-10-15): an arc on its circle, a segment on its line looked like nothing at all in
  the same default line. When a tool adds a figure that would get the default `Line`
  (`FigureCreator.AddFiguresAndRestart`), and it lies on a figure the drawing shows - the
  same circle (one center, and a point of one on the other: its second point, a point on
  it or on a figure proven the same, a crossing with it; or the same radius source), the
  same ellipse (the same three points), the same line (both points of one on the other) -
  and covers a part of it right now (arcs and segments on one circle or line may lie
  apart), it gets `CoincidentLine` instead: 4 px of translucent cyan (#4F00BAD7, both
  themes), so every one more deepens the color. Points are the same when they are one
  figure or by `LinearForms` (see "Exact forms"); never by where they happen to be. Decided
  once, a saved style like any; offered in the picker only to a figure that has it, and an
  arc in it converted to a sector takes the default outline.
- **Point shapes and emoji** (`PointStyle.Shape` / `Character`; `Size` is the shape's or the
  character's, whichever shows: the style keeps both in memory and the file only the one in
  use, so `Character` must be read before `Size`, which declaration order does): every point
  is a `PointMarker` (circle, triangle, square, diamond, pentagon, hexagon; the polygons reach
  past the circle by eye so they look as big), or one character instead, in the embedded
  Twemoji font (`Main/Avalonia/Fonts`, CC-BY, credited in the Emoji tab) with Inter named as
  the fallback: the browser has no system fonts, and a character neither has is not offered
  (`EmojiFont.CanDraw`). The font is 1.5 MB and loaded on first use, or as soon as the
  gallery is built, whose first rows show emoji (`EmojiFont.Open`: a file beside the desktop
  exe, a fetch of `fonts/` in the browser, brotli via web.config and cached as immutable: a
  different font must get a different file name). In the browser `main.js` fetches it, on
  the gallery's page as soon as the page loads (at a low priority, after the runtime's
  files), and hands it over in one piece (`fetchEmojiFont`, `copyEmojiFont`). The browser
  reads a response on the page's thread a piece at a time, and HttpClient took a turn of
  the UI thread for every step besides: fetched once the gallery was up, the font came in
  between the tiles' loads, seconds after its bytes. A font that fails to come (a network
  error) leaves the characters in the fallback font for good: each `EmojiGlyph` asked to be
  drawn again whenever it was drawn without the font, and the page drew without end. The
  style's editor has Shape | Emoji tabs (`IPropertyGridTabs`: the tab shown is what the style
  is; picking Shape drops the character, undoably; a row on two tabs, Size, gets an editor on
  each). "What the style is" is what it is under the theme on screen (`ShownCharacter`,
  and `ThemedValue` reads a property without an override from the resolved style): by
  the base values an emoji picked under Dark opened on the Shape tab, could not be taken
  off, and showed the shape's size. The Emoji tab searches `Emoji/Emoji.txt`
  (CLDR names and subgroups of the single-character emoji the font has): regenerate it with
  `dotnet tools/emoji.cs -- <emoji-test.txt> <font> <Emoji.txt>` when the font changes.
  A color emoji is painted with the style's `Fill`, so the fill's alpha is the emoji's
  opacity: a style written by hand without a `Fill` takes the shape default's (alpha 100)
  and comes out faded.
- **Dashes**: `LineStyle.Dash` is put on in `LineStyle.OnApplied`, not through a setter, because
  `StrokeDashArray` counts in stroke widths. Anything that
  applies a style to a shape by hand (sample glyphs) must call `OnApplied` too. Any enum
  property of a style or figure round-trips by name through the generic `EnumSerializer`.
- **Textured fills** (`Styles/FillTextures.cs`, 2026-10-25; the player's twin is
  `styles/fillTextures.js`): a shape style's `Texture="Crayon"` lays its fill down as a
  colored pencil does, `Texture="FeltTip"` as a felt-tip pen does. Each tile is a scan:
  `CrayonTile`, `FeltTipTile` (`Styles/<Name>Tile.cs`, `styles/<name>Tile.js`, both written by
  `dotnet run tools/texturetile.cs -- <Name> <png> <x> <y>`; the tool's header has the
  commands that made them) are bytes of pigment density against the average (128) from an
  evenly colored patch of one of Catriona Agg's drawings, measured from the paper's white,
  the photo's light divided out (`--flatten`: the felt tip keeps its stroke bands), specks
  clamped, and the edges cross-faded so it repeats without a seam. The felt tip is the
  bright yellow hexagon of Three Hexagons on a Circle, whose ink nearly saturates the scan,
  its variation scaled up (`--spread 0.09`) to what her pink measures; no clean pink area is
  big enough. A fill is the
  paper moved toward the fill's one color (a gradient's middle) by the density, so it
  averages to the color, flecks fade to the paper and dense spots deepen; pick the color
  as the average of the area to match. A procedural tooth (value noise) was tried first and
  looked sharper and sparser than pencil. The tile is opaque on the drawing's paper (a
  fleck shows the paper, not the shape under it: one region of a pencil drawing is colored
  on the paper, not over another), so the brush is per color and paper, cached; the
  player's renderer takes the paper per frame (`paperColor`) and anchors the pattern to the
  canvas, the app's `ImageBrush` to the shape, one tile pixel per canvas pixel (drawn
  smaller it blurred and lost its grain). A paper changed by hand re-textures at the next
  style application (a theme switch does).
- **Color/brush picking** (`LiveGeometry/Controls/ColorPicker/`) is layered so parts can be
  swapped: `ColorPalette` -> `ColorPage` (swatches, spectrum) -> `ColorPickerView` ->
  `BrushPickerView` (solid | gradient); in the property grid through `ExpandingPickerEditor`.
  A gradient is its stops and an angle; the brush's two points are the line through the
  middle of the box at that angle, long enough for the end colors to be reached in the
  corners (the CSS rule: (0,0)-(1,1) at 45°, edge to edge at 0° and 90°, outside the box in
  between), so that the stops reach every part of the shape. Gradients made before
  2026-09-29 have a line of length 1, which leaves the corners of a diagonal one flat; they
  load as they are and take the new line when edited.
  Whoever hosts a `SegmentSwitcher` sets its `Surface` to the background it sits on so the
  selected tab blends into it. A click on a swatch gives the page the keyboard, and the
  arrow keys move the current swatch through the grid and pick as a click does
  (`SwatchPage.OnKeyDown`; one undo step for the run, the canvas doesn't pan meanwhile).
  On the Spectrum tab the surface pressed last (field, hue, opacity) takes them as fine
  steps, 1% or 1°, ten with Shift (`DragSurface.Stepped`); a step of 1/255 left the color
  as it was once rounded, and a press that changes nothing looks broken.
- **The paper** is `Drawing.Background`, edited through "Drawing background" on the settings
  page (`AppSettings.EditDrawingBackground`; `DrawingHost.ShowDrawingProperties` puts the
  drawing itself in the property grid). "Reset to default" (the theme's paper) shows only while
  the paper is not the theme's, and "Same paper as in Light" only while there is a Dark override: both come and go
  as the paper changes (`[PropertyGridLiveCondition]`: the button is made hidden and the grid
  asks `CanEdit` again on each named `PropertyChanged`; the divider over destructive buttons
  hides with them), since the usual rebuild on a change
  without a name would fold the color picker under the cursor. A gallery tile takes a drawing's paper as its plate and turns its caption white on a
  dark one.
- **The chrome's colors are a theme** (`UI/AppTheme.cs`, not `Theme`: every control has a
  `Theme` property, its ControlTheme, which would shadow the class). One `Color` property per
  role (`Strip`, `HeaderRow`, `Background`, `Text`, `Ink`, `Accent`...), and `AppTheme.Light`
  and `AppTheme.Dark` are the values; each theme is a `ThemeVariant` whose resource dictionary
  of brushes (under the property names) is registered on the application, and the chrome binds
  to those resources - the code side of DynamicResource - through `ThemeBinding`:
  `border.BindTheme(Border.BackgroundProperty, nameof(AppTheme.Background))`. A control whose
  color follows its state binds again to another name (a null name unbinds, `whenNone:` gives
  the plain value); pens made in `Render` come from a styled property bound the same way with
  `AffectsRender`; a gradient of theme colors is rebuilt from `ObserveTheme`. Never copy a
  color out of a theme brush into a plain brush: it stops following. Fluent's own controls
  (text boxes, combos, scroll bars, menus, tooltips) follow the variant on their own; the
  property grid's styles use `DynamicResourceExtension` setters (`PropertyGridTheme`).
  Switching is `AppTheme.Apply(choice)`: a theme's name, or `System` (`RequestedThemeVariant
  = Default`, the OS's or the browser's `prefers-color-scheme`); `AppTheme.CurrentChanged`
  is for what can't bind. A theme is a `[PropertyGridNoUndo]` object with a color picker per
  row: the settings page (the gear after Redo on the toolbar) opens `AppTheme.Current` through "Theme
  colors", every pick repaints the app at once (the setter writes the dictionary, which
  notifies its owner), and "Copy as code" puts the initializer on the clipboard to paste
  back into `AppTheme.cs`. A write to the dictionary makes every resource binding in the
  app (Fluent's templates included) look its resource up again, whatever the key, and
  a drag in the picker is a set per pointer move: the setter changes the property at
  once but posts the dictionary writes (`AppTheme.SetResource`, one flush per idle
  tick), and `ColorsChanged` - the drawings' refresh - fires after the flush, only for a
  color drawings take (`IsDrawingColor`: the Paper group and `Text`). `Drawing.RefreshTheme`
  applies each figure's style once (`IsRefreshingTheme` holds the styles' own notifications
  back while they re-read the theme), and only while the drawing is on screen: a hidden
  gallery tile or the parked drawing catches up through `RefreshThemeIfStale` against
  `AppTheme.Version` when it shows again (the gallery's `IsVisible`, `MainView.ShowEditor`,
  the attach to a canvas; Avalonia's `IsEffectivelyVisibleChanged` is internal). Without
  this, one pick in the browser (interpreted, no AOT) took over a second. A new theme is another instance with its own variant, inheriting
  Light or Dark for Fluent's sake, added to `All`. Tool icons draw their lines in `Ink`
  (`IconBuilder` binds them, and takes a theme color's name where a `Color` was passed);
  the drawn chrome icons (`MainToolbarIcons`, `PropertyGridIcons`) use a sentinel brush that
  `Shape` rebinds to `IconOutline`. A tool icon's fills are theme colors too: the Paper group
  for what a figure would look like (points), and the
  Icons group for what is only a picture (`ShapeIconFill`/`ShapeOutline` of every shape in
  the Shapes icons, `ImageFill` of a transformation's image - the figure transformed is
  filled as a shape is, though outlined in ink like its image -
  `RulerFill`/`RulerOutline`, `AngleFill`/`AngleOutline` with `ScaleMarks` on both,
  `AreaFill`/`AreaHatch`, `LineAccent` for the line or curve a
  tool makes out of the figures its icon also shows: `IconBuilder.AccentLine`, 1.5 thick)
  - no literal brush in a `CreateIcon`,
  or the theme can't reach it. A shape's icon is not filled with the Paper group's
  `ShapeFill`, the fill of a new polygon: that one is translucent, made for the paper, and
  a gradient there would fill every new polygon with it. `ShapeIconFill` and
  `ImageFill` are a `Brush`, not a `Color`, so
  the grid gives them the brush editor and they can be gradients; to make another fill one,
  change its type and its two initializers (a property bound to it must take a brush, and
  `ObserveTheme` hands out solid colors only). The sun/moon
  beside the Octocat (both pages) flips between light and dark, and landing on what the system
  says stores `System` again. On Windows the title bar goes dark too
  (`LiveGeometry.Desktop/WindowFrameTheme.cs`: the DWM attribute, and then a non-client
  activation cycle plus a frame-changed `SetWindowPos`, since Windows 10 keeps painting the
  old shade until the next activation otherwise). `winauto shot` doesn't show the frame at
  all (PrintWindow): to check it, `shot ... --screen`. On a Mac Avalonia shades the title bar
  itself.
- **Drawings follow the theme through their styles.** A `FigureStyle` has its values (how it
  looks under Light, the base theme) and may hold *overrides* for another theme: the
  properties that differ, with their values (`Overrides`, `SetOverride`). Whoever draws with a
  style resolves it first (`IFigureStyle.Resolve`: the style itself without an override for
  the theme on screen, else a copy with the override applied); `GetWpfStyle` and the
  `Apply` extension do that, so figures and sample glyphs need nothing. The default styles
  (`StyleManager.AddDefaultStyles`, the grid's own in `CartesianGrid`) take their colors from
  the theme's Paper group (`AppTheme.Paper`, `Ink`, `Line` - the default line, translucent
  black in Light and an opaque light gray in Dark, where a translucent one came out too dim -
  `SliderTrack`, the point fills, `ShapeFill`, `Axis`, `GridMajor`/`GridMinor`) through `BindToTheme`: base from Light, an override from every
  other theme, read again when a theme color is tweaked (`AppTheme.ColorsChanged` ->
  `Drawing.RefreshTheme`) - except where the user has given the property another value
  since (a red fill for the default point style): `FigureStyle.ReadFromTheme` remembers
  what the binding last gave each theme and leaves a value that differs alone. It was
  overwritten whenever the themes were read again (also when the theme switched while the
  drawing was behind the gallery), and then not saved. The palette's colored styles (see
  "The default styles") are literal, with a literal Dark override each. A drawing's paper is
  the theme's unless it has one of its own (`Drawing.OwnBackground`, null for the theme's;
  files leave it out; the readers of foreign formats take white as none). A theme switch
  (`AppTheme.CurrentChanged`) re-applies every figure's style and the paper
  (`Drawing.RefreshTheme`, hooked by `DrawingControl` and the gallery's `DrawingThumbnail`,
  which doesn't paint the paper: `Drawing.PaintsPaper`). The canvas's own binding to the
  theme's paper is at `BindingPriority.Style`, below the value the drawing sets: at the same
  priority the binding pushes the theme's paper over a drawing's own on every switch. The
  grid's colors go by the paper it is on (`CartesianGrid.GetColor`): on a solid paper of the
  drawing's own, those of the theme whose paper is nearest in lightness, shifted by as much
  as the paper differs from that theme's. The tool icons draw their points,
  fills and lines from the same Paper group, so they show the figures as the theme would.
  **The property grid edits what is on screen**: under the base theme (Light) a style's
  property itself, under any other theme its override for that theme
  (`ThemedValue.ForCurrentTheme` wraps the value the grid edits; `IThemeOverridable` is
  what a style and a drawing's paper implement), and undo puts the override back, or takes
  it away if there was none. A "Same as in Light" button drops the theme's overrides (shown
  once there are some, on the next opening of the style). The paper works the same:
  `Drawing.Overrides` holds the paper chosen for another theme, "Reset to default" under Dark
  stores a null override (that theme's paper). Both buttons are undo steps.
  `Heading` (a Dark override in the splash's blue), `Explanation` (the chrome's text
  color) and `GalleryLocus` are defaults too. The gallery drawings' own styles got their
  Dark overrides from `dotnet tools/darken.cs -- <folder> [--apply]` (2026-09-29): a stroke,
  text or fill darker than lightness 0.35 is lightened to the same hue at 0.82 - 0.4 ×
  lightness (black lands on the ink), a drawing with a paper of its own is left alone, and a
  style that has a `<Dark>` already is never touched, so it is safe to rerun after adding a
  drawing. The three drawings with a light paper of their own (Castle, Conic, Rose) carry
  `Heading`/`Explanation` copies whose Dark override is the light color, so the caption
  stays dark on their paper. `--check <folder> <out> --dark` renders the pictures under the
  dark theme; a contact sheet of the gallery in each theme is the way to review. Anything
  that reads a style's color itself rather than through `Apply` must resolve it first
  (`Arrow.ApplyStyle`). A show/hide check box takes a text style like a label
  (`ShowHideControl.ApplyStyle`, `Text` by default, `Explanation` in the gallery) and pins
  Fluent's per-state caption resources to that color, or the hover would repaint the
  caption in Fluent's own. The gallery's tiles keep their one list of pastels: under the dark
  theme a tile deepens its pastel (`GalleryTile.Deepen`: same hue, value 0.24) and derives
  a lighter border from a dark plate; a drawing's own paper on a tile is the paper as the
  theme resolves it. A tile follows `AppTheme.CurrentChanged` on its own.
