# File format (.lgf)

No compatibility requirements: the gallery drawings are the only files that matter (see
"Backwards compatibility is not a concern"). The upgrade notes below describe what the
loader still does for files from before; none of it needs extending.

- **Colors are always `#AARRGGBB`** (`ColorText.ToArgbHex`). Never `Color.ToString()`: Avalonia
  writes the *name* of a known color. `ToColor()` accepts names too, for old files. A gradient
  is a child element (`<Fill><LinearGradientBrush>`, `<Background>...` on `<Viewport>` for the
  paper), not an attribute.
- **`<Drawing Version="1">`** (`Settings.CurrentDrawingVersion`) says label offsets are pixels:
  a `LabelWithOffset` (point labels, measurements) sits at `Offset` pixels from its anchor in
  the plane, so the gap keeps its size at every zoom. A file without the version is upgraded on
  load, after the viewport, at the zoom it opens at (`UpgradeOffsetFromUnits`) - the best guess,
  since files don't say. Anything new that places text by something in the plane should keep
  its distance in pixels the same way. A point label is also kept in an orbit around its point
  (`PointLabel.ClampPosition`); up close, what is kept clear is the box of the letters (from
  the `TextLine` ink metrics), not the line box, so a name can touch the point's outline.
- **`<Drawing IntersectionOrder="Legacy">`**: `Math.GetIntersectionOfCircleAndLine` once swapped
  P1/P2 for a line through the center, and files carried no version then, so old drawings (the
  phone ones) that pick the other intersection opt in to a swap on load
  (`IntersectionPoint.UpgradeLegacyCircleAndLineOrder`). No file in the repo carries the mark
  (`LiveGeometry.Desktop.exe --modernize <folder>` writes into a file what loading it upgrades);
  the code stays for old files from elsewhere.
- **Styles in files** (since 2026-09-29): a figure on the default style of its kind (a free
  point on `FreePoint`, a segment on `Line`) has no `Style` attribute, and gets it on loading
  (`EnsureStyleAssigned`); one on another default names it (`Style="PointOnFigure"` on an
  intersection point drawn green); only a custom style is carried as an element. The default
  names are the constants in `StyleManager` (`FreePoint`, `PointOnFigure`,
  `IntersectionPoint`, `Midpoint`, `DependentPoint` - `DependentPointStyle` in older files,
  `Line`, `SliderTrack`, `CoincidentLine`, `Shape`, `OutlinedShape`, `Text`, `Heading`, `Explanation`, and the palette ones
  named by hue: `RedLine`, `ThickRedLine`, `DashedRedLine`, `RedOutline`, `RedShape`,
  `RedPoint`, `RedBead`... - see "The default styles"; `OtherLine`, `DottedLine` and
  `OtherShape` are gone). Saved are the styles the figures name (`DrawingSerializer.Write` writes the
  figures aside first and collects their `Style` attributes - a figure may name another's
  style, a vector its arrow's) and a default the drawing changed; a default as a new drawing
  has it is left out (`StyleManager.IsUnchangedDefault`). `Name` comes first, and an
  attribute at the value a fresh style has (`IsFilled="true"`, `Dash="Solid"`) is left out; a
  missing one reads as that value. A point style's `Size` is compared with a fresh style
  showing the same character (`DrawingSerializer.FreshStyleFor`): a character's default is
  24, a shape's 10, and an emoji of size 10 was left out and came back at 24. What differs
  under another theme is a child element
  named after the theme, its attributes (and gradient elements) the properties as in the
  style's own element: `<LineStyle Name="RedLine" Color="#FFD83B3B" StrokeWidth="1.5"><Dark
  Color="#FF00BFFF" /></LineStyle>`. The paper's is the same on `<Viewport>`: `<Dark
  Color="..." />`, `<Dark><Background>gradient</Background></Dark>`, or an empty `<Dark />`
  for that theme's own paper. Loading lays the styles out in a new drawing's order, a file's
  style taking the place of the default of the same name unless it looks the same in Light
  (files from before the themes carry every default they use: those are dropped for the
  default itself, which follows the theme), others after (`StyleManager.AddWithDefaults`).
  A default made to look under a theme as it does in Light ("Same as in Light") and
  otherwise unchanged would be taken for such a copy: it is written with an empty theme
  element (`<PointStyle Name="FreePoint" ...><Dark /></PointStyle>`), and a style read
  with any theme element is kept (`FigureStyle.SaysThemes`).
  A file style under any name that looks like what a default used to be (the numbered
  copies of the phone and CD drawings: the yellow, green and gray 10 px points with a black
  rim, the 18 and 40 pt black Segoe UI text; `StyleManager.LegacyDefaults`) is dropped too,
  and its name resolves to the default (`aliases`). Not the old opaque black line: the
  default line is translucent, and every old drawing would turn gray; those stay file styles
  (`dotnet tools/stylecensus.cs -- <folder>` lists what a folder's files carry, by values).
  The order matters: it is the style picker's, and a figure that isn't a point or a slider
  takes `Line` or `Shape` by name, else the first style that fits
  (`StyleManager.GetDefaultStyle`) - but a sector or a circular segment, an arc with an
  inside that draws its own outline (`StyleManager.IsOutlinedShape`), takes `OutlinedShape`
  (2026-10-09: both a line and a shape, it took `Line`, which comes first, and had no
  fill; a circle keeps the line, a polygon's outline is its side segments). Labels leave
  out `DecimalsToShow` at the default. `LiveGeometry.Desktop.exe --rewrite <folder>` loads
  and saves every drawing of a folder, keeping each file's viewport.
- **A figure is asked before it is worked out** while a file is read: whatever a `ReadXml`
  works out (an intersection point does) sees figures that have been read but never
  recalculated nor added. A Bézier curve had no points yet (`Bezier.CurveInfo` works them
  out on demand) and a locus an empty list (`Math.GetPointOnPolylineFromParameter` answers
  "no such point"): both threw for a point on them, and opening the drawing showed an
  error, though the point came out right once the drawing was recalculated. A point on a
  figure no longer works itself out there: it is where the file says until its figure
  comes into the drawing. One that doesn't exist kept the (0, 0) it got from figures that
  were nowhere yet, and wrote it into the next file; in a session it is where its
  parameter says, also off the end of an arc that turned away (`PointOnFigure.Recalculate`),
  so that the same drawing is saved with the same numbers. A figure given by expressions
  asked then compiles what it can: what names a figure doesn't compile yet, and the rest
  must not set its dependencies (`DrawingExpression.Recalculate` keeps the file's until
  all of its expressions compile) - a line by equation `x = P.X` lost P, whenever the
  intersection read before it happened to exist.
- **A file is read as far as it goes.** What the file is comes first
  (`DrawingControl.ParseDrawing`: XML with a `<Drawing>` root; `GeoGebraReader.ReadWorksheet`
  checks for a zip; a `.dgf` must have its `[General]` section): a file that is none, or
  damaged, is said so in the status and leaves the page as it was - it used to switch to
  the editor under the file's name first, so a picture picked by mistake made a gallery
  drawing the user's own and dropped one parked behind the gallery, and any other XML (an
  exported .svg) opened as an empty drawing that Save then wrote over. Within a drawing,
  what can't be read is left out with a line in words (`DrawingDeserializer.ReportError`:
  a kind of figure or style this version doesn't have, a figure built on one the file
  lacks or on itself, a second figure of the same name, a `ReadXml` that threw), and the
  rest comes in; it threw half way, with a bare name or "the given key was not present"
  for a message. A figure that says too little for its kind - a segment on one point, a
  perpendicular to a point - is caught where it asks (`FigureBase.Point(index)`,
  `IFigureExtensions.Point`/`Line`, 2026-10-09): while the file is read the figure is noted
  (`Drawing.InvalidFigures`), made not to exist (its shape must not be laid out at
  infinity) and given a point that is nowhere, and the deserializer leaves it out with
  what is built on it once every figure is in (`LeaveOutInvalidFigures`); at any other
  time the same access is a bug and throws. A direct `Dependencies[i]` cast is not
  covered. Such a drawing keeps the list (`Drawing.LoadErrors`), gets no name and
  so no file to be saved over, and the status says the first line and how many more.
  `IniFile` skips the lines of a `.dgf` it can't read. A viewport without a size (a file
  saved from a window that had none) keeps the view (`CoordinateSystem.SetViewport`): it
  threw, and the file was refused whole.
- **Label text** is one attribute: a line break is the two characters `\n` and a
  backslash is two backslashes (a typed `C:\notes` came back as two lines); characters an
  XML file can't hold are left out (one pasted in made Save throw and write nothing).
- **A show/hide box is read as a box** (`ShowHideControl.ReadXml`): each figure says in
  the file whether it is hidden. Applying the box to its figures on load hid again one
  that had been shown by hand since the box was last clicked. (The `.dgf` reader does
  apply it: those files have only the buttons.) The box exists whatever its figures do
  (`ShowHideControl.UpdateExistence`): built on them like any figure, it vanished while one
  of them didn't exist (a firework that hadn't burst yet). Its empty box takes the
  caption's color, as the caption does: Fluent's outline was dark on a drawing's own dark
  paper.
- **Saved files declare `encoding="utf-8"`**; files from older builds say `utf-16`, which our
  own loader tolerates but `XDocument.Load` does not.
- **`TranslatedPoint`** without `DistanceSource`/`DirectionSource`/`FreeDistance`/`FreeDirection`
  is the old format (typed values as attributes, roles by position, direction in radians);
  `UpgradeLegacyValues` gives the typed values auxiliary Numbers. A `RotatedPoint` or
  `DilatedPoint` always has its angle or factor as its third dependency (since 2026-09-29;
  the `Angle`/`Factor` attribute before that is not read).
- **`AngleMeasurement`** says `Radians="true"` when it shows radians (it was not saved).
- **Scenes** (`<Scene Left Top Right Bottom />` under `<Drawing>`, 1 or 2) are opt-in suggested
  views for drawings whose content has no useful bounds (endless ground: Castle, The Falling
  Ladder; a locus: Spiral). Where they exist, fit and tile show the scene nearest in shape to
  the room instead of the content bounds, and a gradient paper spans the scene, not the canvas.
