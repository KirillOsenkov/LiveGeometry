# .dgf (DG 1.0) reader facts

Learned from `Reference/VB6/Source` while making the CD library load (`DGFReader.cs`):
- `modFileIO.bas` writes an `AuxInfo(n)` / `AuxPoints(n).X/Y` only when it is not 0: absent
  means 0 (`GetAuxInfo`/`GetAuxPoint`). An analytic line `x - y = 0` has no `AuxInfo(3)`.
- A point's `Type` is the DrawState that made it; 0 = free point, and old files write
  `ParentFigure=0` for those (which is *not* Figure0). Every dependent point has `Type != 0`.
- `ParentFigure` of a point and figure indices are 0-based; Points, Labels, Buttons 1-based.
- Buttons: type 0 show/hide, 1 message box, 2 sound, 3 open another drawing. Only 0 becomes a
  figure; a show/hide button's list may reference the others. The list goes into the
  `ShowHideControl`'s own `Dependencies` (what it shows and hides, and what the `.lgf` saves);
  `AddDependencies` alone only registers the dependents side and the boxes do nothing.
- DG's angle bisector (`Math.bas GetBisector`) is the *interior* bisector and a whole line;
  the reader leaves the new bisector's `Sweep` at `Smaller` (the angle under 180°, see
  "Which of the two angles") and sets `IsLine` ("Whole line", saved as `Line="true"`); the
  verbs it would inherit from `Ray` (Convert to line/segment, Reverse) are vetoed in
  `CanEdit`. Morley's trisectors are expressions and get the same effect from
  `SGN(pi - OANG(...)) * ANG(...)` - `OANG` is the counterclockwise angle in [0, 2pi).
  VB6's names are translated into ours as the file is read (`LegacyExpressions`, `fromDG`:
  VB6's `ANGLE` is radians, like `ANG`).
- A point on a figure is placed by moving it to its saved X,Y in one `MoveTo` (setting X then Y
  projects twice from off the figure and lands elsewhere); `AuxInfo(1)` (t on a line, clockwise
  angle on a circle) is ignored.
- `AuxPoints(6)` of a measurement is the label's shift from its default place, in pixels, y
  down - taken as our `Offset` as it is.
- Intersection points are picked by the saved coordinates of both solution points, so the
  Legacy circle/line order problem does not apply to `.dgf`.
- The VB6 expression language has more than ours: `[A,B]` distance with a comma, comparison and
  logical operators, `IF`/`MAX`/`MIN`, `°` inside brackets. Those labels show the error text;
  the compiler never throws out of a label (`Compiler.CompileExpression` catches).
- `IniFile` skips blank lines (every CD file has them between sections).
- A polygon lists its first point again at the end, to close it; the reader drops repeated
  vertices (kept, the inner triangle of Morley's was a "Quadrilateral" named JJKL). The
  gallery's polygons had the same repeat, also the hand-made and generated ones: removed
  2026-10-01.
- **`.dgf` paper**: `PaperColor1`/`PaperColor2`/`GradientPaper` of `[General]`, top to bottom.
