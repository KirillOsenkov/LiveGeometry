#:project ..\Main\Avalonia\LiveGeometry.Desktop\LiveGeometry.Desktop.csproj

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Threading;
using DynamicGeometry;
using LiveGeometry;

namespace LiveGeometryRegression;

public class Program
{
    [STAThread]
    public static int Main()
    {
        App.UseInvariantCulture();
        AppBuilder.Configure<App>().UsePlatformDetect().WithInterFont().SetupWithoutStarting();
        Settings.Instance.AutoLabelPoints = false;
        var tests = new (string Name, Action Run)[]
        {
            ("Undo zero length", UndoZeroLength),
            ("Undo negative length", UndoNegativeLength),
            ("Point-on-circle parameter survives undo", PointParameterSurvivesUndo),
            ("Length edits across figure kinds", LengthEditsAcrossKinds),
            ("Length panel preserves endpoint on undo", LengthPanelUndo),
            ("A direction typed, fixed, made horizontal or vertical, freed, and undone", DirectionEditsAcrossKinds),
            ("A vector's X and Y move its end, and undo moves it back", VectorComponentsEdit),
            ("Point functions reject dependency cycles", PointFunctionsRejectCycles),
            ("Expression strategies agree", ExpressionStrategiesAgree),
            ("Labels reject dependency cycles", LabelsRejectCycles),
            ("Clearing label text clears display", ClearLabelText),
            ("Functions reject dependency cycles", FunctionsRejectCycles),
            ("Repeated function references survive replacement", FunctionReplacement),
            ("Construction cancellation and undo", ConstructionUndo),
            ("Dragging and Alt snapping undo", DraggingUndo),
            ("Alt dragging a point onto its neighbor joins them", AltDragJoinsNeighbor),
            ("A point inside a shape stays inside and follows it", PointInsideShape),
            ("Dragging a point says what Shift and Alt do", DragModifierHints),
            ("Two fingers drag two points at once", TwoFingersDragTwoPoints),
            ("A finger holds a point while others pan and zoom the paper", FingerHoldsPointWhilePaperMoves),
            ("Two fingers under a construction tool pinch", ConstructionToolPinches),
            ("A tap on a figure selects it", TapSelectsFigure),
            ("A finger on a graph pans with the paper fingers", FingerOnGraphIsPaper),
            ("Editor commits and undo", EditorUndo),
            ("GeoGebra dependent numeric stays live", GeoGebraDependentNumeric),
            ("GeoGebra untranslated numeric keeps its value", GeoGebraUntranslatedNumeric),
            ("GeoGebra worked-out numeric stays hidden", GeoGebraShownNumericStaysHidden),
            ("GeoGebra numeric rotation uses radians", GeoGebraNumericRotation),
            ("GeoGebra live expressions round trip", GeoGebraExpressionsRoundTrip),
            ("GeoGebra polygon centroid", GeoGebraCentroid),
            ("GeoGebra bare degree and a square on A'", GeoGebraPrimedSquare),
            ("GeoGebra constant slider stays a slider", GeoGebraConstantSlider),
            ("GeoGebra worksheet rejects unrelated XML", RejectUnrelatedWorksheet),
            ("A point renamed A' stays named in expressions", PrimeNames),
            ("A square takes the side drawn already", SquareOnSegment),
            ("A square on a side running the other way keeps its side", SquareOnReversedSide),
            ("Squares grown into a lattice share their corners and sides", SquareLattice),
            ("A point there by chance is no corner of a square; the same square again makes nothing", ChancePointsAreNoCorners),
            ("Hexagons around a hexagon and triangles around a point share their vertices", RegularTilings),
            ("The middle of a vector gives its midpoint", VectorMidpoint),
            ("A dashed vector has a dashed shaft", DashedVector),
            ("Regular polygons by their center and on a side are plain figures; the panel after changes their sides", RegularPolygonTools),
            ("The Bezier path tool clicks, pulls handles and closes", BezierPathTool),
            ("A Bezier path's handles show, Tab takes one, Alt mirrors", BezierPathHandles),
            ("Alt dragging an anchor onto its neighbor drops it from the path", BezierPathDropAnchor),
            ("A Bezier path: points on it, deleted anchors, holes, images", BezierPathFigure),
            ("A point on a Bezier path becomes an anchor, the curve the same", BezierPathInsertAnchor),
            ("Alt+click takes a Bezier path's handles while it is drawn", BezierPathAltClickHandles),
            ("A point as a handle: followed, deleted, dropped onto, split beside", BezierPathPointHandles),
            ("A Bezier path smooths its automatic handles", BezierPathSmoothing),
            ("A Bezier path among other figures: deletion, transforms, selection, images", BezierPathWithOthers),
            ("Paste into another drawing keeps the look", PasteBringsStyles),
            ("Pasting plain text is no error", PastePlainText),
            ("The axes are lines to build on, one of each", AxisLines),
            ("Tab chooses among overlapping figures", ChoiceAmongOverlaps),
            ("One click inside an angle serves the bisector, Line at Angle, Rotate and Translate", AngleAtVertexShortcuts),
            ("A filled style fills the angle under a mark", FilledAngleMark),
            ("The sweep chooses which of the two angles a mark, a number, a bisector and an arc are", AngleSweeps),
            ("A click on an arc or an angle gives the Angle and Bisector tools three points and the sweep", AngleFromClickedFigure),
            ("A sector is filled: new, and made of an arc in its hue", ArcConversionStyles),
            ("The Arc tool's panel makes a sector or a segment of the arc", ArcPanelConverts),
            ("A hidden name shows again with Show name", HiddenNameShowsAgain),
            ("Figures without a value don't exist", FiguresWithoutValue),
            ("Length measurements", LengthMeasurements),
            ("The Length and Area tools on each kind of figure: what a click takes, the cursor", MeasureToolsOnEachKind),
            ("A length's number keeps its place through a new quantity and its undo, and a path closed", LengthPlaceThroughUndo),
            ("Undo of a length's prefix keeps a typed prefix typed", LengthPrefixUndo),
            ("GeoGebra: a Radius or an Area read in expressions and texts stays live", GeoGebraMeasurementsInExpressions),
            ("A tool defined on expressions builds on its inputs", DefinedToolOnExpressions),
            ("Defined tools are stored, read back and deleted", StoredToolsRoundTrip),
            ("Deleting what a show/hide box holds, and undo", ShowHideBoxDeletion),
            ("A show/hide box clicked in the middle of a construction stays as it was", ShowHideBoxDuringConstruction),
            ("A box's Edit figures from another tab edits that box", EditFiguresFromAnotherTab),
            ("Free, drag, fix again by a point's own box: undo gives back the first value", FreedomRefixUndo),
            ("Undoing a slider's range puts its value back", SliderRangeUndo),
            ("A vector keeps its Z", CompositeZRoundTrips),
            ("A tool on a segment's length asks for a segment; a figure that can't come in leaves nothing", DefinedToolOnSegmentLength),
            ("A point inside a shape has no X or Y to type", PointInsideCoordinatesReadOnly),
            ("Alt does nothing to a point dragged with a selection", AltDragOfSelection),
            ("A changed default style comes along with its copies", PasteBringsChangedDefault),
            ("Styles renamed in one paste keep apart", PasteRenamesKeepStylesApart),
            ("A theme's emoji and size mean the same whichever was set first", ThemeOverrideOrder),
            ("A vector's direction: not offered when its start follows its end; undo puts the end back", VectorDirectionEdit),
            ("A point inside a path with holes stays out of them", PointInsidePathWithHoles),
            ("No Convert to arc while a point is kept inside the sector", NoArcUnderPointInside),
            ("A numeric label edited moves what is built on its value, and undo moves it back", NumericLabelEditMovesDependents),
            ("A number expressions name renamed pi, e, or x where a function names it is said in backticks", ReservedNamesForNumbers),
            ("Any name can be said in an expression, in backticks", QuotedNames),
            ("No conversion that takes away what an expression reads; a failed expression has no value", ConversionKeepsExpressionsWhole),
            ("Fix length, Free length, Free direction and Convert to path anchor keep a point whose number an expression reads", PointKindKeepsExpressionsWhole),
            ("The X row of two points and the function row of two graphs: checked for each, then set", SharedExpressionRows),
            ("No Free point while a graph reads its Parameter; a graph that stops compiling is drawn no more", ReleaseKeepsFunctionsWhole),
            ("An emoji of a shape's default size keeps its size", EmojiSizeRoundTrips),
            ("LGF partial load", PartialLoad),
            ("LGF rejects abstract figures", AbstractFigureLoad),
            ("LGF leaves out a figure built on too few figures", TooFewDependenciesLoad),
            ("Gallery LGF round trips", GalleryRoundTrips)
        };
        int failures = 0;
        foreach (var test in tests)
        {
            try
            {
                test.Run();
                Console.WriteLine("PASS " + test.Name);
            }
            catch (Exception ex)
            {
                failures++;
                Console.WriteLine("FAIL " + test.Name + ": " + ex);
            }
        }

        Console.WriteLine($"{tests.Length - failures}/{tests.Length} passed.");
        return failures == 0 ? 0 : 1;
    }

    static Drawing NewDrawing()
    {
        var canvas = new Canvas { Width = 1000, Height = 700 };
        canvas.Measure(new Size(1000, 700));
        canvas.Arrange(new Rect(0, 0, 1000, 700));
        var drawing = new Drawing(canvas);
        drawing.UnhandledException += (_, arguments) => throw arguments.Exception;
        return drawing;
    }

    static FreePoint AddPoint(Drawing drawing, double x, double y)
    {
        var point = Factory.CreateFreePoint(drawing, new Point(x, y));
        Actions.Add(drawing, point);
        return point;
    }

    static void Set(Drawing drawing, object figure, string property, object value)
    {
        Actions.SetProperty(drawing.ActionManager, PropertyDiscoveryStrategy.CreateValueProvider(figure, property), value);
    }

    static void UndoZeroLength()
    {
        CheckLengthUndo(length: 0);
    }

    static void UndoNegativeLength()
    {
        CheckLengthUndo(length: -2);
    }

    static void CheckLengthUndo(double length)
    {
        var drawing = NewDrawing();
        var first = AddPoint(drawing, x: 1, y: 2);
        var second = AddPoint(drawing, x: 4, y: 6);
        var segment = Factory.CreateSegment(drawing, first, second);
        Actions.Add(drawing, segment);
        string before = drawing.SaveAsText();
        Set(drawing, segment, nameof(Segment.Length), length);
        string after = drawing.SaveAsText();
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo did not restore the endpoints: " + second.Coordinates);
        drawing.ActionManager.Redo();
        Require(drawing.SaveAsText() == after, "Redo did not restore the changed drawing.");
        drawing.Figures.CheckConsistency();
    }

    static void LengthEditsAcrossKinds()
    {
        foreach (string kind in new[] { "segment", "circle", "vector", "fixed", "onLine" })
        {
            var drawing = NewDrawing();
            var first = AddPoint(drawing, x: 1, y: 2);
            IPoint second = AddPoint(drawing, x: 4, y: 6);
            if (kind == "onLine")
            {
                var line = Factory.CreateLineTwoPoints(drawing, new IFigure[] { first, second });
                Actions.Add(drawing, line);
                second = Factory.CreatePointOnFigure(drawing, line, parameter: 0.8);
                Actions.Add(drawing, second);
            }

            IFixableLength figure = kind switch
            {
                "circle" => Factory.CreateCircle(drawing, new IFigure[] { first, second }),
                "vector" => Factory.CreateVector(drawing, new IFigure[] { first, second }),
                _ => Factory.CreateSegment(drawing, first, second)
            };
            Actions.Add(drawing, figure);
            if (kind == "fixed")
            {
                figure.FixLength();
            }

            foreach (double length in new[] { 0, -2, 0.5, 2.675 })
            {
                string before = drawing.SaveAsText();
                Set(drawing, figure, nameof(IFixableLength.Length), length);
                string after = drawing.SaveAsText();
                drawing.ActionManager.Undo();
                Require(drawing.SaveAsText() == before, kind + ": undo length " + length);
                drawing.ActionManager.Redo();
                Require(drawing.SaveAsText() == after, kind + ": redo length " + length);
                drawing.Figures.CheckConsistency();
                drawing.ActionManager.Undo();
            }
        }
    }

    static void LengthPanelUndo()
    {
        var drawing = NewDrawing();
        var first = AddPoint(drawing, x: 1, y: 2);
        var second = AddPoint(drawing, x: 4, y: 6);
        var segment = Factory.CreateSegment(drawing, first, second);
        Actions.Add(drawing, segment);
        string before = drawing.SaveAsText();
        var value = new DimensionsPanel(segment).GetProperties().Single(property => property.Name == nameof(Segment.Length));
        Actions.SetProperty(drawing.ActionManager, value, value: 0.0);
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Panel length lost the endpoint.");
    }

    /// <summary>
    /// A segment's, a vector's, a ray's or a line's direction, in the panel after the tool:
    /// typed, it turns the end once; fixed, the end slides along the line through the start,
    /// and with the length fixed too it is fully determined; Horizontal and Vertical fix it
    /// along the axis the nearer way; each step undoes to the drawing before it and redoes to
    /// the one after. "firstEnd": the second end is locked, so the first turns about it.
    /// </summary>
    static void DirectionEditsAcrossKinds()
    {
        foreach (string kind in new[] { "segment", "vector", "fixedLength", "firstEnd", "ray", "line" })
        {
            var drawing = NewDrawing();
            var first = AddPoint(drawing, x: 1, y: 2);
            var second = AddPoint(drawing, x: -3, y: 5);
            var points = new IFigure[] { first, second };
            IFixableDirection figure = kind switch
            {
                "vector" => Factory.CreateVector(drawing, points),
                "ray" => Factory.CreateRay(drawing, points),
                "line" => Factory.CreateLineTwoPoints(drawing, points),
                _ => Factory.CreateSegment(drawing, first, second)
            };
            Actions.Add(drawing, figure);
            var withLength = figure as IFixableLength;
            if (kind == "fixedLength")
            {
                withLength.FixLength();
            }

            if (kind == "firstEnd")
            {
                second.Locked = true;
            }

            var panel = new DimensionsPanel(figure);
            var row = panel.GetProperties().SingleOrDefault(property => property.Name == nameof(DimensionsPanel.Direction));
            Require(row != null, kind + ": the panel has no direction.");
            Require((withLength != null) == panel.GetProperties().Any(property => property.Name == nameof(DimensionsPanel.Length)), kind + ": the panel's length row is wrong.");

            void Step(string what, Action action)
            {
                string before = drawing.SaveAsText();
                action();
                string after = drawing.SaveAsText();
                Require(after != before, kind + ": " + what + " changed nothing.");
                drawing.ActionManager.Undo();
                Require(drawing.SaveAsText() == before, kind + ": undo of " + what + " didn't restore the drawing.");
                drawing.ActionManager.Redo();
                Require(drawing.SaveAsText() == after, kind + ": redo of " + what + " changed the drawing.");
                drawing.Figures.CheckConsistency();
            }

            bool Offers(string verb) => panel.GetMethods().Any(method => method.Name == verb);

            // the distance between the two points, which a segment calls its length
            double Apart() => ((IPoint)figure.Dependencies[0]).Coordinates.Distance(((IPoint)figure.Dependencies[1]).Coordinates);

            Step("direction 30", () => Actions.SetProperty(drawing.ActionManager, row, value: 30.0));
            Near(figure.Direction, 30);
            Near(Apart(), 5);
            Require(Offers(nameof(DimensionsPanel.FixDirection)) && !Offers(nameof(DimensionsPanel.FreeDirection)), kind + ": Fix direction isn't offered.");
            Step("fix direction", panel.FixDirection);
            Require(!Offers(nameof(DimensionsPanel.FixDirection)) && Offers(nameof(DimensionsPanel.FreeDirection)), kind + ": Free direction isn't offered.");
            if (withLength != null)
            {
                Step("length 2", () => Set(drawing, figure, nameof(IFixableLength.Length), 2.0));
                Near(figure.Direction, 30);
                Near(withLength.Length, 2);
            }

            double apart = Apart();
            Step("horizontal", panel.MakeHorizontal);
            Near(figure.Direction, 0);
            Near(Apart(), apart);
            Require(!Offers(nameof(DimensionsPanel.MakeHorizontal)) && Offers(nameof(DimensionsPanel.MakeVertical)), kind + ": Horizontal is offered on a horizontal direction.");
            Step("vertical", panel.MakeVertical);
            Near(figure.Direction, 90);
            Step("direction 200", () => Actions.SetProperty(drawing.ActionManager, row, value: 200.0));
            Near(figure.Direction, 200);
            Step("horizontal from 200", panel.MakeHorizontal);
            Near(figure.Direction, 180);
            Step("free direction", panel.FreeDirection);
            Require(Offers(nameof(DimensionsPanel.FixDirection)), kind + ": Fix direction isn't offered again.");
            Near(figure.Direction, 180);
            Near(Apart(), apart);
        }

        // lines built otherwise point the way what they are built on says: a perpendicular
        // bisector (on two points, but across them), a parallel, an angle bisector
        {
            var drawing = NewDrawing();
            var a = AddPoint(drawing, x: 0, y: 0);
            var b = AddPoint(drawing, x: 3, y: 1);
            var c = AddPoint(drawing, x: 1, y: 4);
            var line = Factory.CreateLineTwoPoints(drawing, new IFigure[] { a, b });
            Actions.Add(drawing, line);
            var others = new IFixableDirection[]
            {
                Factory.CreateSegmentBisector(drawing, new IFigure[] { a, b }),
                Factory.CreateParallelLine(drawing, new IFigure[] { line, c }),
                Factory.CreateAngleBisector(drawing, new IFigure[] { a, b, c })
            };
            foreach (var other in others)
            {
                Actions.Add(drawing, other);
                Require(!other.CanEdit(nameof(IFixableDirection.Direction)) && !other.CanEdit(nameof(IFixableDirection.MakeHorizontal)), other.GetType().Name + " offers a direction.");
                Require(new DimensionsPanel(other).GetProperties().All(property => property.Name != nameof(DimensionsPanel.Direction)), other.GetType().Name + "'s panel has a direction.");
            }
        }

        // a line at an angle, in the panel after its tool: Horizontal and Vertical set its typed
        // angle the nearer way, and an angle taken from a measurement is typed again for them;
        // each one undo step
        {
            var drawing = NewDrawing();
            var a = AddPoint(drawing, x: 0, y: 0);
            var b = AddPoint(drawing, x: 3, y: 1);
            var c = AddPoint(drawing, x: 1, y: 4);
            var angle = Number.CreateAuxiliary(drawing, value: 200);
            Actions.Add(drawing, angle);
            var line = LineAtAngle.Create(drawing, c, angle);
            Actions.Add(drawing, line);
            var panel = new TiedValuesPanel(line, line);

            IOperationDescription Verb(string name) => panel.GetMethods().SingleOrDefault(method => method.Name == name);

            void Press(string name, double expected)
            {
                string before = drawing.SaveAsText();
                Verb(name).Invoke(panel, arguments: null);
                string after = drawing.SaveAsText();
                Near(line.Angle, expected);
                Require(line.AngleSource is Number, name + " left the angle tied.");
                Require(Verb(name) == null, name + " is offered on a line along its axis.");
                drawing.ActionManager.Undo();
                Require(drawing.SaveAsText() == before, "Undo of " + name + " on a line at an angle didn't restore the drawing.");
                drawing.ActionManager.Redo();
                Require(drawing.SaveAsText() == after, "Redo of " + name + " on a line at an angle changed the drawing.");
                drawing.Figures.CheckConsistency();
            }

            Press(nameof(LineAtAngle.MakeHorizontal), expected: 180);
            // (from 180°, 90° and 270° are as near: the axis itself)
            Press(nameof(LineAtAngle.MakeVertical), expected: 90);
            var measurement = Factory.CreateAngleMeasurement(drawing, new IFigure[] { a, b, c });
            Actions.Add(drawing, measurement);
            Require(line.TieTo(nameof(LineAtAngle.Angle), measurement), "The line didn't take the measurement's angle.");
            Press(nameof(LineAtAngle.MakeHorizontal), expected: 0);
        }

        // the end of a segment made horizontal, in its own grid: ticking Free direction (its
        // distance is free already) makes it a free point where it is, one undo step; so does
        // ticking Free distance on the end of a fixed length
        foreach (string quantity in new[] { nameof(TranslatedPoint.FreeDirection), nameof(TranslatedPoint.FreeDistance) })
        {
            var drawing = NewDrawing();
            var first = AddPoint(drawing, x: 1, y: 2);
            var segment = Factory.CreateSegment(drawing, first, AddPoint(drawing, x: 4, y: 3));
            Actions.Add(drawing, segment);
            if (quantity == nameof(TranslatedPoint.FreeDirection))
            {
                segment.MakeHorizontal();
            }
            else
            {
                segment.FixLength();
            }

            var end = (TranslatedPoint)segment.Dependencies[1];
            var place = end.Coordinates;
            var box = PropertyDiscoveryStrategy.CreateValueProvider(end, quantity);
            Require(box.CanSetValue, quantity + " can't be ticked on the end.");
            string before = drawing.SaveAsText();
            Actions.SetProperty(drawing.ActionManager, box, value: true);
            string after = drawing.SaveAsText();
            var free = segment.Dependencies[1];
            Require(PointSnapping.IsFreePoint(free) && free.Name == end.Name, quantity + " didn't make the end a free point.");
            NearPoint(((IPoint)free).Coordinates, place);
            drawing.ActionManager.Undo();
            Require(drawing.SaveAsText() == before && segment.Dependencies[1] == end, quantity + ": undo didn't bring the translated point back.");
            drawing.ActionManager.Redo();
            Require(drawing.SaveAsText() == after, quantity + ": redo changed the drawing.");
            drawing.Figures.CheckConsistency();
        }

        // a circle has a length and no direction
        var circleDrawing = NewDrawing();
        var circle = Factory.CreateCircle(circleDrawing, new IFigure[] { AddPoint(circleDrawing, x: 0, y: 0), AddPoint(circleDrawing, x: 1, y: 0) });
        Actions.Add(circleDrawing, circle);
        var circlePanel = new DimensionsPanel(circle);
        Require(circlePanel.GetProperties().All(property => property.Name != nameof(DimensionsPanel.Direction)), "A circle's panel has a direction.");
        Require(circlePanel.GetMethods().All(method => method.Name != nameof(DimensionsPanel.MakeHorizontal)), "A circle's panel has Horizontal.");
    }

    /// <summary>
    /// A vector's X and Y, the second point less the first: typed, they move the second point
    /// (the first when the second can't move), a fixed length and direction take the new
    /// values; each one undo step. In the panel after the tool, and not in a selection of
    /// several vectors.
    /// </summary>
    static void VectorComponentsEdit()
    {
        foreach (string kind in new[] { "free", "fixed", "firstEnd" })
        {
            var drawing = NewDrawing();
            var first = AddPoint(drawing, x: 1, y: 2);
            var second = AddPoint(drawing, x: 4, y: 6);
            var vector = Factory.CreateVector(drawing, new IFigure[] { first, second });
            Actions.Add(drawing, vector);
            if (kind == "fixed")
            {
                vector.FixLength();
                vector.FixDirection();
            }

            if (kind == "firstEnd")
            {
                second.Locked = true;
            }

            Near(vector.X, 3);
            Near(vector.Y, 4);
            var names = new DimensionsPanel(vector).GetProperties().Select(property => property.Name).ToList();
            Require(names.Contains(nameof(DynamicGeometry.Vector.X)) && names.Contains(nameof(DynamicGeometry.Vector.Y)), kind + ": the panel has no X and Y.");

            void Step(string property, double value, double expectedX, double expectedY)
            {
                string before = drawing.SaveAsText();
                var start = ((IPoint)vector.Dependencies[0]).Coordinates;
                Set(drawing, vector, property, value);
                string after = drawing.SaveAsText();
                Near(vector.X, expectedX);
                Near(vector.Y, expectedY);
                if (kind == "firstEnd")
                {
                    NearPoint(second.Coordinates, new Point(4, 6));
                }
                else
                {
                    NearPoint(((IPoint)vector.Dependencies[0]).Coordinates, start);
                }

                drawing.ActionManager.Undo();
                Require(drawing.SaveAsText() == before, kind + ": undo of " + property + " = " + value + " didn't restore the drawing.");
                drawing.ActionManager.Redo();
                Require(drawing.SaveAsText() == after, kind + ": redo of " + property + " = " + value + " changed the drawing.");
                drawing.Figures.CheckConsistency();
            }

            Step(nameof(DynamicGeometry.Vector.X), value: -2, expectedX: -2, expectedY: 4);
            Step(nameof(DynamicGeometry.Vector.Y), value: 0, expectedX: -2, expectedY: 0);
            Step(nameof(DynamicGeometry.Vector.X), value: 0, expectedX: 0, expectedY: 0);
            Step(nameof(DynamicGeometry.Vector.Y), value: -5, expectedX: 0, expectedY: -5);
            if (kind == "fixed")
            {
                Near(vector.Length, 5);
                Near(vector.Angle, 270);
                Require(vector.Dependencies[1] is TranslatedPoint end && !end.HasFreedom, "The fixed vector's end isn't fixed any more.");
            }
        }

        // the end as a point by coordinates at the start plus the components: it follows the
        // start, the vector keeps them; one undo step; "A.X - 2" and a bare "A.Y" for 0
        foreach (var (endX, endY, expectedX, expectedY) in new[] { (4.0, 6.5, "A.X + 3", "A.Y + 4.5"), (-1.0, 2.0, "A.X - 2", "A.Y") })
        {
            var drawing = NewDrawing();
            var start = AddPoint(drawing, x: 1, y: 2);
            start.Name = "A";
            var end = AddPoint(drawing, endX, endY);
            var vector = Factory.CreateVector(drawing, new IFigure[] { start, end });
            Actions.Add(drawing, vector);
            Require(vector.CanEdit(nameof(DynamicGeometry.Vector.ConvertEndToPointByCoordinates)), "A vector with a free end offers no conversion of it.");
            string before = drawing.SaveAsText();
            vector.ConvertEndToPointByCoordinates();
            string after = drawing.SaveAsText();
            var converted = vector.Dependencies[1] as PointByCoordinates;
            Require(converted != null && converted.Name == end.Name, "The end didn't become a point by coordinates.");
            Require(converted.XExpression.Text == expectedX && converted.YExpression.Text == expectedY, "The end's coordinates read " + converted.XExpression.Text + ", " + converted.YExpression.Text + ".");
            NearPoint(converted.Coordinates, new Point(endX, endY));
            Require(!vector.CanEdit(nameof(DynamicGeometry.Vector.ConvertEndToPointByCoordinates)), "The conversion is offered again.");
            drawing.ActionManager.Undo();
            Require(drawing.SaveAsText() == before && vector.Dependencies[1] == end, "Undo of the conversion didn't bring the free end back.");
            drawing.ActionManager.Redo();
            Require(drawing.SaveAsText() == after, "Redo of the conversion changed the drawing.");
            drawing.Figures.CheckConsistency();
            Set(drawing, start, nameof(FreePoint.X), 10.0);
            NearPoint(((IPoint)vector.Dependencies[1]).Coordinates, new Point(10 + endX - 1, endY));
        }

        // the vector of a midpoint and its end has no end to move; several vectors, no X or Y
        var other = NewDrawing();
        var a = AddPoint(other, x: 0, y: 0);
        var b = AddPoint(other, x: 2, y: 0);
        var middle = Factory.CreateMidPoint(other, new IFigure[] { a, b });
        Actions.Add(other, middle);
        var dependent = Factory.CreateVector(other, new IFigure[] { middle, b });
        var plain = Factory.CreateVector(other, new IFigure[] { a, AddPoint(other, x: 1, y: 3) });
        Actions.Add(other, dependent);
        Actions.Add(other, plain);
        Require(!dependent.CanEdit(nameof(DynamicGeometry.Vector.X)), "A vector from a midpoint of its end offers X to type.");
        Require(!dependent.CanEdit(nameof(DynamicGeometry.Vector.ConvertEndToPointByCoordinates)), "A vector from a midpoint of its end offers its end as a point by coordinates.");
        var selection = new FigureSelection(other, new IFigure[] { dependent, plain }, new IncludeByDefaultValueDiscoveryStrategy());
        Require(selection.GetProperties().All(property => property.Name != nameof(DynamicGeometry.Vector.X)), "Several vectors share an X row.");
    }

    static void PointFunctionsRejectCycles()
    {
        var drawing = NewDrawing();
        var first = AddPoint(drawing, x: 0, y: 0);
        var second = AddPoint(drawing, x: 3, y: 4);
        foreach (string text in new[] { "dist(A, B)", "ang(A, B, A)", "area(A, B, A)", "AB", "B.X" })
        {
            var result = Compiler.Instance.CompileExpression(drawing, text, figure => figure != second);
            Require(!result.IsSuccess && !string.IsNullOrWhiteSpace(result.GetErrorText()), text + " accepted a forbidden dependency.");
        }
    }

    /// <summary>
    /// Every way of making an expression's delegate (ExpressionStrategy) gives the same values,
    /// the same "undefined", and an error with words in it for the same wrong texts; the
    /// functions of x too. A, B and C are the 3-4-5 triangle with its right angle at C.
    /// </summary>
    static void ExpressionStrategiesAgree()
    {
        var drawing = NewDrawing();
        var a = AddPoint(drawing, x: 0, y: 0);
        var b = AddPoint(drawing, x: 3, y: 4);
        AddPoint(drawing, x: 3, y: 0);
        Actions.Add(drawing, Factory.CreateSegment(drawing, a, b));
        double pi = System.Math.PI;
        var values = new (string Text, double Expected)[]
        {
            ("AB", 5), ("A.X + B.Y", 4), ("dist(A, B)", 5), ("area(A, C, B)", 6), ("ang(A, C, B)", pi / 2),
            ("sqrt(16)", 4), ("2^3^2", 512), ("-2^2", -4), ("round(2.5)", 3), ("sign(-3)", -1), ("max(1, 2)", 2),
            ("atan2(1, 1)", pi / 4), ("pi", pi), ("e", System.Math.E), ("AB * 2 + 1", 11), ("B.X^2", 9), ("ln(e)", 1),
            ("lg(100)", 2), ("sqr(9)", 3), ("clamp(5, 0, 1)", 1), ("AB.Length", 5), ("3 / 4", 0.75), ("1 - 2 - 3", -4),
            ("2 * (3 + 4)", 14), ("deg(pi)", 180), ("abs(-2.5)", 2.5), ("floor(2.7) + ceiling(2.2)", 5)
        };
        var undefined = new[] { "clamp(5, 1, 0)", "sign(0 / 0)", "sqr(-1)", "sqrt(-4)" };
        var wrong = new[] { "foo(1)", "A.Name", "2x", "A.X +", "dist(1, 2)", "unknown", "AB(", "max(1)", "A.Visible", "ang(A, B)" };
        var functions = new (string Text, double X, double Expected)[]
        {
            ("sin(x) + A.X", 0, 0), ("x^2 + B.X", 2, 7), ("AB * x", 3, 15), ("-x^2", 3, -9), ("round(x)", 2.5, 3)
        };
        var previous = Compiler.Instance.Strategy;
        try
        {
            foreach (ExpressionStrategy strategy in Enum.GetValues(typeof(ExpressionStrategy)))
            {
                Compiler.Instance.Strategy = strategy;
                foreach (var (text, expected) in values)
                {
                    var result = Compiler.Instance.CompileExpression(drawing, text, isFigureAllowed: null);
                    Require(result.IsSuccess, strategy + " did not compile " + text + ": " + result.GetErrorText());
                    double actual = result.Expression();
                    Require(System.Math.Abs(actual - expected) < 1e-9, strategy + " gave " + actual + " for " + text + ", expected " + expected);
                }

                foreach (var text in undefined)
                {
                    var result = Compiler.Instance.CompileExpression(drawing, text, isFigureAllowed: null);
                    Require(result.IsSuccess && double.IsNaN(result.Expression()), strategy + " did not give undefined for " + text);
                }

                foreach (var text in wrong)
                {
                    var result = Compiler.Instance.CompileExpression(drawing, text, isFigureAllowed: null);
                    Require(!result.IsSuccess && !string.IsNullOrWhiteSpace(result.GetErrorText()), strategy + " accepted " + text);
                }

                foreach (var (text, x, expected) in functions)
                {
                    var result = Compiler.Instance.CompileFunction(drawing, text);
                    Require(result.IsSuccess, strategy + " did not compile the function " + text + ": " + result.GetErrorText());
                    double actual = result.Function(x);
                    Require(System.Math.Abs(actual - expected) < 1e-9, strategy + " gave " + actual + " for " + text + " at " + x + ", expected " + expected);
                }
            }
        }
        finally
        {
            Compiler.Instance.Strategy = previous;
        }
    }

    static void LabelsRejectCycles()
    {
        var drawing = NewDrawing();
        var label = Factory.CreateLabel(drawing);
        Actions.Add(drawing, label);
        label.Text = "[1]";
        label.Text = "[" + label.Name + ".Value]";
        Require(!label.Dependencies.Contains(label), "A label depends on itself.");
        Require(!label.IsNumber, "A cyclic label was evaluated.");
        label.Text = "[2]";
        var dependent = Factory.CreatePointByCoordinates(drawing, label.Name + ".Value", "0");
        Actions.Add(drawing, dependent);
        label.Text = "[" + dependent.Name + ".X]";
        Require(!label.Dependencies.Contains(dependent), "A label depends on its descendant.");
        drawing.Figures.CheckConsistency();
    }

    static void ClearLabelText()
    {
        var drawing = NewDrawing();
        var label = Factory.CreateLabel(drawing);
        Actions.Add(drawing, label);
        label.Text = "Words";
        Set(drawing, label, nameof(DynamicGeometry.Label.Text), value: "");
        Require(string.IsNullOrEmpty(label.ProcessedText), "The old text is still displayed.");
        drawing.ActionManager.Undo();
        Require(label.ProcessedText == "Words", "Undo did not restore the text.");
        drawing.ActionManager.Redo();
        Require(string.IsNullOrEmpty(label.ProcessedText), "Redo did not clear the text.");
    }

    static void FunctionsRejectCycles()
    {
        var drawing = NewDrawing();
        var graph = new FunctionGraph { Drawing = drawing, FunctionText = "x" };
        Actions.Add(drawing, graph);
        var point = Factory.CreatePointOnFigure(drawing, graph, parameter: 2);
        Actions.Add(drawing, point);
        string expression = point.Name + ".Y + x";
        var compiled = Compiler.Instance.CompileFunction(drawing, expression, figure => !figure.DependsOn(graph));
        Require(!compiled.IsSuccess, "A graph accepted a point on itself.");
        graph.FunctionText = expression;
        Require(!graph.Dependencies.Contains(point), "A graph installed a dependency cycle.");
        drawing.Figures.CheckConsistency();
    }

    static void FunctionReplacement()
    {
        var drawing = NewDrawing();
        var point = AddPoint(drawing, x: 3, y: 4);
        var graph = new FunctionGraph { Drawing = drawing, FunctionText = "A.X + A.Y + x" };
        Actions.Add(drawing, graph);
        string before = drawing.SaveAsText();
        PointSnapping.ConvertToPointByCoordinates(point);
        drawing.Figures.CheckConsistency();
        Near(graph.Function(0), expected: 7);
        drawing.ActionManager.Undo();
        drawing.Figures.CheckConsistency();
        Require(drawing.SaveAsText() == before, "Undo of replacing a repeated function dependency changed the drawing.");
    }

    static void ConstructionUndo()
    {
        var creators = new (FigureCreator Creator, int Points)[]
        {
            (new SegmentCreator(), 2), (new RayCreator(), 2), (new LineTwoPointsCreator(), 2),
            (new VectorCreator(), 2), (new CircleCreator(), 2), (new MidpointCreator(), 2),
            (new TriangleCreator(), 3), (new RegularPolygonCreator(), 2), (new RegularPolygonOnSideCreator(), 2), (new SquareCreator(), 2),
            (new CircleArcCreator(), 3), (new EllipseCreator(), 3),
            (new AngleMeasurementCreator(), 3), (new AngleBisectorCreator(), 3),
            (new SegmentBisectorCreator(), 2), (new BezierCreator(), 4),
            (new PolygonCreator(), 3), (new AreaMeasurementCreator(), 3)
        };
        var positions = new[] { new Point(0, 0), new Point(3, 0), new Point(0, 4), new Point(-2, 1) };
        foreach (var (creator, pointCount) in creators)
        {
            var drawing = NewDrawing();
            drawing.Behavior = creator;
            string before = drawing.SaveAsText();
            new FigureCreator.Dialog(creator) { X = "0", Y = "0" }.AddPoint();
            creator.Restart();
            Require(!drawing.IsRecordingTransaction, creator.Name + " left a transaction open.");
            Require(drawing.SaveAsText() == before, creator.Name + " cancellation left figures.");
            for (int index = 0; index < pointCount; index++)
            {
                var position = positions[index];
                new FigureCreator.Dialog(creator) { X = position.X.ToString(), Y = position.Y.ToString() }.AddPoint();
            }

            if (creator is PolygonCreator or AreaMeasurementCreator)
            {
                creator.KeyDown(drawing.Canvas, new KeyEventArgs { Key = Key.Enter });
            }

            Require(!drawing.IsRecordingTransaction, creator.Name + " did not complete.");
            string after = drawing.SaveAsText();
            Require(after != before, creator.Name + " made nothing.");
            drawing.ActionManager.Undo();
            Require(drawing.SaveAsText() == before, creator.Name + " undo.");
            drawing.ActionManager.Redo();
            Require(drawing.SaveAsText() == after, creator.Name + " redo.");
            drawing.Figures.CheckConsistency();
        }
    }

    static void DraggingUndo()
    {
        foreach (bool snap in new[] { false, true })
        {
            var drawing = NewDrawing();
            var point = AddPoint(drawing, x: -3, y: 1);
            var first = AddPoint(drawing, x: 0, y: -3);
            var second = AddPoint(drawing, x: 0, y: 3);
            var segment = Factory.CreateSegment(drawing, first, second);
            Actions.Add(drawing, segment);
            using var window = new TestWindow(drawing.Canvas);
            drawing.Behavior = new Dragger();
            string before = drawing.SaveAsText();
            var from = drawing.CoordinateSystem.ToPhysical(point.Coordinates);
            var to = drawing.CoordinateSystem.ToPhysical(new Point(snap ? 0 : -1, 1));
            using var pointer = new Pointer(Pointer.GetNextFreeId(), PointerType.Mouse, isPrimary: true);
            var modifiers = snap ? KeyModifiers.Alt : KeyModifiers.None;
            var down = new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.LeftButtonPressed);
            var moving = new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.Other);
            var up = new PointerPointProperties(RawInputModifiers.None, PointerUpdateKind.LeftButtonReleased);
            var canvas = drawing.Canvas;
            canvas.RaiseEvent(new PointerPressedEventArgs(
                canvas,
                pointer,
                canvas,
                from,
                timestamp: 0,
                down,
                modifiers,
                clickCount: 1));
            for (int step = 1; step <= 12; step++)
            {
                canvas.RaiseEvent(new PointerEventArgs(
                    InputElement.PointerMovedEvent,
                    canvas,
                    pointer,
                    canvas,
                    from + (to - from) * (step / 12.0),
                    timestamp: (ulong)step,
                    moving,
                    modifiers));
            }

            canvas.RaiseEvent(new PointerReleasedEventArgs(
                canvas,
                pointer,
                canvas,
                to,
                timestamp: 13,
                up,
                modifiers,
                MouseButton.Left));
            Require(!drawing.IsRecordingTransaction, "Drag left an open transaction.");
            if (snap)
            {
                Require(Find(drawing, point.Name) is PointOnFigure, "Alt drag did not snap onto the segment.");
            }
            else
            {
                Near(point.X, expected: -1);
                Near(point.Y, expected: 1);
            }

            string after = drawing.SaveAsText();
            drawing.ActionManager.Undo();
            Require(drawing.SaveAsText() == before, "Drag undo changed the drawing.");
            drawing.ActionManager.Redo();
            Require(drawing.SaveAsText() == after, "Drag redo changed the drawing.");
            drawing.Figures.CheckConsistency();
        }
    }

    /// <summary>
    /// Triangle ABC with a point inside it from the Point inside tool: dragged out, the
    /// point stops at the side; the triangle dragged carries it along; a vertex dragged moves
    /// the center and the point with it; the triangle flattened past the point pushes it
    /// back in, and undo gives it its place back; a circle without a fill takes one too; Tab
    /// chooses the outer of two nested triangles; the file round trips.
    /// </summary>
    static void PointInsideShape()
    {
        var drawing = NewDrawing();
        using var window = new TestWindow(drawing.Canvas);
        var system = drawing.CoordinateSystem;
        string status = null;
        drawing.ChoiceStatus += text => status = text;
        var a = AddPoint(drawing, x: 0, y: 0);
        var b = AddPoint(drawing, x: 6, y: 0);
        var c = AddPoint(drawing, x: 0, y: 6);
        var triangle = Factory.CreatePolygon(drawing, new IFigure[] { a, b, c });
        Actions.Add(drawing, triangle);

        // the tool: a click inside makes the point, a click inside nothing makes nothing
        drawing.Behavior = new PointInRegionCreator();
        Click(drawing, system.ToPhysical(new Point(1, 1)));
        var point = drawing.Figures.OfType<PointInRegion>().SingleOrDefault();
        Require(point != null && point.Dependencies[0] == triangle, "The click inside the triangle made no point inside it.");
        Near(point.X, expected: 1);
        Near(point.Y, expected: 1);
        Require(point.Style?.Name == StyleManager.PointOnFigureStyleName, "The point is not green.");
        Click(drawing, system.ToPhysical(new Point(5, 5)));
        Require(drawing.Figures.OfType<PointInRegion>().Count() == 1, "A click outside every shape made a point.");

        // dragged out, it stops at the side x + y = 6
        drawing.Behavior = new Dragger();
        string before = drawing.SaveAsText();
        Drag(drawing, system.ToPhysical(point.Coordinates), system.ToPhysical(new Point(5, 5)));
        Near(point.X, expected: 3);
        Near(point.Y, expected: 3);
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the drag did not put the point back.");

        // the triangle dragged carries the point
        Drag(drawing, system.ToPhysical(new Point(2, 2)), system.ToPhysical(new Point(4, 3)));
        Near(a.X, expected: 2);
        Near(point.X, expected: 3);
        Near(point.Y, expected: 2);
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the triangle's drag did not put the point back.");

        // a vertex dragged moves the center by a third, and the point with it
        static void MovePoint(Drawing drawing, IPoint moved, Point offset)
        {
            var toRecalculate = DependencyAlgorithms.FindDescendants(f => f.Dependents, new IFigure[] { moved });
            toRecalculate.Reverse();
            Actions.Move(drawing, new List<IMovable> { (IMovable)moved }, offset, toRecalculate);
        }

        MovePoint(drawing, c, new Point(0, -3));
        Near(point.X, expected: 1);
        Near(point.Y, expected: 0);
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the vertex's drag did not put the point back.");

        // flattened past the point, it is pushed back in; undo gives its place back
        MovePoint(drawing, c, new Point(0, -5.5));
        Require(point.Exists && point.Y <= 0.5 + 1e-9, "The point was not pushed back into the flattened triangle: " + point.Coordinates);
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the flattening did not put the point back.");

        // a circle without a fill takes a point, and the point stops at the rim
        var center = AddPoint(drawing, x: 7, y: 6);
        var circle = Factory.CreateCircle(drawing, new IFigure[] { center, AddPoint(drawing, x: 8.5, y: 6) });
        Actions.Add(drawing, circle);
        drawing.Behavior = new PointInRegionCreator();
        Click(drawing, system.ToPhysical(new Point(7.5, 6.5)));
        var inCircle = drawing.Figures.OfType<PointInRegion>().SingleOrDefault(p => p.Dependencies[0] == circle);
        Require(inCircle != null, "The click inside the circle made no point inside it.");
        MovePoint(drawing, inCircle, new Point(5, 0));
        Near(inCircle.Coordinates.Distance(center.Coordinates), expected: 1.5);
        Require(inCircle.Y > 6, "The point did not stop on the rim towards where it was dragged: " + inCircle.Coordinates);

        // nested: the inner triangle first, Tab for the outer
        var inner = Factory.CreatePolygon(drawing, new IFigure[] { AddPoint(drawing, x: 0.5, y: 0.5), AddPoint(drawing, x: 2.5, y: 0.5), AddPoint(drawing, x: 0.5, y: 2.5) });
        Actions.Add(drawing, inner);
        var tool = new PointInRegionCreator();
        drawing.Behavior = tool;
        var at = system.ToPhysical(new Point(0.8, 0.8));
        Hover(drawing, at);
        Require(status != null && status.StartsWith("A point inside ") && status.Contains("(1 of 2)"), "The status inside two triangles: " + status);
        Require(tool.StepChoice(backwards: false), "Tab chose nothing.");
        Click(drawing, at);
        var chosen = drawing.Figures.OfType<PointInRegion>().Single(p => p.Coordinates.Distance(new Point(0.8, 0.8)) < 1e-9);
        Require(chosen.Dependencies[0] == triangle, "Tab did not take the outer triangle.");

        // the file round trips, offset and all
        var saved = drawing.SaveAsText();
        var loaded = ReadLgf(saved);
        Require(loaded.LoadErrors == null, "Reload: " + loaded.LoadErrors);
        Require(loaded.SaveAsText() == saved, "The file changed on a round trip.");
        var reloaded = loaded.Figures.OfType<PointInRegion>().Single(p => p.Name == point.Name);
        Near(reloaded.X, expected: 1);
        Near(reloaded.Y, expected: 1);
        drawing.Figures.CheckConsistency();
    }

    /// <summary>
    /// Points A B C D, segments AB BC CD and polygon ABCD: C dragged with Alt onto B joins it
    /// (it slid onto segment AB beside B instead), segment BC goes, CD becomes BD, the
    /// polygon ABD
    /// </summary>
    static void AltDragJoinsNeighbor()
    {
        var drawing = NewDrawing();
        var a = AddPoint(drawing, x: -3, y: 0);
        var b = AddPoint(drawing, x: 0, y: 0);
        var c = AddPoint(drawing, x: 2, y: 2);
        var d = AddPoint(drawing, x: 4, y: 0);
        var ab = Factory.CreateSegment(drawing, a, b);
        var bc = Factory.CreateSegment(drawing, b, c);
        var cd = Factory.CreateSegment(drawing, c, d);
        var polygon = Factory.CreatePolygon(drawing, new IFigure[] { a, b, c, d });
        Actions.Add(drawing, polygon);
        Actions.Add(drawing, ab);
        Actions.Add(drawing, bc);
        Actions.Add(drawing, cd);
        using var window = new TestWindow(drawing.Canvas);
        drawing.Behavior = new Dragger();
        string before = drawing.SaveAsText();
        var from = drawing.CoordinateSystem.ToPhysical(c.Coordinates);
        var to = drawing.CoordinateSystem.ToPhysical(b.Coordinates);
        using var pointer = new Pointer(Pointer.GetNextFreeId(), PointerType.Mouse, isPrimary: true);
        var down = new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.LeftButtonPressed);
        var moving = new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.Other);
        var up = new PointerPointProperties(RawInputModifiers.None, PointerUpdateKind.LeftButtonReleased);
        var canvas = drawing.Canvas;
        canvas.RaiseEvent(new PointerPressedEventArgs(
            canvas,
            pointer,
            canvas,
            from,
            timestamp: 0,
            down,
            KeyModifiers.Alt,
            clickCount: 1));
        for (int step = 1; step <= 12; step++)
        {
            canvas.RaiseEvent(new PointerEventArgs(
                InputElement.PointerMovedEvent,
                canvas,
                pointer,
                canvas,
                from + (to - from) * (step / 12.0),
                timestamp: (ulong)step,
                moving,
                KeyModifiers.Alt));
        }

        canvas.RaiseEvent(new PointerReleasedEventArgs(
            canvas,
            pointer,
            canvas,
            to,
            timestamp: 13,
            up,
            KeyModifiers.Alt,
            MouseButton.Left));
        Require(!drawing.IsRecordingTransaction, "Drag left an open transaction.");
        Require(!drawing.Figures.Contains(c), "C was not joined into B.");
        Require(!drawing.Figures.Contains(bc), "Segment BC is still there.");
        Require(ab.Dependencies.SequenceEqual(new IFigure[] { a, b }), "Segment AB changed.");
        Require(cd.Dependencies.SequenceEqual(new IFigure[] { b, d }), "Segment CD is not BD: " + string.Join(", ", cd.Dependencies));
        Require(polygon.Dependencies.SequenceEqual(new IFigure[] { a, b, d }), "The polygon is not ABD: " + string.Join(", ", polygon.Dependencies));
        drawing.Figures.CheckConsistency();
        string after = drawing.SaveAsText();
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the join changed the drawing.");
        drawing.Figures.CheckConsistency();
        drawing.ActionManager.Redo();
        Require(drawing.SaveAsText() == after, "Redo of the join changed the drawing.");
        drawing.Figures.CheckConsistency();
    }

    /// <summary>
    /// A finger on the canvas: pressed where it is made, moved, lifted - touch pointer
    /// events raised by hand, the pointer captured as a real touch is
    /// </summary>
    class Finger : IDisposable
    {
        static readonly PointerPointProperties down = new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.LeftButtonPressed);
        static readonly PointerPointProperties held = new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.Other);
        static readonly PointerPointProperties up = new PointerPointProperties(RawInputModifiers.None, PointerUpdateKind.LeftButtonReleased);
        static ulong time;

        readonly Canvas canvas;
        readonly Pointer pointer;

        public Point Position { get; private set; }

        public Finger(Canvas canvas, Point at)
        {
            this.canvas = canvas;
            Position = at;
            pointer = new Pointer(Pointer.GetNextFreeId(), PointerType.Touch, isPrimary: false);
            pointer.Capture(canvas);
            canvas.RaiseEvent(new PointerPressedEventArgs(canvas, pointer, canvas, at, time += 20, down, KeyModifiers.None, clickCount: 1));
        }

        public void MoveTo(Point to, int steps = 6)
        {
            var from = Position;
            for (int step = 1; step <= steps; step++)
            {
                Position = from + (to - from) * (step / (double)steps);
                canvas.RaiseEvent(new PointerEventArgs(InputElement.PointerMovedEvent, canvas, pointer, canvas, Position, time += 10, held, KeyModifiers.None));
            }
        }

        public void Lift()
        {
            canvas.RaiseEvent(new PointerReleasedEventArgs(canvas, pointer, canvas, Position, time += 20, up, KeyModifiers.None, MouseButton.Left));
            pointer.Capture(null);
        }

        public void Dispose()
        {
            pointer.Dispose();
        }
    }

    /// <summary>Points A and B with a segment: a finger on each, moved at once, is two drags and one undo step</summary>
    static void TwoFingersDragTwoPoints()
    {
        var drawing = NewDrawing();
        var a = AddPoint(drawing, x: -3, y: 0);
        var b = AddPoint(drawing, x: 3, y: 0);
        Actions.Add(drawing, Factory.CreateSegment(drawing, a, b));
        using var window = new TestWindow(drawing.Canvas);
        drawing.Behavior = new Dragger();
        string before = drawing.SaveAsText();
        var system = drawing.CoordinateSystem;
        using var first = new Finger(drawing.Canvas, system.ToPhysical(a.Coordinates));
        using var second = new Finger(drawing.Canvas, system.ToPhysical(b.Coordinates));
        var firstFrom = first.Position;
        var secondFrom = second.Position;
        var firstTo = system.ToPhysical(new Point(-3, 2));
        var secondTo = system.ToPhysical(new Point(3, -2));
        for (int step = 1; step <= 6; step++)
        {
            first.MoveTo(firstFrom + (firstTo - firstFrom) * (step / 6.0), steps: 1);
            second.MoveTo(secondFrom + (secondTo - secondFrom) * (step / 6.0), steps: 1);
        }

        Near(a.X, expected: -3);
        Near(a.Y, expected: 2);
        Near(b.X, expected: 3);
        Near(b.Y, expected: -2);
        Require(drawing.IsRecordingTransaction, "Two point drags have no undo step open.");
        first.Lift();
        Require(drawing.IsRecordingTransaction, "Lifting one finger ended the step while the other still drags.");
        second.Lift();
        Require(!drawing.IsRecordingTransaction, "The two-finger drag left an open transaction.");
        string after = drawing.SaveAsText();
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the two-finger drag changed the drawing.");
        drawing.ActionManager.Redo();
        Require(drawing.SaveAsText() == after, "Redo of the two-finger drag changed the drawing.");
        drawing.Figures.CheckConsistency();
    }

    /// <summary>
    /// A finger drags a point and holds it; another finger on the paper pans, two more zoom:
    /// the view follows them, the point stays under its finger, and the gesture is one undo
    /// step (the pans are no steps)
    /// </summary>
    static void FingerHoldsPointWhilePaperMoves()
    {
        var drawing = NewDrawing();
        var a = AddPoint(drawing, x: 0, y: 0);
        using var window = new TestWindow(drawing.Canvas);
        drawing.Behavior = new Dragger();
        var system = drawing.CoordinateSystem;
        var canvas = drawing.Canvas;
        using var held = new Finger(canvas, system.ToPhysical(a.Coordinates));
        held.MoveTo(system.ToPhysical(new Point(1, 1)));
        Near(a.X, expected: 1);
        Near(a.Y, expected: 1);
        var originBefore = system.Origin;
        using var paper = new Finger(canvas, new Point(100, 100));
        paper.MoveTo(new Point(250, 180));
        Near(system.Origin.X - originBefore.X, expected: 150);
        Near(system.Origin.Y - originBefore.Y, expected: 80);
        Require(system.ToPhysical(a.Coordinates).Distance(held.Position) < 1, "The held point drifted off its finger while the paper was panned.");
        paper.Lift();
        double unitBefore = system.UnitLength;
        using var left = new Finger(canvas, new Point(150, 600));
        using var right = new Finger(canvas, new Point(350, 600));
        left.MoveTo(new Point(100, 600));
        right.MoveTo(new Point(400, 600));
        Near(system.UnitLength / unitBefore, expected: 1.5);
        Require(system.ToPhysical(a.Coordinates).Distance(held.Position) < 1, "The held point drifted off its finger while the paper was zoomed.");
        left.Lift();
        right.Lift();
        Require(drawing.IsRecordingTransaction, "The paper fingers ended the point drag's step.");
        held.Lift();
        Require(!drawing.IsRecordingTransaction, "The drag left an open transaction.");
        var dragged = a.Coordinates;
        drawing.ActionManager.Undo();
        Near(a.X, expected: 0);
        Near(a.Y, expected: 0);
        drawing.ActionManager.Redo();
        Near(a.X, expected: dragged.X);
        Near(a.Y, expected: dragged.Y);
        drawing.Figures.CheckConsistency();
    }

    /// <summary>Under a tool that constructs, two fingers zoom and make nothing</summary>
    static void ConstructionToolPinches()
    {
        var drawing = NewDrawing();
        using var window = new TestWindow(drawing.Canvas);
        drawing.Behavior = new SegmentCreator();
        var system = drawing.CoordinateSystem;
        int count = drawing.Figures.Count();
        double unitBefore = system.UnitLength;
        using var left = new Finger(drawing.Canvas, new Point(400, 350));
        using var right = new Finger(drawing.Canvas, new Point(600, 350));
        left.MoveTo(new Point(350, 350));
        right.MoveTo(new Point(650, 350));
        left.Lift();
        right.Lift();
        Near(system.UnitLength / unitBefore, expected: 1.5);
        Require(drawing.Figures.Count() == count, "A pinch under the Segment tool made figures.");
        Require(!drawing.IsRecordingTransaction, "A pinch under the Segment tool left a construction open.");
        Require(!drawing.ActionManager.CanUndo, "A pinch under the Segment tool left an undo step.");
    }

    /// <summary>A finger put down on a point and lifted there selects it, as a click does</summary>
    static void TapSelectsFigure()
    {
        var drawing = NewDrawing();
        var a = AddPoint(drawing, x: 0, y: 0);
        using var window = new TestWindow(drawing.Canvas);
        drawing.Behavior = new Dragger();
        string before = drawing.SaveAsText();
        using var finger = new Finger(drawing.Canvas, drawing.CoordinateSystem.ToPhysical(a.Coordinates));
        finger.Lift();
        Require(a.Selected, "A tap on the point did not select it.");
        Require(!drawing.IsRecordingTransaction, "A tap left an open transaction.");
        Require(drawing.SaveAsText() == before, "A tap changed the drawing.");
    }

    /// <summary>
    /// A finger on a function graph, which a drag of pans the view, is a paper finger: with
    /// a second finger it zooms, the paper under the two stays under them, and nothing is
    /// recorded
    /// </summary>
    static void FingerOnGraphIsPaper()
    {
        var drawing = NewDrawing();
        var graph = new FunctionGraph() { Drawing = drawing };
        graph.FunctionText = "0";
        Actions.Add(drawing, graph);
        using var window = new TestWindow(drawing.Canvas);
        drawing.Behavior = new Dragger();
        var system = drawing.CoordinateSystem;
        var onGraph = system.ToPhysical(new Point(2, 0));
        var onPaper = system.ToPhysical(new Point(-2, -3));
        var middle = DynamicGeometry.Math.Midpoint(onGraph, onPaper);
        var under = system.ToLogical(middle);
        double unitBefore = system.UnitLength;
        double spanBefore = onGraph.Distance(onPaper);
        using var first = new Finger(drawing.Canvas, onGraph);
        using var second = new Finger(drawing.Canvas, onPaper);
        var firstTo = onGraph + new Avalonia.Vector(50, 30);
        var secondTo = onPaper - new Avalonia.Vector(50, 30);
        first.MoveTo(firstTo);
        second.MoveTo(secondTo);
        first.Lift();
        second.Lift();
        Near(system.UnitLength / unitBefore, expected: firstTo.Distance(secondTo) / spanBefore);
        Require(system.ToPhysical(under).Distance(middle) < 1, "The paper under the two fingers did not stay under them.");
        Require(!drawing.IsRecordingTransaction, "The pinch left an open transaction.");
        drawing.ActionManager.Undo();
        Require(!drawing.Figures.Contains(graph), "The pinch left an undo step before the graph's.");
    }

    public class TestWindow : Window, IDisposable
    {
        public TestWindow(Canvas canvas)
        {
            Content = canvas;
            Width = canvas.Width;
            Height = canvas.Height;
            Show();
            Dispatcher.UIThread.RunJobs();
        }

        public void Dispose()
        {
            Close();
        }
    }

    static void EditorUndo()
    {
        var drawing = NewDrawing();
        var first = AddPoint(drawing, x: 1, y: 2);
        var second = AddPoint(drawing, x: 4, y: 6);
        var segment = Factory.CreateSegment(drawing, first, second);
        Actions.Add(drawing, segment);
        var editor = new UpDownEditor
        {
            ActionManager = drawing.ActionManager,
            Value = PropertyDiscoveryStrategy.CreateValueProvider(segment, nameof(Segment.Length))
        };
        Dispatcher.UIThread.RunJobs();
        string before = drawing.SaveAsText();
        editor.TextBox.Text = "0";
        editor.TextBox.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter });
        Dispatcher.UIThread.RunJobs();
        Near(segment.Length, expected: 0);
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Editor undo lost the endpoint.");
    }

    static void PointParameterSurvivesUndo()
    {
        var drawing = NewDrawing();
        var center = AddPoint(drawing, x: 0, y: 0);
        var rim = AddPoint(drawing, x: 3, y: 0);
        var circle = Factory.CreateCircle(drawing, new IFigure[] { center, rim });
        Actions.Add(drawing, circle);
        var point = Factory.CreatePointOnFigure(drawing, circle, parameter: 7);
        Actions.Add(drawing, point);
        string before = drawing.SaveAsText();
        var moving = new List<IMovable> { point };
        Actions.Move(drawing, moving, new Point(0.2, 0.4), new IFigure[] { point });
        drawing.ActionManager.Undo();
        Require(point.Parameter == 7, "Parameter changed to " + point.Parameter);
        Require(drawing.SaveAsText() == before, "Undo changed the file.");
    }

    static Drawing ReadGeoGebra(string construction)
    {
        var (drawing, reader) = ReadGeoGebraWithReport(construction);
        Require(reader.IsSuccess, reader.Details);
        return drawing;
    }

    static (Drawing Drawing, GeoGebraReader Reader) ReadGeoGebraWithReport(string construction)
    {
        var drawing = NewDrawing();
        var reader = new GeoGebraReader();
        using var bytes = new MemoryStream();
        using (var archive = new ZipArchive(bytes, ZipArchiveMode.Create, leaveOpen: true))
        {
            using var writer = new StreamWriter(archive.CreateEntry(GeoGebraReader.WorksheetEntry).Open(), Encoding.UTF8);
            writer.Write("<geogebra><construction>" + construction + "</construction></geogebra>");
        }

        var worksheet = GeoGebraReader.ReadWorksheet(bytes.ToArray(), out string problem);
        Require(worksheet != null && problem == null, problem ?? "No worksheet was read.");
        reader.ReadDrawing(drawing, worksheet);
        drawing.Figures.CheckConsistency();
        return (drawing, reader);
    }

    static IFigure Find(Drawing drawing, string name)
    {
        return drawing.Figures.Single(figure => figure.Name == name);
    }

    static void GeoGebraDependentNumeric()
    {
        var drawing = ReadGeoGebra("""
            <element type="numeric" label="a"><value val="2"/></element>
            <expression label="b" exp="2*a"/>
            <element type="numeric" label="b"><value val="4"/></element>
            <element type="point" label="O"><coords x="0" y="0" z="1"/></element>
            <command name="Circle"><input a0="O" a1="b"/><output a0="c"/></command>
            <element type="conic" label="c"/>
            """);
        ((Number)Find(drawing, "a")).Value = 3;
        drawing.Recalculate();
        Near(((ICircle)Find(drawing, "c")).Radius, expected: 6);
    }

    static void GeoGebraUntranslatedNumeric()
    {
        var (drawing, reader) = ReadGeoGebraWithReport("""
            <element type="numeric" label="a"><value val="2"/></element>
            <expression label="b" exp="If(a > 1, 5, 6)"/>
            <element type="numeric" label="b"><value val="5"/></element>
            <element type="point" label="O"><coords x="0" y="0" z="1"/></element>
            <command name="Circle"><input a0="O" a1="b"/><output a0="c"/></command>
            <element type="conic" label="c"/>
            """);
        Require(!reader.IsSuccess, "A number kept at its value was not reported.");
        Near(((ICircle)Find(drawing, "c")).Radius, expected: 5);
    }

    static void GeoGebraShownNumericStaysHidden()
    {
        var drawing = ReadGeoGebra("""
            <element type="numeric" label="a"><value val="2"/></element>
            <expression label="b" exp="2*a"/>
            <element type="numeric" label="b"><value val="4"/><show object="true" label="true"/></element>
            """);
        Require(!Find(drawing, "b").Visible, "A worked-out number showed as a label.");
    }

    static void GeoGebraNumericRotation()
    {
        var drawing = ReadGeoGebra("""
            <element type="point" label="P"><coords x="1" y="0" z="1"/></element>
            <element type="point" label="O"><coords x="0" y="0" z="1"/></element>
            <element type="numeric" label="a"><value val="1.5707963267948966"/></element>
            <command name="Rotate"><input a0="P" a1="a" a2="O"/><output a0="Q"/></command>
            <element type="point" label="Q"><coords x="0" y="1" z="1"/></element>
            """);
        var point = (IPoint)Find(drawing, "Q");
        Near(point.Coordinates.X, expected: 0);
        Near(point.Coordinates.Y, expected: 1);
    }

    static void GeoGebraExpressionsRoundTrip()
    {
        var drawing = ReadGeoGebra("""
            <element type="numeric" label="a"><value val="2"/></element>
            <expression label="b" exp="2*a"/>
            <element type="numeric" label="b"><value val="4"/></element>
            <expression label="d" exp="b+1"/>
            <element type="numeric" label="d"><value val="5"/></element>
            <element type="point" label="P"><coords x="1" y="0" z="1"/></element>
            <element type="point" label="O"><coords x="0" y="0" z="1"/></element>
            <command name="Circle"><input a0="O" a1="d"/><output a0="c"/></command>
            <element type="conic" label="c"/>
            <command name="Rotate"><input a0="P" a1="a" a2="O"/><output a0="Q"/></command>
            <element type="point" label="Q"/>
            <expression label="alpha" exp="a*45°"/>
            <element type="angle" label="alpha"><value val="1.5707963267948966"/></element>
            <command name="Rotate"><input a0="P" a1="alpha" a2="O"/><output a0="R"/></command>
            <element type="point" label="R"/>
            <expression label="undefinedValue" exp="sqrt(a-3)"/>
            <element type="numeric" label="undefinedValue"><value val="NaN"/></element>
            <expression label="laterValue" exp="undefinedValue+1"/>
            <element type="numeric" label="laterValue"><value val="NaN"/></element>
            <expression label="angleText" exp="alpha"/>
            <element type="text" label="angleText"/>
            """);
        foreach (var current in new[] { drawing, ReadLgf(drawing.SaveAsText()) })
        {
            var number = (Number)Find(current, "a");
            number.Value = 3;
            current.Recalculate();
            Near(((ICircle)Find(current, "c")).Radius, expected: 7);
            Near(((IPoint)Find(current, "Q")).Coordinates.X, System.Math.Cos(3));
            Near(((IPoint)Find(current, "R")).Coordinates.X, System.Math.Cos(3 * System.Math.PI / 4));
            Near(((DynamicGeometry.Label)Find(current, "laterValue")).Value, expected: 1);
            Near(((DynamicGeometry.Label)Find(current, "angleText")).Value, expected: 135);
            current.Figures.CheckConsistency();
        }
    }

    static void GeoGebraCentroid()
    {
        var drawing = ReadGeoGebra("""
            <element type="point" label="A"><coords x="0" y="0" z="1"/></element>
            <element type="point" label="B"><coords x="4" y="0" z="1"/></element>
            <element type="point" label="C"><coords x="4" y="1" z="1"/></element>
            <element type="point" label="D"><coords x="0" y="3" z="1"/></element>
            <command name="Polygon"><input a0="A" a1="B" a2="C" a3="D"/><output a0="poly"/></command>
            <element type="polygon" label="poly"/>
            <command name="Centroid"><input a0="poly"/><output a0="G"/></command>
            <element type="point" label="G"><coords x="1.6666666666666667" y="1.0833333333333333" z="1"/></element>
            """);
        var center = (IPoint)Find(drawing, "G");
        Near(center.Coordinates.X, expected: 5.0 / 3);
        Near(center.Coordinates.Y, expected: 13.0 / 12);
    }

    static void GeoGebraPrimedSquare()
    {
        var drawing = ReadGeoGebra("""
            <element type="point" label="A"><coords x="0" y="0" z="1"/></element>
            <element type="point" label="B"><coords x="1" y="0" z="1"/></element>
            <command name="Rotate"><input a0="B" a1="°" a2="A"/><output a0="B'"/></command>
            <element type="point" label="B'"/>
            <command name="Polygon"><input a0="A" a1="B'" a2="4"/><output a0="poly" a1="s1" a2="s2" a3="s3" a4="s4" a5="C" a6="D"/></command>
            <element type="polygon" label="poly"/>
            <element type="point" label="C"/>
            """);
        double angle = System.Math.PI / 180;
        var turned = ((IPoint)Find(drawing, "B'")).Coordinates;
        Near(turned.X, System.Math.Cos(angle));
        Near(turned.Y, System.Math.Sin(angle));
        var corner = ((IPoint)Find(drawing, "C")).Coordinates;
        Near(corner.X, System.Math.Cos(angle) - System.Math.Sin(angle));
        Near(corner.Y, System.Math.Sin(angle) + System.Math.Cos(angle));
    }

    static void GeoGebraConstantSlider()
    {
        var drawing = ReadGeoGebra("""
            <expression label="a" exp="2"/>
            <element type="numeric" label="a">
              <value val="2"/><slider x="0" y="0"/><show object="true"/>
            </element>
            """);
        Require(Find(drawing, "a") is DynamicGeometry.Slider, "A constant expression lost its slider.");
    }

    static void RejectUnrelatedWorksheet()
    {
        using var bytes = new MemoryStream();
        using (var archive = new ZipArchive(bytes, ZipArchiveMode.Create, leaveOpen: true))
        {
            using var writer = new StreamWriter(archive.CreateEntry(GeoGebraReader.WorksheetEntry).Open(), Encoding.UTF8);
            writer.Write("<unrelated />");
        }

        var worksheet = GeoGebraReader.ReadWorksheet(bytes.ToArray(), out string problem);
        Require(worksheet == null && !string.IsNullOrEmpty(problem), "A zip containing unrelated XML was accepted as a drawing.");
    }

    static Drawing ReadLgf(string xml)
    {
        var drawing = NewDrawing();
        bool suppressed = PointBase.SuppressAutoLabelPoints;
        PointBase.SuppressAutoLabelPoints = true;
        try
        {
            drawing.AddFromXml(XElement.Parse(xml));
        }
        finally
        {
            PointBase.SuppressAutoLabelPoints = suppressed;
        }

        drawing.Figures.CheckConsistency();
        return drawing;
    }

    static void PrimeNames()
    {
        var drawing = NewDrawing();
        var point = AddPoint(drawing, x: 3, y: 4);
        var other = AddPoint(drawing, x: 0, y: 0);
        var label = Factory.CreateLabel(drawing);
        Actions.Add(drawing, label);
        label.Text = "[A.X + AB]";
        Set(drawing, point, nameof(IFigure.Name), value: "A'");
        Require(label.Text == "[A'.X + A'B]", "The label says " + label.Text);
        Require(label.IsNumber, "The label lost its value: " + label.ProcessedText);
        Near(label.Value, expected: 8);
        var reloaded = ReadLgf(drawing.SaveAsText());
        var copy = reloaded.Figures.OfType<DynamicGeometry.Label>().Single();
        Near(copy.Value, expected: 8);
        Require(Scanner.IsName("A''") && Scanner.IsName("P_1'") && !Scanner.IsName("'A") && !Scanner.IsName("my point"), "Scanner.IsName");

        // a caption for a name; where an expression names the figure, it says it in backticks
        var free = AddPoint(drawing, x: 5, y: 5);
        Require(Rename(free, "Drag me!") == null && free.Name == "Drag me!", "A caption was refused for a name.");
        Require(Rename(other, "Drag me too") == null && other.Name == "Drag me too", "A name the label says in backticks was refused.");
        Require(label.Text == "[A'.X + dist(A', `Drag me too`)]", "The label says " + label.Text);
        Near(label.Value, expected: 8);
        Near(ReadLgf(drawing.SaveAsText()).Figures.OfType<DynamicGeometry.Label>().Single().Value, expected: 8);

        static string Rename(IFigure figure, string name)
        {
            var editor = new NameEditor { Value = PropertyDiscoveryStrategy.CreateValueProvider(figure, nameof(IFigure.Name)) };
            editor.TextBox.Text = name;
            editor.TextBox.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter });
            return string.IsNullOrEmpty(editor.ErrorText) ? null : editor.ErrorText;
        }
    }

    static void VectorMidpoint()
    {
        var drawing = NewDrawing();
        var first = AddPoint(drawing, x: -2, y: 0);
        var second = AddPoint(drawing, x: 4, y: 2);
        var vector = Factory.CreateVector(drawing, new IFigure[] { first, second });
        Actions.Add(drawing, vector);
        using var window = new TestWindow(drawing.Canvas);
        var placement = PointPlacement.Find(drawing, new Point(1, 1), snapToMidpoint: false);
        Require(placement.Kind == PointPlacementKind.Midpoint, "The middle of a vector gave " + placement.Kind);
        var midpoint = placement.Create(drawing);
        Require(midpoint is MidPoint && midpoint.Dependencies.SequenceEqual(new IFigure[] { first, second }), "Not the midpoint of the vector's ends.");

        // the Midpoint tool: one click on the vector
        drawing.Behavior = new MidpointCreator();
        Click(drawing, drawing.CoordinateSystem.ToPhysical(new Point(2.5, 1.5)));
        var made = drawing.Figures.OfType<MidPoint>().SingleOrDefault();
        Require(made != null, "A click on a vector made no midpoint.");
        Near(made.Coordinates.X, expected: 1);
        Near(made.Coordinates.Y, expected: 1);
    }

    static void DashedVector()
    {
        var drawing = NewDrawing();
        var first = AddPoint(drawing, x: -2, y: 0);
        var second = AddPoint(drawing, x: 4, y: 2);
        var vector = Factory.CreateVector(drawing, new IFigure[] { first, second });
        Actions.Add(drawing, vector);
        using var window = new TestWindow(drawing.Canvas);
        var dashed = drawing.StyleManager["DashedRedLine"];
        Require(dashed is LineStyle { Dash: not LineDash.Solid }, "No dashed palette style.");
        Set(drawing, vector, nameof(FigureBase.StyleDisplay), dashed);
        var shaft = vector.Line.Shape;
        Require(shaft.StrokeDashArray != null && shaft.StrokeDashArray.Count > 0, "The shaft is not dashed.");
        Require(
            shaft.Stroke is Avalonia.Media.ISolidColorBrush brush && brush.Color == ((LineStyle)dashed.Resolve()).Color,
            "The shaft is not in the style's color.");

        // the shaft ends where the head begins, not at its tip
        var tip = drawing.CoordinateSystem.ToPhysical(second.Coordinates);
        var end = new Point(shaft.EndPoint.X, shaft.EndPoint.Y);
        Require(end.Distance(tip) > Arrow.HeadLength, "The shaft runs through the head.");
        var reloaded = ReadLgf(drawing.SaveAsText());
        Require(reloaded.Figures.OfType<DynamicGeometry.Vector>().Single().Style?.Name == "DashedRedLine", "The vector's style was not saved.");
    }

    /// <summary>
    /// The regular polygon tools make plain figures - points by coordinates for the vertices
    /// they add, a segment per side and a polygon - regular whichever way they are built.
    /// Right after, the polygon's panel builds it again with another number of sides, one
    /// undo step, and the next polygon gets that count; its side moves the vertex clicked
    /// last, and undo puts it back - on a side through the side's own panel, which also
    /// turns it vertical; OK brings the tool's panel back. Read again, the drawing is the
    /// same.
    /// </summary>
    static void RegularPolygonTools()
    {
        int next = RegularPolygonCreator.NextNumberOfSides;
        try
        {
            foreach (var (tool, onSide) in new (RegularPolygonCreator Tool, bool OnSide)[] { (new RegularPolygonCreator(), false), (new RegularPolygonOnSideCreator(), true) })
            {
                var drawing = NewDrawing();
                object shown = null;
                drawing.DisplayProperties += (_, args) => shown = args.Object;
                drawing.Behavior = tool;
                Require(tool.PropertyBag is RegularPolygonCreator.SidesPanel, tool.Name + "'s panel is " + tool.PropertyBag);
                RegularPolygonCreator.NextNumberOfSides = 5;
                string empty = drawing.SaveAsText();
                new FigureCreator.Dialog(tool) { X = "0", Y = "0" }.AddPoint();
                new FigureCreator.Dialog(tool) { X = "2", Y = "0" }.AddPoint();
                Require(!drawing.IsRecordingTransaction, tool.Name + " did not finish.");
                CheckRegular(drawing, sides: 5, onSide, tool.Name);
                var panel = PolygonPanel(shown, onSide);
                Require(shown.ToString() == drawing.Figures.OfType<Polygon>().Single().Title, "The panel's title is " + shown);
                string made = drawing.SaveAsText();

                // another number of sides: the polygon again, one undo step
                Actions.SetProperty(drawing.ActionManager, PropertyDiscoveryStrategy.CreateValueProvider(panel, nameof(RegularPolygonPanel.NumberOfSides)), value: 8);
                CheckRegular(drawing, sides: 8, onSide, tool.Name + " with 8 sides");
                Require(RegularPolygonCreator.NextNumberOfSides == 8, "The next polygon gets " + RegularPolygonCreator.NextNumberOfSides + " sides.");
                string rebuilt = drawing.SaveAsText();
                drawing.ActionManager.Undo();
                Require(drawing.SaveAsText() == made, tool.Name + ": undo of 8 sides changed the drawing.");
                drawing.ActionManager.Redo();
                Require(drawing.SaveAsText() == rebuilt, tool.Name + ": redo of 8 sides changed the drawing.");
                drawing.Figures.CheckConsistency();

                // the side, undone; on a side its direction too, as after the Square; then OK
                // brings back the tool's panel
                panel = PolygonPanel(shown, onSide);
                if (shown is DimensionsPanel dimensions)
                {
                    Require(dimensions.GetProperties().Any(row => row.Name == nameof(DimensionsPanel.Direction)), "The side's panel has no direction.");
                    Actions.SetProperty(drawing.ActionManager, PropertyDiscoveryStrategy.CreateValueProvider(dimensions, nameof(DimensionsPanel.Length)), value: 3.0);
                    Near(dimensions.Length, expected: 3);
                }
                else
                {
                    Actions.SetProperty(drawing.ActionManager, PropertyDiscoveryStrategy.CreateValueProvider(panel, nameof(RegularPolygonPanel.Side)), value: 3.0);
                    Near(panel.Side, expected: 3);
                }

                CheckRegular(drawing, sides: 8, onSide, tool.Name + " with a side of 3");
                drawing.ActionManager.Undo();
                Require(drawing.SaveAsText() == rebuilt, tool.Name + ": undo of the side changed the drawing.");
                if (shown is DimensionsPanel sidePanel)
                {
                    var polygon = drawing.Figures.OfType<Polygon>().Single();
                    var side = (IFixableDirection)drawing.Figures.FindLine((IPoint)polygon.Dependencies[0], (IPoint)polygon.Dependencies[1]);
                    sidePanel.MakeVertical();
                    Require(System.Math.Abs(System.Math.Cos(side.Direction * System.Math.PI / 180)) < 1e-9, "Vertical left the side at " + side.Direction + "°.");
                    CheckRegular(drawing, sides: 8, onSide, tool.Name + " made vertical");
                    drawing.ActionManager.Undo();
                    Require(drawing.SaveAsText() == rebuilt, tool.Name + ": undo of Vertical changed the drawing.");
                    sidePanel.OK();
                }
                else
                {
                    panel.OK();
                }

                Require(shown is RegularPolygonCreator.SidesPanel, "OK brought back " + shown);

                var reread = ReadLgf(rebuilt);
                Require(reread.LoadErrors == null && reread.Figures.Count() == drawing.Figures.Count(), tool.Name + ": the drawing did not read back whole.");
                CheckRegular(reread, sides: 8, onSide, tool.Name + " read again");
                drawing.ActionManager.Undo();
                drawing.ActionManager.Undo();
                Require(drawing.SaveAsText() == empty, tool.Name + ": undo of the polygon left " + drawing.SaveAsText());
            }

            // on a side the drawing had, running the other way (segment B to A, the polygon on
            // A then B): the side's Horizontal replaces A, the first click, and a new count
            // built the polygon on the A that had left - a corner below the side and a second
            // side AB2
            {
                var drawing = ReadLgf("""
                    <Drawing Version="1">
                      <Figures>
                        <FreePoint Name="A" X="-3" Y="-1" />
                        <FreePoint Name="B" X="0" Y="0" />
                        <Segment Name="AB">
                          <Dependency Name="B" />
                          <Dependency Name="A" />
                        </Segment>
                      </Figures>
                    </Drawing>
                    """);
                object shown = null;
                drawing.DisplayProperties += (_, args) => shown = args.Object;
                var tool = new RegularPolygonOnSideCreator();
                drawing.Behavior = tool;
                RegularPolygonCreator.NextNumberOfSides = 5;
                new FigureCreator.Dialog(tool) { X = "-3", Y = "-1" }.AddPoint();
                new FigureCreator.Dialog(tool) { X = "0", Y = "0" }.AddPoint();
                CheckRegular(drawing, sides: 5, onSide: true, "A pentagon on side BA");
                var panel = PolygonPanel(shown, onSide: true);
                ((DimensionsPanel)shown).MakeHorizontal();
                Require(Find(drawing, "A") is TranslatedPoint, "Horizontal on side BA left A a " + Find(drawing, "A").GetType().Name);
                string horizontal = drawing.SaveAsText();
                Actions.SetProperty(drawing.ActionManager, PropertyDiscoveryStrategy.CreateValueProvider(panel, nameof(RegularPolygonPanel.NumberOfSides)), value: 6);
                CheckRegular(drawing, sides: 6, onSide: true, "On side BA made horizontal, with 6 sides");
                var polygon = drawing.Figures.OfType<Polygon>().Single();
                Require(polygon.Dependencies.All(drawing.Figures.Contains), "The hexagon is built on a point the drawing doesn't have.");
                var side = (IFixableDirection)drawing.Figures.FindLine((IPoint)polygon.Dependencies[0], (IPoint)polygon.Dependencies[1]);
                Require(System.Math.Abs(System.Math.Sin(side.Direction * System.Math.PI / 180)) < 1e-9, "The hexagon's side is at " + side.Direction + "°.");
                string rebuilt = drawing.SaveAsText();
                var reread = ReadLgf(rebuilt);
                CheckRegular(reread, sides: 6, onSide: true, "On side BA with 6 sides, read again");
                var vertices = polygon.Dependencies.Cast<IPoint>().Select(point => point.Coordinates).ToList();
                var rereadVertices = reread.Figures.OfType<Polygon>().Single().Dependencies.Cast<IPoint>().Select(point => point.Coordinates).ToList();
                for (int i = 0; i < vertices.Count; i++)
                {
                    NearPoint(rereadVertices[i], vertices[i]);
                }

                drawing.ActionManager.Undo();
                Require(drawing.SaveAsText() == horizontal, "Undo of 6 sides on side BA changed the drawing.");
                drawing.ActionManager.Redo();
                Require(drawing.SaveAsText() == rebuilt, "Redo of 6 sides on side BA changed the drawing.");
            }
        }
        finally
        {
            RegularPolygonCreator.NextNumberOfSides = next;
        }

        // by its center the polygon's own panel; on a side the side's, the number of sides on top
        static RegularPolygonPanel PolygonPanel(object shown, bool onSide)
        {
            if (onSide)
            {
                Require(shown is DimensionsPanel { RegularPolygon: not null }, "After the polygon on a side the panel shows " + shown);
                return ((DimensionsPanel)shown).RegularPolygon;
            }

            Require(shown is RegularPolygonPanel, "After the polygon the panel shows " + shown);
            return (RegularPolygonPanel)shown;
        }

        static void CheckRegular(Drawing drawing, int sides, bool onSide, string what)
        {
            var polygon = drawing.Figures.OfType<Polygon>().Single();
            var vertices = polygon.Dependencies.Cast<IPoint>().Select(point => point.Coordinates).ToList();
            Require(vertices.Count == sides, what + ": " + vertices.Count + " vertices");
            var center = new Point(vertices.Average(point => point.X), vertices.Average(point => point.Y));
            for (int i = 0; i < sides; i++)
            {
                Near(vertices[i].Distance(vertices[(i + 1) % sides]), vertices[0].Distance(vertices[1]));
                Near(vertices[i].Distance(center), vertices[0].Distance(center));
            }

            int added = drawing.Figures.OfType<PointByCoordinates>().Count();
            Require(added == sides - (onSide ? 2 : 1), what + ": " + added + " points by coordinates");
            Require(drawing.Figures.OfType<Segment>().Count() == sides, what + ": " + drawing.Figures.OfType<Segment>().Count() + " segments");
            drawing.Figures.CheckConsistency();
        }
    }

    /// <summary>A press, moves along the way and a release, as a mouse makes them</summary>
    static void Drag(Drawing drawing, Point from, Point to, KeyModifiers modifiers = KeyModifiers.None)
    {
        var canvas = drawing.Canvas;
        using var pointer = new Pointer(Pointer.GetNextFreeId(), PointerType.Mouse, isPrimary: true);
        canvas.RaiseEvent(new PointerPressedEventArgs(
            canvas,
            pointer,
            canvas,
            from,
            timestamp: 0,
            new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.LeftButtonPressed),
            modifiers,
            clickCount: 1));
        for (int step = 1; step <= 12; step++)
        {
            canvas.RaiseEvent(new PointerEventArgs(
                InputElement.PointerMovedEvent,
                canvas,
                pointer,
                canvas,
                from + (to - from) * (step / 12.0),
                timestamp: (ulong)step,
                new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.Other),
                modifiers));
        }

        canvas.RaiseEvent(new PointerReleasedEventArgs(
            canvas,
            pointer,
            canvas,
            to,
            timestamp: 13,
            new PointerPointProperties(RawInputModifiers.None, PointerUpdateKind.LeftButtonReleased),
            modifiers,
            MouseButton.Left));
    }

    static void NearPoint(Point actual, Point expected)
    {
        Near(actual.X, expected.X);
        Near(actual.Y, expected.Y);
    }

    static Point HandleOffset(BezierPath path, int anchor, bool isIn)
    {
        return ((BezierPath.BezierPathHandle)path.GetPart((isIn ? "In" : "Out") + (anchor + 1))).Offset;
    }

    /// <summary>A click, a press dragged out, a click, and a click on the first anchor: a closed path with one smooth anchor</summary>
    static void BezierPathTool()
    {
        var drawing = NewDrawing();
        using var window = new TestWindow(drawing.Canvas);
        var system = drawing.CoordinateSystem;
        string empty = drawing.SaveAsText();
        var tool = new BezierPathCreator();
        drawing.Behavior = tool;
        Click(drawing, system.ToPhysical(new Point(0, 0)));
        Drag(drawing, system.ToPhysical(new Point(3, 0)), system.ToPhysical(new Point(4, 1)));
        Click(drawing, system.ToPhysical(new Point(3, -3)));
        Require(drawing.IsRecordingTransaction, "The path ended before it was closed.");
        Click(drawing, system.ToPhysical(new Point(0, 0)));
        Require(!drawing.IsRecordingTransaction, "A click on the first anchor did not close the path.");
        var path = drawing.Figures.OfType<BezierPath>().Single();
        Require(path.Closed && path.Filled && path.AnchorCount == 3, "The path: closed " + path.Closed + ", " + path.AnchorCount + " anchors.");
        Require(drawing.Figures.OfType<FreePoint>().Count() == 3, "The anchors are not three free points.");
        NearPoint(HandleOffset(path, anchor: 1, isIn: false), new Point(1, 1));
        NearPoint(HandleOffset(path, anchor: 1, isIn: true), new Point(-1, -1));
        // the clicked anchors' handles are the path's, worked out: smooth, not on the anchor
        var outA = (BezierPath.BezierPathHandle)path.GetPart("Out1");
        Require(outA.Auto && outA.Offset != default, "A clicked anchor's handle is not worked out by the path.");
        Require(path.Name == "ABC", "The path is named " + path.Name);
        drawing.Figures.CheckConsistency();
        string after = drawing.SaveAsText();
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == empty, "Undo of the path left something.");
        drawing.ActionManager.Redo();
        Require(drawing.SaveAsText() == after, "Redo of the path changed it.");
        var reloaded = ReadLgf(after);
        Require(reloaded.LoadErrors == null && reloaded.SaveAsText() == after, "The path did not round trip.");

        // open, with Enter: not filled
        Click(drawing, system.ToPhysical(new Point(-5, 3)));
        Click(drawing, system.ToPhysical(new Point(-3, 4)));
        tool.KeyDown(null, new KeyEventArgs() { Key = Key.Enter });
        var open = drawing.Figures.OfType<BezierPath>().Single(p => p != path);
        Require(!open.Closed && !open.Filled && open.AnchorCount == 2, "Enter did not leave an open path of two anchors.");
        drawing.Figures.CheckConsistency();
    }

    /// <summary>
    /// Path ABC (open, every handle on its anchor): with B selected its two handles show and
    /// the ones of A and C that face it; Tab at B takes a handle, which a drag pulls out; with
    /// Alt the handle across B goes the opposite way; undo puts both back
    /// </summary>
    static void BezierPathHandles()
    {
        var drawing = NewDrawing();
        using var window = new TestWindow(drawing.Canvas);
        var system = drawing.CoordinateSystem;
        var a = AddPoint(drawing, x: 0, y: 0);
        var b = AddPoint(drawing, x: 4, y: 0);
        var c = AddPoint(drawing, x: 8, y: 0);
        var zero = new Point[3];
        var path = BezierPath.Create(drawing, new IFigure[] { a, b, c }, zero, zero, closed: false, filled: false);
        Actions.Add(drawing, path);
        drawing.Behavior = new Dragger();
        Require(!path.Handles.Any(h => h.Visible), "Handles show with nothing selected.");

        // not shown, the handles on an anchor are still there to be taken with Tab
        string hidden = null;
        Action<string> watch = text => hidden = text;
        drawing.ChoiceStatus += watch;
        Hover(drawing, system.ToPhysical(a.Coordinates));
        // (the point, its hidden handle, the path)
        Require(hidden != null && hidden.Contains("(1 of 3)"), "Over A, nothing selected: " + hidden);
        Require(drawing.Behavior.StepChoice(backwards: false) && hidden.StartsWith("Handle of A toward B"), "Tab at A took: " + hidden);
        string untouched = drawing.SaveAsText();
        Drag(drawing, system.ToPhysical(a.Coordinates), system.ToPhysical(new Point(1, 1)));
        NearPoint(a.Coordinates, new Point(0, 0));
        NearPoint(HandleOffset(path, anchor: 0, isIn: false), new Point(1, 1));
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == untouched, "Undo of pulling out a hidden handle.");
        drawing.ChoiceStatus -= watch;

        // (a test window draws no frames: a tool takes one move until the next press, so
        // the hovers below are a fresh tool's)
        drawing.Behavior = new Dragger();

        b.Selected = true;
        drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());
        var shown = path.Handles.Where(h => h.Visible).Select(h => path.GetPartName(h)).OrderBy(n => n).ToList();
        Require(string.Join(" ", shown) == "In2 In3 Out1 Out2", "Handles shown: " + string.Join(" ", shown));

        string status = null;
        drawing.ChoiceStatus += text => status = text;
        var atB = system.ToPhysical(b.Coordinates);
        Hover(drawing, atB);
        // (the point and its two handles, which the path's own hit test answers with)
        Require(status != null && status.Contains("(1 of 3)"), "No choice at B: " + status);
        Require(drawing.Behavior.StepChoice(backwards: false), "Tab chose nothing at B.");
        Require(status.StartsWith("Handle of B toward A"), "Tab took: " + status);
        string before = drawing.SaveAsText();
        Drag(drawing, atB, system.ToPhysical(new Point(4, 2)));
        NearPoint(b.Coordinates, new Point(4, 0));
        NearPoint(HandleOffset(path, anchor: 1, isIn: true), new Point(0, 2));
        NearPoint(HandleOffset(path, anchor: 1, isIn: false), new Point(0, -2));
        Require(b.Selected, "The drag of a handle lost the selection.");
        string pulled = drawing.SaveAsText();
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the handle's drag.");
        drawing.ActionManager.Redo();
        Require(drawing.SaveAsText() == pulled, "Redo of the handle's drag.");

        // with Alt the out handle stays where it is (a corner)
        Drag(drawing, system.ToPhysical(new Point(4, 2)), system.ToPhysical(new Point(3, 3)), KeyModifiers.Alt);
        NearPoint(HandleOffset(path, anchor: 1, isIn: true), new Point(-1, 3));
        NearPoint(HandleOffset(path, anchor: 1, isIn: false), new Point(0, -2));
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == pulled, "Undo of the drag with Alt.");

        // without, it snaps back to the mirror image
        Drag(drawing, system.ToPhysical(new Point(4, 2)), system.ToPhysical(new Point(5, 3)));
        NearPoint(HandleOffset(path, anchor: 1, isIn: false), new Point(-1, -3));
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == pulled, "Undo of the mirrored drag.");

        // a click on the handle keeps B selected
        Click(drawing, system.ToPhysical(new Point(4, 2)));
        Require(b.Selected && !path.Selected, "A click on a handle changed the selection.");
        drawing.Figures.CheckConsistency();

        // no other tool takes a handle
        drawing.Behavior = new FreePointCreator();
        var placement = PointPlacement.Find(drawing, new Point(4, 2), snapToMidpoint: false);
        Require(placement.ExistingPoint == null, "The Point tool took a handle.");
    }

    /// <summary>Path ABCD: B dragged with Alt onto C leaves the path ACD, C taking B's handle toward A</summary>
    static void BezierPathDropAnchor()
    {
        var drawing = NewDrawing();
        using var window = new TestWindow(drawing.Canvas);
        var system = drawing.CoordinateSystem;
        var points = new[] { AddPoint(drawing, x: 0, y: 0), AddPoint(drawing, x: 3, y: 2), AddPoint(drawing, x: 6, y: 0), AddPoint(drawing, x: 9, y: 2) };
        var ins = new[] { new Point(0, 0), new Point(-1, 1), new Point(-0.5, 0.5), new Point(0, 0) };
        var outs = new[] { new Point(1, 1), new Point(1, -1), new Point(0.5, -0.5), new Point(0, 0) };
        var path = BezierPath.Create(drawing, points, ins, outs, closed: false, filled: false);
        Actions.Add(drawing, path);
        drawing.Behavior = new Dragger();
        string before = drawing.SaveAsText();
        Drag(drawing, system.ToPhysical(points[1].Coordinates), system.ToPhysical(points[2].Coordinates), KeyModifiers.Alt);
        Require(!drawing.Figures.Contains(points[1]), "B was not joined into C.");
        Require(drawing.Figures.Contains(path) && path.AnchorCount == 3, "The path is not ACD.");
        Require(path.Dependencies.SequenceEqual(new IFigure[] { points[0], points[2], points[3] }), "The anchors: " + string.Join(", ", path.Dependencies));
        NearPoint(HandleOffset(path, anchor: 1, isIn: true), new Point(-1, 1));
        NearPoint(HandleOffset(path, anchor: 1, isIn: false), new Point(0.5, -0.5));
        drawing.Figures.CheckConsistency();
        string after = drawing.SaveAsText();
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the drop changed the drawing.");
        drawing.Figures.CheckConsistency();
        drawing.ActionManager.Redo();
        Require(drawing.SaveAsText() == after, "Redo of the drop changed the drawing.");

        // A onto D, not next to each other: no join
        Drag(drawing, system.ToPhysical(points[0].Coordinates), system.ToPhysical(points[3].Coordinates), KeyModifiers.Alt);
        Require(drawing.Figures.Contains(points[0]) && path.AnchorCount == 3, "A was joined into D, which is not next to it.");
        drawing.Figures.CheckConsistency();
    }

    static void BezierPathFigure()
    {
        var drawing = NewDrawing();
        using var window = new TestWindow(drawing.Canvas);
        var a = AddPoint(drawing, x: 0, y: 4);
        var b = AddPoint(drawing, x: 4, y: -2);
        var c = AddPoint(drawing, x: -4, y: -2);
        var path = BezierPath.Create(
            drawing,
            new IFigure[] { a, b, c },
            new[] { new Point(-1, 0), new Point(0, 1), new Point(1, 0) },
            new[] { new Point(1, 0), new Point(0, -1), new Point(-1, 0) },
            closed: true,
            filled: true);
        Actions.Add(drawing, path);

        // a point on the second piece, where the cubic's own parameter says
        var onPath = Factory.CreatePointOnFigure(drawing, path, new Point(0, -3));
        Actions.Add(drawing, onPath);
        Require(onPath.Parameter >= 1 && onPath.Parameter <= 2, "The point near the bottom is on piece " + onPath.Parameter);
        NearPoint(onPath.Coordinates, path.GetPointFromParameter(onPath.Parameter));
        Near(path.GetNearestParameterFromPoint(path.GetPointFromParameter(0.3)), expected: 0.3);
        Require(onPath.Exists, "The point on the path doesn't exist.");

        // opened, the closing piece goes, and a point on it with it
        var onClosing = Factory.CreatePointOnFigure(drawing, path, path.GetPointFromParameter(2.5));
        Actions.Add(drawing, onClosing);
        Set(drawing, path, nameof(BezierPath.Closed), false);
        drawing.Recalculate();
        Require(!onClosing.Exists && onPath.Exists, "Opening the path: the points on it.");
        drawing.ActionManager.Undo();
        drawing.Recalculate();
        Require(onClosing.Exists, "Undo of opening the path.");

        // a deleted anchor leaves a path of two
        string before = drawing.SaveAsText();
        Actions.Remove(c);
        Require(drawing.Figures.Contains(path) && path.AnchorCount == 2 && onPath.Exists, "Deleting an anchor: " + path.AnchorCount + " anchors.");
        drawing.Figures.CheckConsistency();
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of deleting an anchor.");
        drawing.Figures.CheckConsistency();

        // a hole
        var hole = BezierPath.Create(
            drawing,
            new IFigure[] { AddPoint(drawing, x: -1, y: 0), AddPoint(drawing, x: 1, y: 0), AddPoint(drawing, x: 0, y: 1) },
            new Point[3],
            new Point[3],
            closed: true,
            filled: true);
        Actions.Add(drawing, hole);
        Require(BezierPath.CanCutHoles(new IFigure[] { path, hole }), "Two paths can't be cut.");
        string withHole = drawing.SaveAsText();
        BezierPath.CutHoles(drawing, new IFigure[] { hole, path });
        Require(path.Holes.Single() == hole && !hole.Filled, "The small path is not an unfilled hole of the big one.");
        Require(drawing.Figures.IndexOf(hole) < drawing.Figures.IndexOf(path), "The hole comes after the path built on it.");
        Require(path.HitTest(new Point(0, 0.3)) == null && path.HitTest(new Point(0, 2.5)) != null, "The hole is not left out of the inside.");
        drawing.Figures.CheckConsistency();
        string cut = drawing.SaveAsText();
        var reloaded = ReadLgf(cut);
        Require(reloaded.LoadErrors == null && reloaded.SaveAsText() == cut, "The holes did not round trip.");
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == withHole, "Undo of cutting the hole.");
        drawing.ActionManager.Redo();
        Actions.Remove(hole);
        Require(drawing.Figures.Contains(path) && !path.Holes.Any(), "Deleting the hole took the path.");
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == cut, "Undo of deleting the hole.");
        drawing.Figures.CheckConsistency();

        // reflected in the y axis: the image's handles are the images of the source's
        var mirror = Factory.CreateLineTwoPoints(drawing, new IFigure[] { AddPoint(drawing, x: 6, y: 0), AddPoint(drawing, x: 6, y: 1) });
        Actions.Add(drawing, mirror);
        Require(Transformer.CanBeSource(path), "The path can't be reflected.");
        var images = Transformer.CreateReflectedFigure(drawing, path, mirror);
        foreach (var image in images)
        {
            Actions.Add(drawing, image);
        }

        var reflected = (BezierPath)images.Last();
        Require(reflected.HasHandlePoints && reflected.AnchorCount == 3 && reflected.Holes.Count() == 1, "The image is not a path of three anchors with a hole.");
        NearPoint(reflected.GetPointFromParameter(0.4), new Point(12 - path.GetPointFromParameter(0.4).X, path.GetPointFromParameter(0.4).Y));
        drawing.Figures.CheckConsistency();
        string imaged = drawing.SaveAsText();
        var reread = ReadLgf(imaged);
        Require(reread.LoadErrors == null && reread.SaveAsText() == imaged, "The image did not round trip.");

        // the image follows the source's handles
        var handle = (BezierPath.BezierPathHandle)path.GetPart("Out1");
        handle.MoveTo(handle.Coordinates.Plus(new Point(1, 1)));
        drawing.Recalculate();
        NearPoint(reflected.GetPointFromParameter(0.4), new Point(12 - path.GetPointFromParameter(0.4).X, path.GetPointFromParameter(0.4).Y));

        // not in a circle
        var circle = Factory.CreateCircle(drawing, new IFigure[] { AddPoint(drawing, x: 10, y: 10), AddPoint(drawing, x: 11, y: 10) });
        Actions.Add(drawing, circle);
        Require(!Transformer.CanFigureBeMirrorForSource(circle, path), "A circle reflects a path.");
    }

    /// <summary>
    /// Closed path ABC; P halfway along its first piece becomes an anchor: four anchors, the
    /// same curve (a place three quarters along the old piece is half along the new second
    /// one), the other points on the path where they were
    /// </summary>
    static void BezierPathInsertAnchor()
    {
        var drawing = NewDrawing();
        using var window = new TestWindow(drawing.Canvas);
        var a = AddPoint(drawing, x: 0, y: 4);
        var b = AddPoint(drawing, x: 4, y: -2);
        var c = AddPoint(drawing, x: -4, y: -2);
        var path = BezierPath.Create(
            drawing,
            new IFigure[] { a, b, c },
            new[] { new Point(-2, 0), new Point(0, 2), new Point(1, 1) },
            new[] { new Point(2, 0), new Point(0, -2), new Point(-1, -1) },
            closed: true,
            filled: true);
        Actions.Add(drawing, path);
        var oldQuarter = path.GetPointFromParameter(0.25);
        var oldThreeQuarters = path.GetPointFromParameter(0.75);
        var oldSecond = path.GetPointFromParameter(1.5);
        var p = Factory.CreatePointOnFigure(drawing, path, path.GetPointFromParameter(0.5));
        var q = Factory.CreatePointOnFigure(drawing, path, oldQuarter);
        var r = Factory.CreatePointOnFigure(drawing, path, oldSecond);
        Actions.Add(drawing, p);
        Actions.Add(drawing, q);
        Actions.Add(drawing, r);
        Near(p.Parameter, expected: 0.5);
        Require(BezierPath.CanBecomeAnchor(p) && !BezierPath.CanBecomeAnchor(a), "Which points can become anchors.");
        string name = p.Name;
        string before = drawing.SaveAsText();
        p.ConvertToPathAnchor();
        Require(path.AnchorCount == 4 && path.Anchor(1) is FreePoint anchor && anchor.Name == name, "P is not the second anchor.");
        Require(drawing.Figures.IndexOf((IFigure)path.Anchor(1)) < drawing.Figures.IndexOf(path), "The new anchor comes after the path.");
        NearPoint(path.GetPointFromParameter(0.5), oldQuarter);
        NearPoint(path.GetPointFromParameter(1.5), oldThreeQuarters);
        NearPoint(path.GetPointFromParameter(2.5), oldSecond);
        drawing.Recalculate();
        NearPoint(q.Coordinates, oldQuarter);
        NearPoint(r.Coordinates, oldSecond);
        drawing.Figures.CheckConsistency();
        string after = drawing.SaveAsText();
        Require(ReadLgf(after).SaveAsText() == after, "The path with the new anchor did not round trip.");
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the new anchor changed the drawing.");
        drawing.Figures.CheckConsistency();
        drawing.ActionManager.Redo();
        Require(drawing.SaveAsText() == after, "Redo of the new anchor changed the drawing.");
        drawing.Figures.CheckConsistency();
    }

    /// <summary>
    /// A, then Alt+click on point P (A's out handle), Alt+click on empty paper (the in handle of
    /// the next anchor), B, Enter: the path AB pulled towards P and the spot
    /// </summary>
    static void BezierPathAltClickHandles()
    {
        var drawing = NewDrawing();
        using var window = new TestWindow(drawing.Canvas);
        var system = drawing.CoordinateSystem;
        var p = AddPoint(drawing, x: 2, y: 4);
        string before = drawing.SaveAsText();
        string status = null;
        drawing.Status += text => status = text;
        var tool = new BezierPathCreator();
        drawing.Behavior = tool;
        Click(drawing, system.ToPhysical(new Point(6, 0)), KeyModifiers.Alt);
        Require(status != null && status.Contains("first point"), "Alt+click before any point: " + status);
        Click(drawing, system.ToPhysical(new Point(0, 0)));
        Click(drawing, system.ToPhysical(p.Coordinates), KeyModifiers.Alt);
        Click(drawing, system.ToPhysical(new Point(6, 3)), KeyModifiers.Alt);
        Click(drawing, system.ToPhysical(new Point(6, -3)), KeyModifiers.Alt);
        Require(status.Contains("Both handles"), "A third Alt+click: " + status);
        Click(drawing, system.ToPhysical(new Point(8, 0)));
        tool.KeyDown(null, new KeyEventArgs() { Key = Key.Enter });
        var path = drawing.Figures.OfType<BezierPath>().Single();
        Require(path.AnchorCount == 2 && !path.Closed, "The path is not two anchors, open.");
        var outA = (BezierPath.BezierPathHandle)path.GetPart("Out1");
        Require(outA.Point == p && path.Dependencies.Contains(p), "A's out handle is not P.");
        NearPoint(HandleOffset(path, anchor: 1, isIn: true), new Point(-2, 3));
        Require(drawing.Figures.OfType<FreePoint>().Count() == 3, "Alt+click made points of its own.");
        drawing.Figures.CheckConsistency();
        string after = drawing.SaveAsText();
        Require(after.Contains("Path=\"C #2 -2,3 C a a\""), "The path's text: " + after);
        Require(ReadLgf(after).SaveAsText() == after, "The path with a point for a handle did not round trip.");
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the path.");
        drawing.ActionManager.Redo();
        Require(drawing.SaveAsText() == after, "Redo of the path.");
    }

    /// <summary>
    /// Open path AB whose A has point P for its out handle: P moved bends the curve; P deleted
    /// leaves an ordinary handle where it was; the in handle of B dropped with Alt onto Q is Q;
    /// a point on the path split beside P makes P's handle an ordinary one, the curve the same
    /// </summary>
    static void BezierPathPointHandles()
    {
        var drawing = NewDrawing();
        using var window = new TestWindow(drawing.Canvas);
        var system = drawing.CoordinateSystem;
        var a = AddPoint(drawing, x: 0, y: 0);
        var b = AddPoint(drawing, x: 8, y: 0);
        var p = AddPoint(drawing, x: 2, y: 4);
        var q = AddPoint(drawing, x: 7, y: 5);
        var path = BezierPath.Create(
            drawing,
            new IFigure[] { a, b },
            new[] { new BezierPath.HandleSpec(), new BezierPath.HandleSpec(new Point(-1, 2), null) },
            new[] { new BezierPath.HandleSpec(default, p), new BezierPath.HandleSpec() },
            holes: Array.Empty<IFigure>(),
            closed: false,
            filled: false);
        Actions.Add(drawing, path);
        drawing.Figures.CheckConsistency();
        var middle = path.GetPointFromParameter(0.5);
        p.MoveTo(new Point(2, 6));
        p.RecalculateAllDependents();
        Require(path.GetPointFromParameter(0.5).Y > middle.Y + 0.5, "Moving P did not bend the curve.");
        p.MoveTo(new Point(2, 4));
        p.RecalculateAllDependents();
        NearPoint(path.GetPointFromParameter(0.5), middle);

        // deleted: an ordinary handle where P was, the curve the same
        string withP = drawing.SaveAsText();
        Actions.Remove(p);
        Require(drawing.Figures.Contains(path), "Deleting P took the path.");
        var outA = (BezierPath.BezierPathHandle)path.GetPart("Out1");
        Require(outA.Point == null, "A's handle is still a point.");
        NearPoint(outA.Offset, new Point(2, 4));
        NearPoint(path.GetPointFromParameter(0.5), middle);
        drawing.Figures.CheckConsistency();
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == withP, "Undo of deleting P.");
        drawing.Figures.CheckConsistency();

        // B's in handle dragged with Alt onto Q: it is Q
        drawing.Behavior = new Dragger();
        b.Selected = true;
        drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());
        Drag(drawing, system.ToPhysical(new Point(7, 2)), system.ToPhysical(q.Coordinates), KeyModifiers.Alt);
        var inB = (BezierPath.BezierPathHandle)path.GetPart("In2");
        Require(inB.Point == q, "B's in handle is not Q.");
        Require(drawing.Figures.IndexOf(q) < drawing.Figures.IndexOf(path), "Q comes after the path built on it.");
        drawing.Figures.CheckConsistency();
        string onQ = drawing.SaveAsText();
        Require(ReadLgf(onQ).SaveAsText() == onQ, "Two points as handles did not round trip.");
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == withP, "Undo of the drop onto Q.");
        drawing.ActionManager.Redo();
        Require(drawing.SaveAsText() == onQ, "Redo of the drop onto Q.");

        // split: the piece's handles that are points become ordinary ones, the curve the same
        var quarter = path.GetPointFromParameter(0.25);
        var threeQuarters = path.GetPointFromParameter(0.75);
        var onPath = Factory.CreatePointOnFigure(drawing, path, path.GetPointFromParameter(0.5));
        Actions.Add(drawing, onPath);
        onPath.ConvertToPathAnchor();
        Require(path.AnchorCount == 3 && !path.HasHandlePoints, "The split left points as handles.");
        NearPoint(path.GetPointFromParameter(0.5), quarter);
        NearPoint(path.GetPointFromParameter(1.5), threeQuarters);
        drawing.Figures.CheckConsistency();
        drawing.ActionManager.Undo();
        Require(path.AnchorCount == 2 && inB.Point == q && outA.Point == p, "Undo of the split.");
        drawing.Figures.CheckConsistency();
    }

    static void Close(Point actual, Point expected, string what)
    {
        Require(
            double.IsFinite(actual.X) && double.IsFinite(actual.Y) && actual.Distance(expected) <= 1e-9,
            what + $": expected {expected}, got {actual}.");
    }

    /// <summary>
    /// The smoothing methods: four points on a circle give the handles each method is known
    /// for; moved, turned, scaled and mirrored anchors give the handles so moved, also next
    /// to handles the user set; then a path through A, B, C, D on a circle whose tension is a
    /// slider: read, followed, untied, the slider deleted; another method; a handle dragged
    /// (the user's from then on), the anchor verbs, a split between automatic handles; and
    /// the tool's clicks, whose path takes its tension from a slider clicked right after
    /// </summary>
    static void BezierPathSmoothing()
    {
        var circle = new[] { new Point(1, 0), new Point(0, 1), new Point(-1, 0), new Point(0, -1) };
        var unknown = new Point?[4];
        void Expect(BezierPathSmoothing smoothing, double tension, double length)
        {
            var (ins, outs) = BezierPathSmoother.Smooth(circle, unknown, unknown, closed: true, smoothing, tension);
            for (int i = 0; i < circle.Length; i++)
            {
                var tangent = new Point(-circle[i].Y, circle[i].X);
                Close(outs[i], tangent.Scale(length), smoothing + " out handle " + i);
                Close(ins[i], tangent.Scale(-length), smoothing + " in handle " + i);
            }
        }

        double kappa = 4 * (System.Math.Sqrt(2) - 1) / 3;
        Expect(DynamicGeometry.BezierPathSmoothing.Hobby, tension: 1, kappa);
        Expect(DynamicGeometry.BezierPathSmoothing.Hobby, tension: 2, kappa / 2);
        Expect(DynamicGeometry.BezierPathSmoothing.CatmullRom, tension: 1, 1.0 / 3);
        Expect(DynamicGeometry.BezierPathSmoothing.NaturalSpline, tension: 1, 0.5);
        Expect(DynamicGeometry.BezierPathSmoothing.None, tension: 1, 0);

        // the same curve for anchors moved, turned, scaled, mirrored: next to a handle set by
        // the user (C's out handle) and a corner (D's in handle on D)
        var anchors = new[] { new Point(0, 0), new Point(3, 1), new Point(5, 4), new Point(2, 6), new Point(-1, 3) };
        var givenIns = new Point?[] { null, null, null, new Point(0, 0), null };
        var givenOuts = new Point?[] { null, null, new Point(0.5, 2), null, null };
        Point Turn(Point vector, bool mirror)
        {
            double cos = System.Math.Cos(0.5);
            double sin = System.Math.Sin(0.5);
            var turned = new Point(2.5 * (vector.X * cos - vector.Y * sin), 2.5 * (vector.X * sin + vector.Y * cos));
            return mirror ? new Point(turned.X, -turned.Y) : turned;
        }

        foreach (var smoothing in new[] { DynamicGeometry.BezierPathSmoothing.Hobby, DynamicGeometry.BezierPathSmoothing.CatmullRom, DynamicGeometry.BezierPathSmoothing.NaturalSpline })
        {
            foreach (bool closed in new[] { true, false })
            {
                var (ins, outs) = BezierPathSmoother.Smooth(anchors, givenIns, givenOuts, closed, smoothing, tension: 1.3);
                var kept = ins[2];
                Require(
                    kept.X * givenOuts[2].Value.Y - kept.Y * givenOuts[2].Value.X is var cross && System.Math.Abs(cross) < 1e-9 && kept.X < 0,
                    smoothing + ": the automatic handle across a handle set is not straight on.");
                foreach (bool mirror in new[] { false, true })
                {
                    var (movedIns, movedOuts) = BezierPathSmoother.Smooth(
                        anchors.Select(p => Turn(p, mirror).Plus(new Point(7, -2))).ToList(),
                        givenIns.Select(o => o.HasValue ? Turn(o.Value, mirror) : (Point?)null).ToList(),
                        givenOuts.Select(o => o.HasValue ? Turn(o.Value, mirror) : (Point?)null).ToList(),
                        closed,
                        smoothing,
                        tension: 1.3);
                    for (int i = 0; i < anchors.Length; i++)
                    {
                        Close(movedIns[i], Turn(ins[i], mirror), smoothing + " moved, in handle " + i);
                        Close(movedOuts[i], Turn(outs[i], mirror), smoothing + " moved, out handle " + i);
                    }
                }
            }
        }

        // a path whose tension is the slider
        string text = """
            <Drawing Version="1">
              <Figures>
                <FreePoint Name="A" X="3" Y="0" />
                <FreePoint Name="B" X="0" Y="3" />
                <FreePoint Name="C" X="-3" Y="0" />
                <FreePoint Name="D" X="0" Y="-3" />
                <Slider Name="s" X="-6" Y="5" Value="2" />
                <BezierPath Name="ABCD" Closed="true" Tension="#4" Path="C a a C a a C a a C a a">
                  <Dependency Name="A" />
                  <Dependency Name="B" />
                  <Dependency Name="C" />
                  <Dependency Name="D" />
                  <Dependency Name="s" />
                </BezierPath>
              </Figures>
            </Drawing>
            """;
        var drawing = ReadLgf(text);
        using var window = new TestWindow(drawing.Canvas);
        var system = drawing.CoordinateSystem;
        var path = drawing.Figures.OfType<BezierPath>().Single();
        var slider = (DynamicGeometry.Slider)Find(drawing, "s");
        var a = (FreePoint)Find(drawing, "A");
        var outA = (BezierPath.BezierPathHandle)path.GetPart("Out1");
        var inA = (BezierPath.BezierPathHandle)path.GetPart("In1");
        Require(drawing.LoadErrors == null && path.GetSource(nameof(BezierPath.Tension)) == slider, "The tension is not the slider.");
        Close(outA.Offset, new Point(0, 3 * kappa / 2), "Tension 2");
        string original = drawing.SaveAsText();
        Require(original.Contains("Tension=\"#4\"") && original.Contains("Path=\"C a a C a a C a a C a a\""), "The path's text: " + original);
        Require(ReadLgf(original).SaveAsText() == original, "The smoothed path did not round trip.");
        Set(drawing, slider, nameof(DynamicGeometry.Slider.Value), 1.0);
        Close(outA.Offset, new Point(0, 3 * kappa), "Tension 1 from the slider");
        drawing.ActionManager.Undo();
        Close(outA.Offset, new Point(0, 3 * kappa / 2), "Undo of the slider");

        // untied: typed at the value it had; the slider deleted: the same
        Require(path.Detach(nameof(BezierPath.Tension)) && path.GetSource(nameof(BezierPath.Tension)) == null, "Untie left the slider.");
        Require(!path.Dependencies.Contains(slider) && path.Tension == 2, "Untied, the tension is " + path.Tension);
        Close(outA.Offset, new Point(0, 3 * kappa / 2), "Untied");
        drawing.Figures.CheckConsistency();
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == original, "Undo of the untie.");
        Actions.Remove(slider);
        Require(drawing.Figures.Contains(path) && path.Tension == 2 && path.Exists, "Deleting the slider took the path or its tension.");
        drawing.Figures.CheckConsistency();
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == original, "Undo of deleting the slider.");
        drawing.Figures.CheckConsistency();

        // another method
        Set(drawing, path, nameof(BezierPath.Smoothing), DynamicGeometry.BezierPathSmoothing.CatmullRom);
        Close(outA.Offset, new Point(0, 0.5), "Catmull-Rom at tension 2");
        Require(drawing.SaveAsText().Contains("Smoothing=\"CatmullRom\""), "The method is not saved.");
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == original, "Undo of the method.");

        // A's out handle dragged: the user's, B's still the path's; undo: the path's again
        drawing.Behavior = new Dragger();
        a.Selected = true;
        drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());
        Drag(drawing, system.ToPhysical(a.Coordinates.Plus(outA.Offset)), system.ToPhysical(new Point(4, 2)));
        Require(!outA.Auto && !inA.Auto, "A dragged handle is still automatic.");
        Close(outA.Offset, new Point(1, 2), "The handle dragged");
        Close(inA.Offset, new Point(-1, -2), "The handle across it");
        Require(((BezierPath.BezierPathHandle)path.GetPart("In2")).Auto, "B's handle stopped being automatic.");
        Require(drawing.SaveAsText().Contains("Path=\"C 1,2 a C a a C a a C a -1,-2\""), "The path's text after the drag: " + drawing.SaveAsText());
        drawing.ActionManager.Undo();
        Require(outA.Auto && inA.Auto && drawing.SaveAsText() == original, "Undo of the drag.");
        a.Selected = false;
        drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());

        // A a sharp corner, then smooth again
        Require(BezierPath.CanSharpenAnchor(a) && !BezierPath.CanSmoothAnchor(a), "The verbs offered at a smooth anchor.");
        a.SharpenAnchor();
        Require(!outA.Auto && outA.Offset == default && inA.Offset == default, "A is no sharp corner.");
        Require(BezierPath.CanSmoothAnchor(a) && !BezierPath.CanSharpenAnchor(a), "The verbs offered at a corner.");
        a.SmoothAnchor();
        Require(outA.Auto && inA.Auto && drawing.SaveAsText() == original, "Smooth automatically did not undo the corner.");
        drawing.ActionManager.Undo();
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == original, "Undo of the verbs.");

        // a split between automatic handles: the new anchor is smooth too
        var onPath = Factory.CreatePointOnFigure(drawing, path, path.GetPointFromParameter(0.5));
        Actions.Add(drawing, onPath);
        onPath.ConvertToPathAnchor();
        Require(path.AnchorCount == 5 && path.Handles.All(h => h.Auto), "The split made handles that aren't automatic.");
        drawing.Figures.CheckConsistency();
        drawing.ActionManager.Undo();
        Require(path.AnchorCount == 4, "Undo of the split.");
        drawing.Figures.CheckConsistency();

        // the tool: clicks make smooth anchors, a click on the slider right after ties the tension
        BezierPathCreator.LastSmoothing = DynamicGeometry.BezierPathSmoothing.Hobby;
        BezierPathCreator.LastTension = 1;
        var tool = new BezierPathCreator();
        drawing.Behavior = tool;
        foreach (var point in new[] { new Point(9, 0), new Point(6, 3), new Point(3, 0), new Point(6, -3), new Point(9, 0) })
        {
            Click(drawing, system.ToPhysical(point));
        }

        var made = drawing.Figures.OfType<BezierPath>().Single(p => p != path);
        Require(made.Closed && made.AnchorCount == 4, "The tool did not close a path of four anchors.");
        Close(((BezierPath.BezierPathHandle)made.GetPart("Out1")).Offset, new Point(0, 3 * kappa), "The tool's circle");
        Click(drawing, system.ToPhysical(new Point(-5, 5)));
        Require(made.GetSource(nameof(BezierPath.Tension)) == slider, "A click on the slider did not tie the tension.");
        Close(((BezierPath.BezierPathHandle)made.GetPart("Out1")).Offset, new Point(0, 3 * kappa / 2), "The tool's circle at tension 2");
        drawing.Figures.CheckConsistency();
    }

    /// <summary>
    /// Path CDF where D is on segment CE: deleting C takes D too, and the path, which would
    /// be left with one anchor. Path ABC whose tension is slider s: it can still be
    /// transformed; its image keeps it from being split (the image would stay the old
    /// curve); a press on a handle while the path is selected with A drags the handle, not
    /// the selection.
    /// </summary>
    static void BezierPathWithOthers()
    {
        var drawing = NewDrawing();
        var c = AddPoint(drawing, x: 0, y: 0);
        var e = AddPoint(drawing, x: 4, y: 0);
        var f = AddPoint(drawing, x: 2, y: 4);
        var ce = Factory.CreateSegment(drawing, c, e);
        Actions.Add(drawing, ce);
        var d = Factory.CreatePointOnFigure(drawing, ce, new Point(2, 0));
        Actions.Add(drawing, d);
        var zero = new Point[3];
        var cdf = BezierPath.Create(drawing, new IFigure[] { c, d, f }, zero, zero, closed: true, filled: false);
        Actions.Add(drawing, cdf);
        string before = drawing.SaveAsText();
        Actions.Remove(c);
        Require(!drawing.Figures.Contains(cdf), "Deleting C and D with it left a path of " + cdf.AnchorCount + " anchor.");
        drawing.Figures.CheckConsistency();
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of deleting C.");
        drawing.Figures.CheckConsistency();

        drawing = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <FreePoint Name="A" X="0" Y="0" />
                <FreePoint Name="B" X="6" Y="0" />
                <FreePoint Name="C" X="3" Y="5" />
                <Slider Name="s" X="-8" Y="6" Value="2" />
                <BezierPath Name="ABC" Closed="true" Tension="#3" Path="C 2,2 a C a a C a a">
                  <Dependency Name="A" />
                  <Dependency Name="B" />
                  <Dependency Name="C" />
                  <Dependency Name="s" />
                </BezierPath>
              </Figures>
            </Drawing>
            """);
        using var window = new TestWindow(drawing.Canvas);
        var system = drawing.CoordinateSystem;
        var path = drawing.Figures.OfType<BezierPath>().Single();
        var a = (FreePoint)Find(drawing, "A");
        Require(Transformer.CanBeSource(path), "A path whose tension is a slider can't be transformed.");

        // a press on A's handle, with A and the path selected
        drawing.Behavior = new Dragger();
        a.Selected = true;
        path.Selected = true;
        drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());
        Drag(drawing, system.ToPhysical(new Point(2, 2)), system.ToPhysical(new Point(3, 3)));
        Require(a.Coordinates == new Point(0, 0), "The press on the handle dragged the selection.");
        NearPoint(((BezierPath.BezierPathHandle)path.GetPart("Out1")).Offset, new Point(3, 3));
        drawing.ActionManager.Undo();
        drawing.Figures.ClearSelection();
        drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());

        // with an image, not split
        var mirror = Factory.CreateLineTwoPoints(drawing, new IFigure[] { AddPoint(drawing, x: 10, y: 0), AddPoint(drawing, x: 10, y: 1) });
        Actions.Add(drawing, mirror);
        foreach (var image in Transformer.CreateReflectedFigure(drawing, path, mirror))
        {
            Actions.Add(drawing, image);
        }

        var onPath = Factory.CreatePointOnFigure(drawing, path, path.GetPointFromParameter(0.5));
        Actions.Add(drawing, onPath);
        Require(!BezierPath.CanBecomeAnchor(onPath), "A path with an image can be split.");
        drawing.Figures.CheckConsistency();

        // an anchor deleted: the source goes on, the image goes, its helpers with it
        string imaged = drawing.SaveAsText();
        Actions.Remove(a);
        Require(drawing.Figures.Contains(path) && path.AnchorCount == 2, "Deleting A took the source.");
        Require(drawing.Figures.OfType<BezierPath>().Count() == 1, "The image stayed without A's image.");
        Require(!drawing.Figures.Any(figure => figure.Auxiliary), "The image's helpers stayed.");
        drawing.Figures.CheckConsistency();
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == imaged, "Undo of deleting A.");
        drawing.Figures.CheckConsistency();

        // a point that is a handle joined into an anchor of the same path: both at once
        var other = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <FreePoint Name="A" X="0" Y="0" />
                <FreePoint Name="B" X="6" Y="0" />
                <FreePoint Name="C" X="3" Y="5" />
                <FreePoint Name="P" X="2" Y="3" />
                <BezierPath Name="ABC" Closed="true" Path="C #3 a C a a C a a">
                  <Dependency Name="A" />
                  <Dependency Name="B" />
                  <Dependency Name="C" />
                  <Dependency Name="P" />
                </BezierPath>
              </Figures>
            </Drawing>
            """);
        var p = (FreePoint)Find(other, "P");
        var b = (IPoint)Find(other, "B");
        Require(PointSnapping.CanJoin(p, b), "P can't join B.");
        PointSnapping.Join(p, b);
        other.Figures.CheckConsistency();
        string joined = other.SaveAsText();
        var rejoined = ReadLgf(joined);
        Require(rejoined.LoadErrors == null && rejoined.SaveAsText() == joined, "The joined path did not round trip: " + joined);
        Actions.Remove((IFigure)b);
        other.Figures.CheckConsistency();
        other.ActionManager.Undo();
        Require(other.SaveAsText() == joined, "Undo of deleting B.");
        other.Figures.CheckConsistency();
    }

    static void Click(Drawing drawing, Point at, KeyModifiers modifiers)
    {
        var canvas = drawing.Canvas;
        using var pointer = new Pointer(Pointer.GetNextFreeId(), PointerType.Mouse, isPrimary: true);
        canvas.RaiseEvent(new PointerPressedEventArgs(
            canvas,
            pointer,
            canvas,
            at,
            timestamp: 0,
            new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.LeftButtonPressed),
            modifiers,
            clickCount: 1));
        canvas.RaiseEvent(new PointerReleasedEventArgs(
            canvas,
            pointer,
            canvas,
            at,
            timestamp: 1,
            new PointerPointProperties(RawInputModifiers.None, PointerUpdateKind.LeftButtonReleased),
            modifiers,
            MouseButton.Left));
    }

    static void Click(Drawing drawing, Point at)
    {
        var canvas = drawing.Canvas;
        using var pointer = new Pointer(Pointer.GetNextFreeId(), PointerType.Mouse, isPrimary: true);
        canvas.RaiseEvent(new PointerPressedEventArgs(
            canvas,
            pointer,
            canvas,
            at,
            timestamp: 0,
            new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.LeftButtonPressed),
            KeyModifiers.None,
            clickCount: 1));
        canvas.RaiseEvent(new PointerReleasedEventArgs(
            canvas,
            pointer,
            canvas,
            at,
            timestamp: 1,
            new PointerPointProperties(RawInputModifiers.None, PointerUpdateKind.LeftButtonReleased),
            KeyModifiers.None,
            MouseButton.Left));
    }

    static Cursor ToolCursor(string name)
    {
        return (Cursor)typeof(Behavior).GetField(name, System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static).GetValue(null);
    }

    static void MeasureToolsOnEachKind()
    {
        var drawing = NewDrawing();
        using var window = new TestWindow(drawing.Canvas);
        var system = drawing.CoordinateSystem;
        var hand = ToolCursor("HandCursor");
        var arrow = ToolCursor("ArrowCursor");
        var cross = ToolCursor("CrossCursor");
        var errors = new List<string>();
        drawing.UnhandledException += (sender, e) => errors.Add(e.Exception.ToString());
        Cursor CursorAt(double x, double y)
        {
            // (moves are taken one a frame, and the test window draws no frames: as if the last one had)
            typeof(Behavior)
                .GetField("movedThisFrame", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                .SetValue(drawing.Behavior, false);
            Hover(drawing, system.ToPhysical(new Point(x, y)));
            Require(errors.Count == 0, "A tool threw: " + string.Join("\n", errors));
            return drawing.Canvas.Cursor;
        }

        // a filled triangle, a line, a Bezier curve, an open path, a segment, apart from each other
        var triangle = Factory.CreatePolygon(drawing, new IFigure[] { AddPoint(drawing, x: -10, y: 0), AddPoint(drawing, x: -6, y: 0), AddPoint(drawing, x: -8, y: 4) });
        Actions.Add(drawing, triangle);
        var line = Factory.CreateLineTwoPoints(drawing, new IFigure[] { AddPoint(drawing, x: -10, y: -6), AddPoint(drawing, x: -9, y: -6) });
        Actions.Add(drawing, line);
        var bezier = (Bezier)Factory.CreateBezier(drawing, new IFigure[] { AddPoint(drawing, x: 0, y: 0), AddPoint(drawing, x: 0, y: 3), AddPoint(drawing, x: 3, y: 3), AddPoint(drawing, x: 3, y: 0) });
        Actions.Add(drawing, bezier);
        var corner = new BezierPath.HandleSpec(default, null, Auto: false);
        var path = BezierPath.Create(
            drawing,
            new IFigure[] { AddPoint(drawing, x: 6, y: 0), AddPoint(drawing, x: 9, y: 0), AddPoint(drawing, x: 9, y: 4) },
            new[] { corner, corner, corner },
            new[] { corner, corner, corner },
            holes: Array.Empty<IFigure>(),
            closed: false,
            filled: false);
        Actions.Add(drawing, path);
        var a = AddPoint(drawing, x: 0, y: -8);
        var segment = Factory.CreateSegment(drawing, new IFigure[] { AddPoint(drawing, x: 4, y: -8), AddPoint(drawing, x: 8, y: -8) });
        Actions.Add(drawing, segment);
        Near(path.Length, 7);
        Require(double.IsNaN(path.Perimeter), "An open path has a perimeter.");

        var length = new LengthMeasurementCreator();
        drawing.Behavior = length;
        Require(CursorAt(-8, 1) == hand, "Over a filled triangle the Length tool doesn't show the hand.");
        Require(CursorAt(-2, 8) == cross, "Over empty paper the Length tool shows no cross.");
        var overLine = CursorAt(-5, -6);
        Require(overLine == arrow, "Over a line the Length tool shows " + (overLine == hand ? "the hand" : overLine == cross ? "the cross" : "no cursor") + ", not the arrow of a point on it: "
            + PointPlacement.Find(drawing, new Point(-5, -6), snapToMidpoint: true)?.Kind);
        Require(CursorAt(1.5, 2.25) == hand, "Over a Bezier curve the Length tool doesn't show the hand.");
        Require(CursorAt(9, 2) == hand, "Over an open path the Length tool doesn't show the hand.");
        Require(CursorAt(-2, 8) == cross, "Over empty paper the Length tool shows no cross.");

        // the curve: its length, on it
        Click(drawing, system.ToPhysical(new Point(1.5, 2.25)));
        var curveLength = drawing.Figures.OfType<LengthMeasurement>().Last();
        Require(curveLength.Dependencies.Single() == bezier && curveLength.Quantity == LengthQuantity.Length, "A click on a Bezier curve measures " + curveLength.Construction);
        Near(curveLength.Length, new DynamicGeometry.Math.BezierInfo(new Point(0, 0), new Point(0, 3), new Point(3, 3), new Point(3, 0)).Length);

        // the open path: its length, which becomes the perimeter when the path is closed
        Click(drawing, system.ToPhysical(new Point(9, 2)));
        var pathLength = drawing.Figures.OfType<LengthMeasurement>().Last();
        Require(pathLength.Dependencies.Single() == path && pathLength.Quantity == LengthQuantity.Length && pathLength.Text == "7", "The open path says " + pathLength.Text);
        path.Closed = true;
        Require(pathLength.Quantity == LengthQuantity.Perimeter && pathLength.Exists && pathLength.Text == "P = 12", "The closed path says " + pathLength.Text);
        path.Closed = false;
        Require(pathLength.Quantity == LengthQuantity.Length && pathLength.Text == "7", "Opened again, the path says " + pathLength.Text);

        // a line is no length: a click puts a point on it, the first end of a distance
        int pointsBefore = drawing.Figures.OfType<IPoint>().Count();
        Click(drawing, system.ToPhysical(new Point(-5, -6)));
        Require(!drawing.Figures.OfType<LengthMeasurement>().Any(m => m.Dependencies.Contains(line)), "The line was measured.");
        Require(CursorAt(5, -8) == arrow, "With the first end down, a segment offers to measure itself.");
        Click(drawing, system.ToPhysical(new Point(5, -8)));
        var distance = drawing.Figures.OfType<LengthMeasurement>().Last();
        Require(distance.Quantity == LengthQuantity.Distance && distance.Dependencies[0] is PointOnFigure && distance.Dependencies[1] is PointOnFigure,
            "A distance from a line to a segment is " + distance.Quantity + " of " + distance.Construction);
        Require(drawing.Figures.OfType<IPoint>().Count() == pointsBefore + 2, "The distance left points behind.");
        Near(distance.Length, System.Math.Sqrt(10 * 10 + 2 * 2));

        // the Area tool: the hand over a filled shape, and its area
        var area = new AreaMeasurementCreator();
        drawing.Behavior = area;
        Require(CursorAt(-8, 1) == hand, "Over a filled triangle the Area tool doesn't show the hand.");
        Click(drawing, system.ToPhysical(new Point(-8, 1)));
        var triangleArea = drawing.Figures.OfType<AreaMeasurement>().Last();
        Require(triangleArea.Dependencies.Single() == triangle && triangleArea.Text == "8", "The triangle's area says " + triangleArea.Text);

        // a closed path is a shape to it (hit on a side: it isn't filled), an open one is not
        Require(CursorAt(9, 2) != hand, "Over an open path the Area tool shows the hand.");
        path.Closed = true;
        Require(CursorAt(9, 2) == hand, "Over a closed path the Area tool doesn't show the hand.");
        Click(drawing, system.ToPhysical(new Point(9, 2)));
        var pathArea = drawing.Figures.OfType<AreaMeasurement>().Last();
        Require(pathArea.Dependencies.Single() == path && pathArea.Text == "6", "The closed path's area says " + pathArea.Text);
        path.Closed = false;
        Require(!pathArea.Exists, "The area of a path opened and unfilled still shows.");

        // once points are listed, a click inside a shape is the next point
        Click(drawing, system.ToPhysical(new Point(0, -12)));
        Require(drawing.IsRecordingTransaction, "The first point of a list started nothing.");
        Click(drawing, system.ToPhysical(new Point(-8, 1)));
        Require(drawing.Figures.OfType<AreaMeasurement>().Count(m => m.Dependencies.Contains(triangle)) == 1, "A click in a shape threw the points away.");
        drawing.Behavior = new Dragger();
        drawing.Figures.CheckConsistency();
    }

    static void LengthPlaceThroughUndo()
    {
        var drawing = NewDrawing();
        using var window = new TestWindow(drawing.Canvas);

        // a radius dragged away, then measured as the diameter: undo puts the number back where it was dragged to
        var circle = Factory.CreateCircle(drawing, new IFigure[] { AddPoint(drawing, x: 0, y: 0), AddPoint(drawing, x: 2, y: 0) });
        Actions.Add(drawing, circle);
        var radius = Factory.CreateLengthMeasurement(drawing, new IFigure[] { circle });
        Actions.Add(drawing, radius);
        radius.RestorePlace(new Point(28.5, 18));
        string before = drawing.SaveAsText();
        var quantityRow = PropertyDiscoveryStrategy.CreateValueProvider(radius, nameof(LengthMeasurement.Quantity));
        Require(quantityRow is LengthMeasurement.QuantityValue, "The Measure row is a " + quantityRow.GetType().Name);
        Actions.SetProperty(drawing.ActionManager, quantityRow, value: LengthQuantity.Diameter);
        Require(radius.Text == "d = 4" && radius.Offset != new Point(28.5, 18), "The diameter is not placed anew.");
        string after = drawing.SaveAsText();
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the quantity moved the number: " + radius.Offset);
        drawing.ActionManager.Redo();
        Require(drawing.SaveAsText() == after, "Redo of the quantity put the number elsewhere.");

        // an open path's length dragged away, then the path closed: the number stays, before and after a reload
        var corner = new BezierPath.HandleSpec(default, null, Auto: false);
        var path = BezierPath.Create(
            drawing,
            new IFigure[] { AddPoint(drawing, x: 10, y: 0), AddPoint(drawing, x: 13, y: 0), AddPoint(drawing, x: 13, y: 4) },
            new[] { corner, corner, corner },
            new[] { corner, corner, corner },
            holes: Array.Empty<IFigure>(),
            closed: false,
            filled: false);
        Actions.Add(drawing, path);
        var pathLength = Factory.CreateLengthMeasurement(drawing, new IFigure[] { path });
        Actions.Add(drawing, pathLength);
        pathLength.RestorePlace(new Point(48, 48));
        string open = drawing.SaveAsText();
        Actions.SetProperty(drawing.ActionManager, PropertyDiscoveryStrategy.CreateValueProvider(path, nameof(BezierPath.Closed)), value: true);
        Require(pathLength.Text == "P = 12" && pathLength.Offset == new Point(48, 48), "Closing the path moved its number to " + pathLength.Offset);
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == open, "Undo of closing the path left something.");

        var reloaded = ReadLgf(open);
        var reloadedPath = reloaded.Figures.OfType<BezierPath>().Single();
        var reloadedLength = reloaded.Figures.OfType<LengthMeasurement>().Single(m => m.Dependencies.Contains(reloadedPath));
        reloadedPath.Closed = true;
        Require(reloadedLength.Offset == new Point(48, 48), "Read back and closed, the path's number went to " + reloadedLength.Offset);
        drawing.Figures.CheckConsistency();
    }

    static void LengthPrefixUndo()
    {
        var drawing = NewDrawing();
        var circle = Factory.CreateCircle(drawing, new IFigure[] { AddPoint(drawing, x: 0, y: 0), AddPoint(drawing, x: 2, y: 0) });
        Actions.Add(drawing, circle);
        var measurement = Factory.CreateLengthMeasurement(drawing, new IFigure[] { circle });
        Actions.Add(drawing, measurement);
        IValueProvider Row(string name) => PropertyDiscoveryStrategy.CreateValueProvider(measurement, name);

        // "d = " typed on the radius, then the diameter measured, then the prefix cleared and undone twice
        Actions.SetProperty(drawing.ActionManager, Row(nameof(LengthMeasurement.Prefix)), value: "d = ");
        Require(measurement.Text == "d = 2", "The typed prefix says " + measurement.Text);
        Actions.SetProperty(drawing.ActionManager, Row(nameof(LengthMeasurement.Quantity)), value: LengthQuantity.Diameter);
        string typed = drawing.SaveAsText();
        Require(typed.Contains("Prefix=\"d = \""), "The typed prefix is not kept.");
        Actions.SetProperty(drawing.ActionManager, Row(nameof(LengthMeasurement.Prefix)), value: "");
        Require(measurement.Text == "4", "Without a prefix the diameter says " + measurement.Text);
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == typed, "Undo of the prefix lost the typed one.");
        drawing.ActionManager.Undo();
        Require(measurement.Text == "d = 2", "Undo of the quantity says " + measurement.Text);
    }

    static void GeoGebraMeasurementsInExpressions()
    {
        var drawing = ReadGeoGebra(
            "<element type=\"point\" label=\"A\"><coords x=\"0\" y=\"0\" z=\"1\" /></element>"
            + "<element type=\"point\" label=\"B\"><coords x=\"2\" y=\"0\" z=\"1\" /></element>"
            + "<command name=\"Circle\"><input a0=\"A\" a1=\"B\" /><output a0=\"c\" /></command>"
            + "<element type=\"conic\" label=\"c\" />"
            + "<command name=\"Radius\"><input a0=\"c\" /><output a0=\"r\" /></command>"
            + "<element type=\"numeric\" label=\"r\"><value val=\"2\" /></element>"
            + "<element type=\"point\" label=\"Q\"><coords x=\"5\" y=\"0\" z=\"1\" /></element>"
            + "<element type=\"point\" label=\"R\"><coords x=\"7\" y=\"0\" z=\"1\" /></element>"
            + "<element type=\"point\" label=\"S\"><coords x=\"5\" y=\"2\" z=\"1\" /></element>"
            + "<command name=\"Polygon\"><input a0=\"Q\" a1=\"R\" a2=\"S\" /><output a0=\"t\" /></command>"
            + "<element type=\"polygon\" label=\"t\" />"
            + "<command name=\"Area\"><input a0=\"t\" /><output a0=\"a\" /></command>"
            + "<element type=\"numeric\" label=\"a\"><value val=\"2\" /></element>"
            + "<expression label=\"diameter\" exp=\"2 * r\" />"
            + "<element type=\"numeric\" label=\"diameter\"><value val=\"4\" /></element>"
            + "<expression label=\"P\" exp=\"(r, a)\" />"
            + "<element type=\"point\" label=\"P\"><coords x=\"2\" y=\"2\" z=\"1\" /></element>"
            + "<expression label=\"caption\" exp=\"&quot;Radius = &quot; + r\" />"
            + "<element type=\"text\" label=\"caption\"><startPoint x=\"0\" y=\"-4\" z=\"1\" /></element>"
            + "<command name=\"Circle\"><input a0=\"A\" a1=\"diameter\" /><output a0=\"big\" /></command>"
            + "<element type=\"conic\" label=\"big\" />");
        ((FreePoint)Find(drawing, "B")).MoveTo(new Point(3, 0));
        ((FreePoint)Find(drawing, "S")).MoveTo(new Point(5, 4));
        drawing.Recalculate();
        Near(((CircleBase)Find(drawing, "big")).Radius, 6);
        var p = (IPoint)Find(drawing, "P");
        Require(System.Math.Abs(p.Coordinates.X - 3) < 1e-9 && System.Math.Abs(p.Coordinates.Y - 4) < 1e-9, "P is at " + p.Coordinates);
        var caption = (LabelBase)Find(drawing, "caption");
        Require(caption.ProcessedText.Contains("3"), "The caption says " + caption.ProcessedText);
    }

    static void LengthMeasurements()
    {
        var drawing = NewDrawing();
        using var window = new TestWindow(drawing.Canvas);
        var system = drawing.CoordinateSystem;
        var center = AddPoint(drawing, x: 0, y: 0);
        var rim = AddPoint(drawing, x: 2, y: 0);
        var circle = Factory.CreateCircle(drawing, new IFigure[] { center, rim });
        Actions.Add(drawing, circle);
        string before = drawing.SaveAsText();

        // the tool: one click on the circle measures its radius, and the panel offers the others
        var tool = new LengthMeasurementCreator();
        drawing.Behavior = tool;
        Click(drawing, system.ToPhysical(new Point(0, 2)));
        Require(!drawing.IsRecordingTransaction, "The Length tool left a transaction open.");
        var measurement = drawing.Figures.OfType<LengthMeasurement>().Single();
        Require(measurement.Quantity == LengthQuantity.Radius, "A click on a circle measures its " + measurement.Quantity);
        Require(
            measurement.Quantities.SequenceEqual(new[] { LengthQuantity.Radius, LengthQuantity.Diameter, LengthQuantity.Circumference }),
            "A circle offers " + string.Join(", ", measurement.Quantities));
        Require(measurement.Text == "r = 2", "A circle of radius 2 says " + measurement.Text);
        Require(measurement.Exists && measurement.Coordinates.Exists(), "The number is nowhere.");
        Require(measurement.Title == "Radius" && measurement.Construction == "of circle " + circle.Name, "The title says " + measurement.Title + " " + measurement.Construction);
        Require(!drawing.SaveAsText().Contains("Quantity="), "The circle's first length is written out.");
        Require(measurement.Offers(nameof(LengthMeasurement.Quantity), LengthQuantity.Diameter)
            && !measurement.Offers(nameof(LengthMeasurement.Quantity), LengthQuantity.Perimeter), "The Measure row offers the wrong lengths.");

        // another length of it is one undo step, and the prefix follows
        Actions.SetProperty(drawing.ActionManager, PropertyDiscoveryStrategy.CreateValueProvider(measurement, nameof(LengthMeasurement.Quantity)), value: LengthQuantity.Circumference);
        Near(measurement.Length, 4 * System.Math.PI);
        Require(measurement.Text == "C = 12.57", "The circumference says " + measurement.Text);
        Actions.SetProperty(drawing.ActionManager, PropertyDiscoveryStrategy.CreateValueProvider(measurement, nameof(LengthMeasurement.Prefix)), value: "");
        Require(measurement.Text == "12.57", "Without a prefix it says " + measurement.Text);
        string after = drawing.SaveAsText();
        Require(after.Contains("Quantity=\"Circumference\"") && after.Contains("Prefix=\"\""), "The quantity or the prefix is not saved.");
        drawing.ActionManager.Undo();
        Require(measurement.Text == "C = 12.57", "Undo of the prefix says " + measurement.Text);
        drawing.ActionManager.Undo();
        Require(measurement.Quantity == LengthQuantity.Radius && measurement.Text == "r = 2", "Undo of the quantity says " + measurement.Text);
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the measurement left something.");
        drawing.ActionManager.Redo();
        drawing.ActionManager.Redo();
        drawing.ActionManager.Redo();
        Require(drawing.SaveAsText() == after, "Redo of the measurement changed it.");
        var reloaded = ReadLgf(after);
        Require(reloaded.LoadErrors == null && reloaded.SaveAsText() == after, "The measurement did not round trip.");
        Require(reloaded.Figures.OfType<LengthMeasurement>().Single().Text == "12.57", "Read back, the number changed.");

        // the circumference is a length other figures take, and expressions read it by name
        var label = Factory.CreateLabel(drawing);
        Actions.Add(drawing, label);
        label.Text = "[" + circle.Name + ".Perimeter] [" + circle.Name + ".Circumference] [" + circle.Name + ".Diameter]";
        Require(label.ProcessedText == "12.57 12.57 4.00", "The label says " + label.ProcessedText);
        var byRadius = Factory.CreateCircleByRadius(drawing, new IFigure[] { measurement, center });
        Actions.Add(drawing, byRadius);
        Near(byRadius.Radius, 4 * System.Math.PI);
        Require(byRadius.Construction.Contains("circumference of circle " + circle.Name), "The circle by the circumference says " + byRadius.Construction);

        // two clicks on points, a click on a segment
        var first = AddPoint(drawing, x: 0, y: -5);
        var second = AddPoint(drawing, x: 3, y: -9);
        Click(drawing, system.ToPhysical(new Point(0, -5)));
        Click(drawing, system.ToPhysical(new Point(3, -9)));
        var distance = drawing.Figures.OfType<LengthMeasurement>().Last();
        Require(distance.Quantity == LengthQuantity.Distance && distance.Text == "5", "Two points say " + distance.Text);
        Require(distance.Title == "Distance" && distance.Construction == first.Name + second.Name, "The distance's title says " + distance.Title + " " + distance.Construction);
        var side = Factory.CreateSegment(drawing, new IFigure[] { AddPoint(drawing, x: 10, y: -5), AddPoint(drawing, x: 10, y: -8) });
        Actions.Add(drawing, side);
        Click(drawing, system.ToPhysical(new Point(10, -6.5)));
        var length = drawing.Figures.OfType<LengthMeasurement>().Last();
        Require(length.Dependencies.Single() == side && length.Quantity == LengthQuantity.Length && length.Text == "3", "A segment says " + length.Text);
        Require(!length.CanEdit(nameof(LengthMeasurement.Quantity)), "A segment's one length can be switched.");

        // an arc: its length first, then its radius
        var arc = Factory.CreateArc(drawing, new IFigure[] { AddPoint(drawing, x: 20, y: -5), AddPoint(drawing, x: 22, y: -5), AddPoint(drawing, x: 20, y: -3) });
        Actions.Add(drawing, arc);
        Click(drawing, system.ToPhysical(new Point(20 + System.Math.Sqrt(2), -5 + System.Math.Sqrt(2))));
        var arcLength = drawing.Figures.OfType<LengthMeasurement>().Last();
        Require(arcLength.Dependencies.Single() == arc && arcLength.Quantity == LengthQuantity.ArcLength, "A click on an arc measures its " + arcLength.Quantity);
        Near(arcLength.Length, System.Math.PI);
        Require(arcLength.Coordinates.Distance(arc.ArcMiddle) < 2, "The arc's length is not next to the arc.");
        Require(arcLength.Quantities.SequenceEqual(new[] { LengthQuantity.ArcLength, LengthQuantity.Radius }), "An arc offers " + string.Join(", ", arcLength.Quantities));
        drawing.Behavior = new Dragger();

        // the other kinds
        var corners = new IFigure[] { AddPoint(drawing, x: 10, y: 0), AddPoint(drawing, x: 13, y: 0), AddPoint(drawing, x: 13, y: 3), AddPoint(drawing, x: 10, y: 3) };
        var square = Factory.CreatePolygon(drawing, corners);
        Actions.Add(drawing, square);
        Near(square.Perimeter, 12);
        var hexagonFigures = RegularPolygonBuilder.Create(drawing, AddPoint(drawing, x: 20, y: 0), AddPoint(drawing, x: 21, y: 0), sides: 6, onSide: false, out var hexagon);
        foreach (var figure in hexagonFigures)
        {
            Actions.Add(drawing, figure);
        }

        Near(hexagon.Perimeter, 6);
        var hexagonMeasurement = Factory.CreateLengthMeasurement(drawing, new IFigure[] { hexagon });
        Actions.Add(drawing, hexagonMeasurement);
        Near(hexagonMeasurement.Length, 6);
        Require(hexagonMeasurement.Quantity == LengthQuantity.Perimeter && hexagonMeasurement.Text == "P = 6", "The hexagon says " + hexagonMeasurement.Text);
        var ellipse = Factory.CreateEllipse(drawing, new IFigure[] { AddPoint(drawing, x: 30, y: 0), AddPoint(drawing, x: 33, y: 0), AddPoint(drawing, x: 30, y: 2) });
        Actions.Add(drawing, ellipse);
        Require(System.Math.Abs(ellipse.Perimeter - 15.865439589290586) < 1e-6, "An ellipse with axes 3 and 2 has the perimeter " + ellipse.Perimeter);
        var sector = Factory.CreateCircleSector(drawing, new IFigure[] { AddPoint(drawing, x: 40, y: 0), AddPoint(drawing, x: 42, y: 0), AddPoint(drawing, x: 40, y: 2) });
        Actions.Add(drawing, sector);
        Near(sector.Perimeter, System.Math.PI + 4);
        var sectorLength = Factory.CreateLengthMeasurement(drawing, new IFigure[] { sector });
        Actions.Add(drawing, sectorLength);
        Require(sectorLength.Quantity == LengthQuantity.ArcLength && !sector.IsUsedForArea(), "A sector's arc length keeps it from being an arc.");
        sectorLength.Quantity = LengthQuantity.Perimeter;
        Near(sectorLength.Length, System.Math.PI + 4);
        Require(sector.IsUsedForArea(), "A sector's perimeter lets it become an arc.");
        sectorLength.Quantity = LengthQuantity.Radius;
        Require(sectorLength.Text == "r = 2", "The sector's radius says " + sectorLength.Text);
        var segment =Factory.CreateCircleSegment(drawing, new IFigure[] { AddPoint(drawing, x: 50, y: 0), AddPoint(drawing, x: 52, y: 0), AddPoint(drawing, x: 48, y: 0) });
        Actions.Add(drawing, segment);
        Near(segment.Perimeter, 2 * System.Math.PI + 4);

        // a closed path through four points of the unit circle is the circle but for a hair; opened, it has none
        var automatic = new BezierPath.HandleSpec(default, null, Auto: true);
        var anchors = new IFigure[] { AddPoint(drawing, x: 61, y: 0), AddPoint(drawing, x: 60, y: 1), AddPoint(drawing, x: 59, y: 0), AddPoint(drawing, x: 60, y: -1) };
        var path = BezierPath.Create(
            drawing,
            anchors,
            new[] { automatic, automatic, automatic, automatic },
            new[] { automatic, automatic, automatic, automatic },
            holes: Array.Empty<IFigure>(),
            closed: true,
            filled: true);
        Actions.Add(drawing, path);
        Require(System.Math.Abs(path.Perimeter - 2 * System.Math.PI) < 0.01, "The path around the unit circle has the perimeter " + path.Perimeter);
        var pathMeasurement = Factory.CreateLengthMeasurement(drawing, new IFigure[] { path });
        Actions.Add(drawing, pathMeasurement);
        Require(pathMeasurement.Exists, "The closed path's perimeter doesn't exist.");
        path.Closed = false;
        Require(double.IsNaN(path.Perimeter) && pathMeasurement.Quantity == LengthQuantity.Length && pathMeasurement.Exists, "An open path's number is not its length.");
        path.Closed = true;
        Require(pathMeasurement.Quantity == LengthQuantity.Perimeter && pathMeasurement.Exists, "Closed again, the path's number is not its perimeter.");
        var sticking = Factory.CreateLengthMeasurement(drawing, new IFigure[] { path }, LengthQuantity.Perimeter);
        Actions.Add(drawing, sticking);
        path.Closed = false;
        Require(!sticking.Exists, "A perimeter asked for stays on an open path.");
        path.Closed = true;
        drawing.Figures.CheckConsistency();
    }

    static void SquareOnSegment()
    {
        var drawing = NewDrawing();
        var first = AddPoint(drawing, x: 0, y: 0);
        var second = AddPoint(drawing, x: 3, y: 0);
        var segment = Factory.CreateSegment(drawing, first, second);
        Actions.Add(drawing, segment);
        string before = drawing.SaveAsText();
        var creator = new SquareCreator();
        drawing.Behavior = creator;
        new FigureCreator.Dialog(creator) { X = "0", Y = "0" }.AddPoint();
        new FigureCreator.Dialog(creator) { X = "3", Y = "0" }.AddPoint();
        Require(!drawing.IsRecordingTransaction, "The square was not finished.");
        var onAB = drawing.Figures.OfType<Segment>().Count(s => s.Dependencies.Contains(first) && s.Dependencies.Contains(second));
        Require(onAB == 1, onAB + " segments on A and B.");
        var square = drawing.Figures.OfType<Polygon>().Single();
        Require(square.Dependencies.Count == 4 && square.Dependencies.All(d => d.Exists), "The square is not whole.");
        Near(((IPoint)square.Dependencies[2]).Coordinates.Y, expected: 3);

        // the other two corners are points by coordinates over the first two, nothing else
        var corners = square.Dependencies.Skip(2).Cast<PointByCoordinates>().Select(corner => corner.XExpression.Text + ", " + corner.YExpression.Text);
        string a = first.Name;
        string b = second.Name;
        Require(corners.SequenceEqual(new[] { $"{a}.Y + {b}.X - {b}.Y, -{a}.X + {b}.X + {b}.Y", $"{a}.X + {a}.Y - {b}.Y, -{a}.X + {a}.Y + {b}.X" }), "The corners say " + string.Join("; ", corners));
        Require(
            drawing.Figures.OfType<IPoint>().Count() == 4 && drawing.Figures.OfType<Segment>().Count() == 4 && drawing.Figures.OfType<Polygon>().Count() == 1,
            "The square made more than its corners, sides and polygon: " + string.Join(", ", drawing.Figures));
        string after = drawing.SaveAsText();
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the square changed the drawing.");
        drawing.ActionManager.Redo();
        Require(drawing.SaveAsText() == after, "Redo of the square changed the drawing.");
        drawing.Figures.CheckConsistency();
        Require(ReadLgf(after).LoadErrors == null, "The drawing with the square did not load.");
    }

    // a square on the top right and bottom right corners of a square drawn top to bottom:
    // the side between them runs bottom to top, and the second square came out on the first
    static void SquareOnReversedSide()
    {
        var drawing = NewDrawing();
        var creator = new SquareCreator();
        drawing.Behavior = creator;
        new FigureCreator.Dialog(creator) { X = "0", Y = "0" }.AddPoint();
        new FigureCreator.Dialog(creator) { X = "0", Y = "-3" }.AddPoint();
        var first = drawing.Figures.OfType<Polygon>().Single();
        var bottomRight = (IPoint)first.Dependencies[2];
        var topRight = (IPoint)first.Dependencies[3];
        Near(bottomRight.Coordinates.X, expected: 3);
        Near(topRight.Coordinates.X, expected: 3);
        new FigureCreator.Dialog(creator) { X = "3", Y = "0" }.AddPoint();
        new FigureCreator.Dialog(creator) { X = "3", Y = "-3" }.AddPoint();
        Require(!drawing.IsRecordingTransaction, "The second square was not finished.");
        var second = drawing.Figures.OfType<Polygon>().Single(p => p != first);
        Require(second.Dependencies[0] == topRight && second.Dependencies[1] == bottomRight, "The second square is not on the first one's corners.");
        Near(((IPoint)second.Dependencies[2]).Coordinates.X, expected: 6);
        Near(((IPoint)second.Dependencies[3]).Coordinates.X, expected: 6);
        drawing.Figures.CheckConsistency();
    }

    /// <summary>
    /// Squares grown into a lattice, each on a side of one before: a corner where a square
    /// has one already is that corner (worked out exactly, it is the same point), the sides
    /// to it are its sides, and nothing is there twice. The 2 x 2 of the bug report in the
    /// order it was drawn, a 3 x 3 row by row, and four squares round a point, the last one
    /// closing on the first square's free corner. Undone, redone and read again, the same.
    /// </summary>
    static void SquareLattice()
    {
        // AB, then CB, DC, CF
        CheckLattice(new[] { (0, 0, 1, 0), (1, 1, 1, 0), (0, 1, 1, 1), (1, 1, 2, 1) }, points: 9, sides: 12, squares: 4);

        // the bottom row to the right, each on the right side of the one before, then each
        // square on the top side of the one under it
        var clicks = new List<(int, int, int, int)> { (0, 0, 1, 0) };
        for (int i = 1; i < 3; i++)
        {
            clicks.Add((i, 1, i, 0));
        }

        for (int j = 1; j < 3; j++)
        {
            for (int i = 0; i < 3; i++)
            {
                clicks.Add((i, j, i + 1, j));
            }
        }

        CheckLattice(clicks, points: 16, sides: 24, squares: 9);

        // round B = (1, 0): the fourth square's corner is A
        CheckLattice(new[] { (0, 0, 1, 0), (1, 1, 1, 0), (2, 0, 1, 0), (1, -1, 1, 0) }, points: 9, sides: 12, squares: 4);

        static void CheckLattice(IEnumerable<(int X1, int Y1, int X2, int Y2)> clicks, int points, int sides, int squares)
        {
            var drawing = NewDrawing();
            var creator = new SquareCreator();
            drawing.Behavior = creator;
            string empty = drawing.SaveAsText();
            int steps = 0;
            foreach (var (x1, y1, x2, y2) in clicks)
            {
                new FigureCreator.Dialog(creator) { X = x1.ToString(), Y = y1.ToString() }.AddPoint();
                new FigureCreator.Dialog(creator) { X = x2.ToString(), Y = y2.ToString() }.AddPoint();
                Require(!drawing.IsRecordingTransaction, "A square was not finished.");
                steps++;
            }

            string what = squares + " squares";
            CheckNothingTwice(drawing, points, sides, squares, what);
            string done = drawing.SaveAsText();
            for (int i = 0; i < steps; i++)
            {
                drawing.ActionManager.Undo();
            }

            Require(drawing.SaveAsText() == empty, what + ": undo left " + drawing.SaveAsText());
            for (int i = 0; i < steps; i++)
            {
                drawing.ActionManager.Redo();
            }

            Require(drawing.SaveAsText() == done, what + ": redo changed the drawing.");
            drawing.Figures.CheckConsistency();
            var reread = ReadLgf(done);
            Require(reread.LoadErrors == null, what + ": the drawing did not read back.");
            CheckNothingTwice(reread, points, sides, squares, what + " read again");
        }
    }

    /// <summary>So many points, segments and polygons, and no two at one place</summary>
    static void CheckNothingTwice(Drawing drawing, int points, int sides, int polygons, string what)
    {
        static (double, double) Place(IPoint point) => (System.Math.Round(point.Coordinates.X, 6) + 0.0, System.Math.Round(point.Coordinates.Y, 6) + 0.0);
        var places = drawing.Figures.OfType<IPoint>().GroupBy(Place).ToList();
        var twice = places.Where(place => place.Count() > 1).Select(place => string.Join(" and ", place.Select(point => point.Name)) + " at " + place.Key).ToList();
        Require(twice.Count == 0, what + ": " + string.Join("; ", twice));
        Require(places.Count == points, what + ": " + places.Count + " points");
        var segments = drawing.Figures.OfType<Segment>()
            .GroupBy(segment => string.Join(" ", segment.Dependencies.Cast<IPoint>().Select(Place).OrderBy(place => place)))
            .ToList();
        Require(segments.All(segment => segment.Count() == 1), what + ": a side is drawn twice");
        Require(segments.Count == sides, what + ": " + segments.Count + " segments");
        Require(drawing.Figures.OfType<Polygon>().Count() == polygons, what + ": " + drawing.Figures.OfType<Polygon>().Count() + " polygons");
        drawing.Figures.CheckConsistency();
    }

    /// <summary>
    /// Only the same point is taken for a corner, never one there by chance: a free point, a
    /// point by coordinates typed at the place, the corner of a square on other free points.
    /// And the same square on the same two points again makes nothing, no undo step either.
    /// </summary>
    static void ChancePointsAreNoCorners()
    {
        var drawing = NewDrawing();
        var free = AddPoint(drawing, x: 1, y: 1);
        var typed = Factory.CreatePointByCoordinates(drawing, "0", "1");
        Actions.Add(drawing, typed);
        var creator = new SquareCreator();
        drawing.Behavior = creator;
        void Square(int x1, int y1, int x2, int y2)
        {
            new FigureCreator.Dialog(creator) { X = x1.ToString(), Y = y1.ToString() }.AddPoint();
            new FigureCreator.Dialog(creator) { X = x2.ToString(), Y = y2.ToString() }.AddPoint();
        }

        Square(0, 0, 1, 0);
        var square = drawing.Figures.OfType<Polygon>().Single();
        Require(!square.Dependencies.Contains(free) && !square.Dependencies.Contains(typed), "A point there by chance became a corner: " + string.Join(", ", square.Dependencies));

        // on new points above, the square comes down onto the first one's top corners
        Square(1, 2, 0, 2);
        int AtPlace(double x, double y) => drawing.Figures.OfType<IPoint>().Count(point => point.Coordinates.Distance(new Point(x, y)) < 1e-9);
        Require(AtPlace(1, 1) == 3 && AtPlace(0, 1) == 3, "Corners of squares on other free points were taken for each other.");

        string before = drawing.SaveAsText();
        int undoable = drawing.ActionManager.EnumUndoableActions().Count();
        Square(0, 0, 1, 0);
        Require(drawing.SaveAsText() == before, "The same square again changed the drawing: " + drawing.SaveAsText());
        Require(drawing.ActionManager.EnumUndoableActions().Count() == undoable, "The same square again left an undo step.");
        drawing.Figures.CheckConsistency();
    }

    /// <summary>
    /// The regular polygons that tile, around a polygon of their kind, each on a side of one
    /// before: six hexagons round a hexagon (a ring of 24 vertices and 30 sides, the vertices
    /// two neighbors share worked out with sqrt(3) exactly), six triangles round a point (the
    /// last closing on the first one's free vertex). Nothing is there twice.
    /// </summary>
    static void RegularTilings()
    {
        var drawing = NewDrawing();
        var center = Build(drawing, AddPoint(drawing, x: 0, y: 0), AddPoint(drawing, x: 1, y: 0), sides: 6);
        var ring = center.Dependencies.Cast<IPoint>().ToList();
        for (int i = 0; i < 6; i++)
        {
            Build(drawing, ring[(i + 1) % 6], ring[i], sides: 6);
        }

        CheckNothingTwice(drawing, points: 24, sides: 30, polygons: 7, "hexagons");

        drawing = NewDrawing();
        var middle = AddPoint(drawing, x: 0, y: 0);
        IPoint next = AddPoint(drawing, x: 1, y: 0);
        for (int i = 0; i < 6; i++)
        {
            next = (IPoint)Build(drawing, middle, next, sides: 3).Dependencies[2];
        }

        CheckNothingTwice(drawing, points: 7, sides: 12, polygons: 6, "triangles");

        static Polygon Build(Drawing drawing, IPoint first, IPoint second, int sides)
        {
            var figures = RegularPolygonBuilder.Create(drawing, first, second, sides, onSide: true, out var polygon);
            foreach (var figure in figures)
            {
                Actions.Add(drawing, figure);
            }

            return polygon;
        }
    }

    static void PasteBringsStyles()
    {
        var source = NewDrawing();
        var point = AddPoint(source, x: 1, y: 2);
        var big = new PointStyle() { Name = "1", Size = 15, Fill = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Colors.Red) };
        source.StyleManager.Add(big);
        point.Style = big;
        string copied = DrawingSerializer.WriteUsingXmlWriter(writer => new DrawingSerializer().WriteFiguresWithStyles(source, new IFigure[] { point }, writer));

        var target = NewDrawing();
        var clash = new LineStyle() { Name = "1", StrokeWidth = 5 };
        target.StyleManager.Add(clash);
        int styleCount = target.StyleManager.GetAllStyles().Count();
        string before = target.SaveAsText();
        target.PasteFromText(copied);
        var pasted = target.Figures.OfType<FreePoint>().Single();
        Require(pasted.Style is PointStyle { Size: 15 } style && style.Name != "1", "The copy did not keep its look: " + pasted.Style?.Name);
        Require(target.StyleManager["1"] == clash, "The drawing's own style 1 was replaced.");
        string after = target.SaveAsText();
        target.ActionManager.Undo();
        Require(target.StyleManager.GetAllStyles().Count() == styleCount, "Undo left the pasted style.");
        Require(target.SaveAsText() == before, "Undo of the paste changed the drawing.");
        target.ActionManager.Redo();
        Require(target.SaveAsText() == after, "Redo of the paste changed the drawing.");
        var reloaded = ReadLgf(after);
        Require(reloaded.Figures.OfType<FreePoint>().Single().Style is PointStyle { Size: 15 }, "The saved copy lost its look.");

        // into the drawing it came from: its own style, no second one
        int sourceStyles = source.StyleManager.GetAllStyles().Count();
        source.PasteFromText(copied);
        Require(source.StyleManager.GetAllStyles().Count() == sourceStyles, "A paste into the same drawing added a style.");
        Require(source.Figures.OfType<FreePoint>().All(p => p.Style == big), "The copy in the same drawing took another style.");
    }

    /// <summary>
    /// A point on a default style its drawing changed (Edit this style, not a new one) keeps
    /// its look pasted into another drawing, where the style comes along under a free name;
    /// pasted into its own drawing, it stays on that default
    /// </summary>
    static void PasteBringsChangedDefault()
    {
        var source = NewDrawing();
        var point = AddPoint(source, x: 1, y: 2);
        var edited = (PointStyle)point.Style;
        Require(edited.Name == StyleManager.FreePointStyleName, "A new point's style: " + edited.Name);
        edited.Size = 40;
        string copied = DrawingSerializer.WriteUsingXmlWriter(writer => new DrawingSerializer().WriteFiguresWithStyles(source, new IFigure[] { point }, writer));

        var target = NewDrawing();
        target.PasteFromText(copied);
        var pasted = target.Figures.OfType<FreePoint>().Single();
        Require(pasted.Style is PointStyle { Size: 40 } && pasted.Style.Name != StyleManager.FreePointStyleName, "The copy lost the changed default's look: " + pasted.Style?.Name);
        Require(target.StyleManager[StyleManager.FreePointStyleName] is PointStyle { Size: 10 }, "The drawing's own default was changed.");
        var reloaded = ReadLgf(target.SaveAsText());
        Require(reloaded.Figures.OfType<FreePoint>().Single().Style is PointStyle { Size: 40 }, "The saved copy lost its look.");

        int sourceStyles = source.StyleManager.GetAllStyles().Count();
        source.PasteFromText(copied);
        Require(source.StyleManager.GetAllStyles().Count() == sourceStyles, "A paste into the same drawing added a style.");
        Require(source.Figures.OfType<FreePoint>().All(p => p.Style == edited), "The copy in the same drawing left the default.");

        // the other way round: a point on the default as every drawing has it, pasted where
        // the default was made bigger, keeps its size, and the points there keep theirs
        var plain = NewDrawing();
        var plainPoint = AddPoint(plain, x: 1, y: 2);
        string plainCopied = DrawingSerializer.WriteUsingXmlWriter(writer => new DrawingSerializer().WriteFiguresWithStyles(plain, new IFigure[] { plainPoint }, writer));
        var resized = NewDrawing();
        var own = AddPoint(resized, x: 3, y: 4);
        var resizedDefault = (PointStyle)own.Style;
        resizedDefault.Size = 40;
        resized.PasteFromText(plainCopied);
        var copy = resized.Figures.OfType<FreePoint>().Single(p => p != own);
        Require(copy.Style is PointStyle { Size: 10 }, "A point on an unchanged default took the drawing's changed one: " + copy.Style?.Name);
        Require(own.Style == resizedDefault && resizedDefault.Size == 40, "The drawing's own point or default changed.");

        // and where the default is as every drawing has it, the copy takes it: no second style
        foreach (var unchanged in new[] { plain, NewDrawing() })
        {
            int styles = unchanged.StyleManager.GetAllStyles().Count();
            unchanged.PasteFromText(plainCopied);
            Require(unchanged.StyleManager.GetAllStyles().Count() == styles, "A copy on an unchanged default added a style.");
            Require(unchanged.Figures.OfType<FreePoint>().All(p => p.Style == unchanged.StyleManager[StyleManager.FreePointStyleName]), "A copy on an unchanged default left it.");
        }
    }

    /// <summary>
    /// Two styles of the copies renamed in one paste keep apart: red "1", which the drawing
    /// has in green, doesn't take "12", the name of the blue style pasted with it (renamed
    /// one after the other, both copies came out blue)
    /// </summary>
    static void PasteRenamesKeepStylesApart()
    {
        var red = Avalonia.Media.Colors.Red;
        var blue = Avalonia.Media.Colors.Blue;
        var source = NewDrawing();
        var redPoint = AddPoint(source, x: 0, y: 0);
        var bluePoint = AddPoint(source, x: 2, y: 0);
        var redStyle = new PointStyle() { Name = "1", Fill = new Avalonia.Media.SolidColorBrush(red) };
        var blueStyle = new PointStyle() { Name = "12", Fill = new Avalonia.Media.SolidColorBrush(blue) };
        source.StyleManager.Add(redStyle);
        source.StyleManager.Add(blueStyle);
        redPoint.Style = redStyle;
        bluePoint.Style = blueStyle;
        string copied = DrawingSerializer.WriteUsingXmlWriter(writer => new DrawingSerializer().WriteFiguresWithStyles(source, new IFigure[] { redPoint, bluePoint }, writer));

        var target = NewDrawing();
        target.StyleManager.Add(new PointStyle() { Name = "1", Fill = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Colors.Green) });
        target.PasteFromText(copied);
        Avalonia.Media.Color Fill(IFigure figure) => ((Avalonia.Media.SolidColorBrush)((PointStyle)figure.Style).Fill).Color;
        var pasted = target.Figures.OfType<FreePoint>().ToList();
        Require(pasted.Count == 2 && Fill(pasted[0]) == red && Fill(pasted[1]) == blue, "The pasted points are " + string.Join(", ", pasted.Select(p => p.Style.Name + " " + Fill(p))));
        var reloaded = ReadLgf(target.SaveAsText()).Figures.OfType<FreePoint>().ToList();
        Require(Fill(reloaded[0]) == red && Fill(reloaded[1]) == blue, "The saved copies lost their colors.");
    }

    /// <summary>
    /// Under Dark, a point style given a size and an emoji: the same look whichever was set
    /// first, and after the file is read again (the size went to the shape when it came
    /// first, and the emoji showed at its own size while the grid said 60)
    /// </summary>
    static void ThemeOverrideOrder()
    {
        string dark = AppTheme.Dark.Name;
        const string star = "⭐";
        foreach (bool sizeFirst in new[] { true, false })
        {
            var style = new PointStyle() { Name = "Star" };
            if (sizeFirst)
            {
                style.SetOverride(dark, nameof(PointStyle.Size), 60.0);
                style.SetOverride(dark, nameof(PointStyle.Character), star);
            }
            else
            {
                style.SetOverride(dark, nameof(PointStyle.Character), star);
                style.SetOverride(dark, nameof(PointStyle.Size), 60.0);
            }

            Require(style.Resolve(dark) is PointStyle { Character: star, Size: 60 }, $"Size {(sizeFirst ? "before" : "after")} the emoji: the emoji at {((PointStyle)style.Resolve(dark)).Size}.");
            var drawing = NewDrawing();
            drawing.StyleManager.Add(style);
            AddPoint(drawing, x: 0, y: 0).Style = style;
            var reloaded = (FigureStyle)ReadLgf(drawing.SaveAsText()).Figures.OfType<FreePoint>().Single().Style;
            Require(reloaded.Resolve(dark) is PointStyle { Character: star, Size: 60 }, "The file read again shows the emoji at " + ((PointStyle)reloaded.Resolve(dark)).Size);
        }
    }

    /// <summary>
    /// A vector whose start is the midpoint of its end and another point has no Direction to
    /// type (turning the end moved the start, and the vector came out at another angle and
    /// shorter); on an ordinary vector, undo of a run of directions puts the end back exactly
    /// </summary>
    static void VectorDirectionEdit()
    {
        var drawing = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <FreePoint Name="C" X="-2" Y="0" />
                <FreePoint Name="B" X="2" Y="0" />
                <MidPoint Name="A">
                  <Dependency Name="C" />
                  <Dependency Name="B" />
                </MidPoint>
                <Vector Name="AB">
                  <Dependency Name="A" />
                  <Dependency Name="B" />
                </Vector>
                <FreePoint Name="P" X="0" Y="-3" />
                <FreePoint Name="Q" X="2" Y="-3" />
                <Vector Name="PQ">
                  <Dependency Name="P" />
                  <Dependency Name="Q" />
                </Vector>
              </Figures>
            </Drawing>
            """);
        var dependent = (DynamicGeometry.Vector)Find(drawing, "AB");
        Require(!dependent.CanEdit(nameof(DynamicGeometry.Vector.Angle)), "A vector from a midpoint of its end offers a direction to type.");

        var vector = (DynamicGeometry.Vector)Find(drawing, "PQ");
        var end = (IPoint)Find(drawing, "Q");
        Require(vector.CanEdit(nameof(DynamicGeometry.Vector.Angle)), "An ordinary vector offers no direction to type.");
        string before = drawing.SaveAsText();
        Set(drawing, vector, nameof(DynamicGeometry.Vector.Angle), 90.0);
        NearPoint(end.Coordinates, new Point(0, -1));
        Set(drawing, vector, nameof(DynamicGeometry.Vector.Angle), 180.0);
        string after = drawing.SaveAsText();
        drawing.ActionManager.Undo();
        NearPoint(end.Coordinates, new Point(0, -1));
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the directions didn't put the end back: " + end.Coordinates);
        drawing.ActionManager.Redo();
        drawing.ActionManager.Redo();
        Require(drawing.SaveAsText() == after, "Redo of the directions changed the drawing.");
    }

    /// <summary>
    /// A point inside a filled path with a hole that reaches past the path: dragged into the
    /// hole, it stops at the hole's edge, not on the path's own outline inside the hole; and
    /// where the hole covers the whole path, the point is nowhere
    /// </summary>
    static void PointInsidePathWithHoles()
    {
        // the square from (-4, -4) to (4, 4), filled, with rectangular holes, and P inside it
        static string WithHoles(params (double Left, double Bottom, double Right, double Top)[] holes)
        {
            var figures = new StringBuilder();
            var outer = new StringBuilder();
            void Rectangle(string name, (double Left, double Bottom, double Right, double Top) box, string attributes, StringBuilder extra)
            {
                var corners = new[] { (box.Left, box.Bottom), (box.Right, box.Bottom), (box.Right, box.Top), (box.Left, box.Top) };
                for (int i = 0; i < 4; i++)
                {
                    figures.Append($"<FreePoint Name=\"{name}{i}\" X=\"{corners[i].Item1}\" Y=\"{corners[i].Item2}\" />");
                }

                figures.Append($"<BezierPath Name=\"{name}\" Closed=\"true\" {attributes} Path=\"L L L L\">");
                for (int i = 0; i < 4; i++)
                {
                    figures.Append($"<Dependency Name=\"{name}{i}\" />");
                }

                figures.Append(extra).Append("</BezierPath>");
            }

            for (int i = 0; i < holes.Length; i++)
            {
                Rectangle("hole" + i, holes[i], "", new StringBuilder());
                outer.Append($"<Dependency Name=\"hole{i}\" />");
            }

            Rectangle("outer", (-4, -4, 4, 4), "Filled=\"true\"", outer);
            figures.Append("<PointInRegion Name=\"P\" X=\"3\" Y=\"3.5\" OffsetX=\"3\" OffsetY=\"3.5\"><Dependency Name=\"outer\" /></PointInRegion>");
            return "<Drawing Version=\"1\"><Figures>" + figures + "</Figures></Drawing>";
        }

        // a hole reaching past the top: the point stops at the hole's edge, not on the
        // square's top inside the hole
        var drawing = ReadLgf(WithHoles((-2, -3, 2, 5)));
        var point = (PointInRegion)Find(drawing, "P");
        foreach (var target in new[] { new Point(0, 3.5), new Point(0.5, 4.5) })
        {
            point.MoveTo(target);
            var place = point.Coordinates;
            Require(System.Math.Abs(System.Math.Abs(place.X) - 2) < 1e-9 && place.Y <= 4 + 1e-9, $"Dragged to {target}, the point is at {place}, in the hole.");
        }

        // two holes side by side: their seam is no edge
        var seam = ReadLgf(WithHoles((-2, -2, 0, 2), (0, -2, 2, 2)));
        var onSeam = (PointInRegion)Find(seam, "P");
        onSeam.MoveTo(new Point(0, 0.5));
        Require(System.Math.Abs(onSeam.Coordinates.Distance(new Point(0, 0.5)) - 1.5) < 1e-9, "Dragged onto the seam of two holes, the point is at " + onSeam.Coordinates);

        // a hole that covers the square, exactly or with room: the point is nowhere
        foreach (var hole in new[] { (-4.0, -4.0, 4.0, 4.0), (-5.0, -5.0, 5.0, 5.0) })
        {
            var covered = ReadLgf(WithHoles(hole));
            covered.Recalculate();
            Require(!Find(covered, "P").Exists, $"A point inside a path its hole {hole} covers exists.");
        }
    }

    /// <summary>
    /// A sector or a circular segment with a point kept inside has no Convert to arc: a bare
    /// arc has no inside, and the point stopped existing for good. Between kinds with an
    /// inside the conversion stays.
    /// </summary>
    static void NoArcUnderPointInside()
    {
        var drawing = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <FreePoint Name="A" X="0" Y="0" />
                <FreePoint Name="B" X="3" Y="0" />
                <FreePoint Name="C" X="0" Y="3" />
                <CircleSector Name="c">
                  <Dependency Name="A" />
                  <Dependency Name="B" />
                  <Dependency Name="C" />
                </CircleSector>
                <PointInRegion Name="D" X="1" Y="1" OffsetX="1" OffsetY="1">
                  <Dependency Name="c" />
                </PointInRegion>
                <CircleSegment Name="d">
                  <Dependency Name="A" />
                  <Dependency Name="B" />
                  <Dependency Name="C" />
                </CircleSegment>
                <PointInRegion Name="E" X="2" Y="2" OffsetX="2" OffsetY="2">
                  <Dependency Name="d" />
                </PointInRegion>
              </Figures>
            </Drawing>
            """);
        foreach (var name in new[] { "c", "d" })
        {
            var region = (IConditionalProperties)Find(drawing, name);
            Require(!region.CanEdit("ConvertToArc"), name + " offers Convert to arc with a point inside it.");
        }

        Require(((IConditionalProperties)Find(drawing, "c")).CanEdit("ConvertToSegment"), "The sector no longer offers Convert to segment.");
    }

    /// <summary>
    /// A text label that says a number, edited: what is built on its value follows at once
    /// (a circle with it for its radius, a point at its value), and goes back on undo. They
    /// stayed until the next recalculation, and undo left them at the new value.
    /// </summary>
    static void NumericLabelEditMovesDependents()
    {
        var drawing = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <Label Name="radius" Text="[2]" X="3" Y="3" />
                <FreePoint Name="O" X="0" Y="0" />
                <CircleByRadius Name="c">
                  <Dependency Name="radius" />
                  <Dependency Name="O" />
                </CircleByRadius>
                <PointByCoordinates Name="P" X="radius.Value" Y="-3">
                  <Dependency Name="radius" />
                </PointByCoordinates>
              </Figures>
            </Drawing>
            """);
        var label = Find(drawing, "radius");
        var circle = (CircleByRadius)Find(drawing, "c");
        var point = (IPoint)Find(drawing, "P");
        double width = circle.Shape.Width;
        Set(drawing, label, "Text", "[4]");
        NearPoint(point.Coordinates, new Point(4, -3));
        Require(circle.Shape.Width > 1.5 * width, "The circle was not redrawn at the new radius: " + circle.Shape.Width + " wide, from " + width);
        drawing.ActionManager.Undo();
        NearPoint(point.Coordinates, new Point(2, -3));
        Near(circle.Shape.Width, width);

        // no longer a number: its value is 0 (a caption's), and the point follows that too
        Set(drawing, label, "Text", "radius");
        NearPoint(point.Coordinates, new Point(0, -3));
    }

    /// <summary>
    /// A number that expressions name, renamed pi or e in any case, or x where a function
    /// names it, is said in backticks: the texts said the constant once, and the file read
    /// again took them for it. Read again, they still mean the number; undo of the rename
    /// puts the plain name back.
    /// </summary>
    static void ReservedNamesForNumbers()
    {
        var drawing = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <Slider Name="a" X="-3" Y="2" Value="2" Maximum="5" />
                <PointByCoordinates Name="P" X="a" Y="0">
                  <Dependency Name="a" />
                </PointByCoordinates>
                <Slider Name="b" X="-3" Y="4" Value="1" Maximum="5" />
                <FunctionGraph Name="f" Function="b * x">
                  <Dependency Name="b" />
                </FunctionGraph>
              </Figures>
            </Drawing>
            """);
        var a = Find(drawing, "a");
        var point = (PointByCoordinates)Find(drawing, "P");
        foreach (var name in new[] { "pi", "PI", "e", "E" })
        {
            Require(Rename(a, name) == null && a.Name == name, "A slider an expression names was not renamed " + name + ".");
            Require(point.XExpression.Text == "`" + name + "`", "The point's X says " + point.XExpression.Text);
            var reread = (IPoint)Find(ReadLgf(drawing.SaveAsText()), "P");
            Near(reread.Coordinates.X, expected: 2);
        }

        // a plain name again needs no backticks; undo puts them back
        Set(drawing, a, nameof(IFigure.Name), value: "s");
        Require(point.XExpression.Text == "s", "The point's X says " + point.XExpression.Text);
        drawing.ActionManager.Undo();
        Require(a.Name == "E" && point.XExpression.Text == "`E`", "Undo of the rename to s left " + a.Name + " and " + point.XExpression.Text);

        var b = Find(drawing, "b");
        var graph = (FunctionGraph)Find(drawing, "f");
        Require(Rename(b, "x") == null && b.Name == "x", "A slider a function names was not renamed x.");
        Require(graph.FunctionText == "`x` * x", "The function says " + graph.FunctionText);
        var rereadGraph = (FunctionGraph)Find(ReadLgf(drawing.SaveAsText()), "f");
        Near(rereadGraph.Function(3), expected: 3);

        static string Rename(IFigure figure, string name)
        {
            var editor = new NameEditor { Value = PropertyDiscoveryStrategy.CreateValueProvider(figure, nameof(IFigure.Name)) };
            editor.TextBox.Text = name;
            editor.TextBox.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter });
            return string.IsNullOrEmpty(editor.ErrorText) ? null : editor.ErrorText;
        }
    }

    /// <summary>
    /// Any name can be said in an expression, in backticks where it isn't a name of the
    /// language (`Drag me!`, a backtick in it doubled): the figure called exactly that,
    /// never a constant, two points, a function or a figure differing by case. A label's
    /// [...] part takes a ] inside backticks along. Renames write the name so, and the file
    /// read again means the same figures.
    /// </summary>
    static void QuotedNames()
    {
        Require(Scanner.NameInExpression("A'") == "A'" && Scanner.NameInExpression("P_1") == "P_1", "A name was put in backticks.");
        Require(Scanner.NameInExpression("Drag me!") == "`Drag me!`" && Scanner.NameInExpression("q`r") == "`q``r`", "Scanner.NameInExpression");
        Require(Scanner.NameInExpression("e") == "`e`" && Scanner.NameInExpression("Pi") == "`Pi`" && Scanner.NameInExpression("x") == "`x`", "A reserved name was said bare.");
        Require(Scanner.NameInExpression(null) == "``", "A figure without a name");

        var drawing = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <FreePoint Name="Drag me!" X="3" Y="4" />
                <FreePoint Name="a]b" X="1" Y="0" />
                <FreePoint Name="q`r" X="0" Y="2" />
                <FreePoint Name="e" X="5" Y="0" />
                <FreePoint Name="B" X="0" Y="0" />
                <Slider Name="pi" X="-3" Y="2" Value="2" Maximum="5" />
              </Figures>
            </Drawing>
            """);
        foreach (var (text, expected) in new[]
        {
            ("`Drag me!`.X + `a]b`.X", 4.0),
            ("`q``r`.Y * `pi`", 4.0),
            ("dist(`Drag me!`, B) + `e`.X", 10.0),
            ("pi", System.Math.PI),
            ("`pi` + e.X", 7.0),
        })
        {
            var result = Compiler.Instance.CompileExpression(drawing, text, figure => true);
            Require(result.IsSuccess, text + ": " + result);
            Near(result.Expression(), expected);
        }

        foreach (var text in new[] { "`drag me!`.X", "`Drag me!", "``.X", "`B`(1)", "`AB`", "`sin`(1)" })
        {
            Require(!Compiler.Instance.CompileExpression(drawing, text, figure => true).IsSuccess, text + " compiled.");
        }

        var label = Factory.CreateLabel(drawing);
        Actions.Add(drawing, label);
        label.Text = "[`a]b`.X + `Drag me!`.Y] and [B.X]";
        Require(label.ProcessedText == "5.00 and 0.00", "The label says " + label.ProcessedText);

        // renamed to a name of the language and back into backticks
        var point = Find(drawing, "Drag me!");
        Set(drawing, point, nameof(IFigure.Name), value: "Pole");
        Require(label.Text == "[`a]b`.X + Pole.Y] and [B.X]", "The label says " + label.Text);
        Set(drawing, Find(drawing, "B"), nameof(IFigure.Name), value: "my ]point");
        Require(label.Text == "[`a]b`.X + Pole.Y] and [`my ]point`.X]", "The label says " + label.Text);
        var reread = ReadLgf(drawing.SaveAsText()).Figures.OfType<DynamicGeometry.Label>().Single();
        Require(reread.ProcessedText == "5.00 and 0.00", "The label read again says " + reread.ProcessedText);
        drawing.ActionManager.Undo();
        drawing.ActionManager.Undo();
        Require(label.Text == "[`a]b`.X + `Drag me!`.Y] and [B.X]", "Undo left the label saying " + label.Text);
        drawing.Figures.CheckConsistency();
    }

    /// <summary>
    /// A segment whose length an expression reads ([AB.Length]) offers no Convert to ray or
    /// line. Converted all the same, the expression no longer compiles and gives no value:
    /// the point is nowhere at once, as it is when the file is read again (it went on at
    /// its old place, through what had been compiled for the segment that left), and undo
    /// brings it back. A name the expressions say in backticks (`base side`.Length, given by a
    /// rename or typed so) is read all the same: a regular expression over the text missed it,
    /// and the conversions came back once the segment was renamed. So is a label's [...] part.
    /// </summary>
    static void ConversionKeepsExpressionsWhole()
    {
        var drawing = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <FreePoint Name="A" X="0" Y="0" />
                <FreePoint Name="B" X="3" Y="0" />
                <Segment Name="AB">
                  <Dependency Name="A" />
                  <Dependency Name="B" />
                </Segment>
                <PointByCoordinates Name="P" X="AB.Length" Y="-1">
                  <Dependency Name="AB" />
                </PointByCoordinates>
                <FreePoint Name="C" X="0" Y="2" />
                <FreePoint Name="D" X="2" Y="2" />
                <Segment Name="CD">
                  <Dependency Name="C" />
                  <Dependency Name="D" />
                </Segment>
                <Label Name="half" Text="half of it: [CD.Length / 2]" X="0" Y="4" />
              </Figures>
            </Drawing>
            """);
        var segment = (Segment)Find(drawing, "AB");
        Require(!segment.CanEdit("ConvertToRay") && !segment.CanEdit("ConvertToLine"), "A segment whose length an expression reads offers to become a ray or a line.");

        segment.ConvertToRay();
        Require(!Find(drawing, "P").Exists, "The point at (AB.Length, -1) exists after AB became a ray.");
        Require(!Find(ReadLgf(drawing.SaveAsText()), "P").Exists, "Read again, the point exists.");
        drawing.ActionManager.Undo();
        var point = (IPoint)Find(drawing, "P");
        Require(((IFigure)point).Exists, "Undo of the conversion left the point nowhere.");
        NearPoint(point.Coordinates, new Point(3, -1));
        drawing.Figures.CheckConsistency();

        segment = (Segment)Find(drawing, "AB");
        var holder = (IRenamableExpressions)point;
        foreach (var name in new[] { "base side", "my `side" })
        {
            Set(drawing, segment, nameof(IFigure.Name), value: name);
            Require(
                !segment.CanEdit("ConvertToRay") && !segment.CanEdit("ConvertToLine"),
                "Named " + name + ", the segment P reads (X = " + holder.ExpressionTexts[0] + ") offers to become a ray or a line.");
        }

        Set(drawing, segment, nameof(IFigure.Name), value: "AB");
        holder.ExpressionTexts = new[] { "`AB`.Length", "-1" };
        holder.RebindExpressions();
        Require(((IFigure)point).Exists, "P at (`AB`.Length, -1) is nowhere.");
        Require(!segment.CanEdit("ConvertToRay"), "Read as `AB`.Length, the segment offers to become a ray.");

        var other = (Segment)Find(drawing, "CD");
        var label = (DynamicGeometry.Label)Find(drawing, "half");
        Require(!other.CanEdit("ConvertToRay"), "A segment a label's [CD.Length / 2] reads offers to become a ray.");
        Set(drawing, other, nameof(IFigure.Name), value: "other side");
        Require(label.Text == "half of it: [`other side`.Length / 2]", "The label says " + label.Text);
        Require(!other.CanEdit("ConvertToRay"), "Renamed, a segment a label reads offers to become a ray.");
        drawing.Figures.CheckConsistency();
    }

    /// <summary>
    /// A segment's Free length and Free direction make its fixed end a free point when nothing
    /// else holds it - not while an expression reads its Distance or Direction, which a free
    /// point has not (the point at (B.Distance, 3) went nowhere, the label said "Could not find
    /// property"; the point's own Free boxes refused the same). Freeing one of two fixed
    /// quantities keeps the translated point, and stays offered. The same for Fix length of
    /// a point on a line read as B.Parameter (it became a translated point) and Convert to path
    /// anchor of a point on a path read so (a free point); a length typed once still moves it.
    /// </summary>
    static void PointKindKeepsExpressionsWhole()
    {
        foreach (var quantity in new[] { "Distance", "Direction" })
        {
            var drawing = ReadLgf("""
                <Drawing Version="1">
                  <Figures>
                    <FreePoint Name="A" X="0" Y="0" />
                    <FreePoint Name="B" X="4" Y="1" />
                    <Segment Name="AB">
                      <Dependency Name="A" />
                      <Dependency Name="B" />
                    </Segment>
                  </Figures>
                </Drawing>
                """);
            var segment = (Segment)Find(drawing, "AB");
            string fix = quantity == "Distance" ? "FixLength" : "FixDirection";
            string free = quantity == "Distance" ? "FreeLength" : "FreeDirection";
            if (quantity == "Distance")
            {
                segment.FixLength();
            }
            else
            {
                segment.FixDirection();
            }

            Require(Find(drawing, "B") is TranslatedPoint, fix + " left B a " + Find(drawing, "B").GetType().Name);
            var point = Factory.CreatePointByCoordinates(drawing, "B." + quantity, "3");
            Actions.Add(drawing, point);
            Require(point.Exists, "The point at (B." + quantity + ", 3) is nowhere.");
            foreach (var name in new[] { "B", "end point" })
            {
                Set(drawing, Find(drawing, "B"), nameof(IFigure.Name), value: name);
                Require(!segment.CanEdit(free), "With B named " + name + " and read as " + point.XExpression.Text + ", the segment offers " + free + ".");
                if (quantity == "Distance")
                {
                    segment.FreeLength();
                }
                else
                {
                    segment.FreeDirection();
                }

                Require(Find(drawing, name) is TranslatedPoint && point.Exists, free + " took away what " + point.XExpression.Text + " reads.");
                Set(drawing, Find(drawing, name), nameof(IFigure.Name), value: "B");
            }

            // both fixed: freeing this one lets go of it alone, and the point stays what it is
            if (quantity == "Distance")
            {
                segment.FixDirection();
            }
            else
            {
                segment.FixLength();
            }

            Require(segment.CanEdit(free), "With both fixed, the segment offers no " + free + ".");
            if (quantity == "Distance")
            {
                segment.FreeLength();
            }
            else
            {
                segment.FreeDirection();
            }

            Require(Find(drawing, "B") is TranslatedPoint && point.Exists, free + " of one of two fixed quantities took away what " + point.XExpression.Text + " reads.");
            drawing.Figures.CheckConsistency();
        }

        var onLine = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <FreePoint Name="A" X="0" Y="0" />
                <FreePoint Name="C" X="4" Y="0" />
                <LineTwoPoints Name="AC">
                  <Dependency Name="A" />
                  <Dependency Name="C" />
                </LineTwoPoints>
                <PointOnFigure Name="B" X="2" Y="0" Parameter="0.5">
                  <Dependency Name="AC" />
                </PointOnFigure>
                <Segment Name="AB">
                  <Dependency Name="A" />
                  <Dependency Name="B" />
                </Segment>
                <PointByCoordinates Name="P" X="B.Parameter" Y="3">
                  <Dependency Name="B" />
                </PointByCoordinates>
                <FreePoint Name="D" X="0" Y="4" />
                <FreePoint Name="E" X="4" Y="4" />
                <FreePoint Name="F" X="2" Y="7" />
                <BezierPath Name="DEF" Closed="true" Path="L L L">
                  <Dependency Name="D" />
                  <Dependency Name="E" />
                  <Dependency Name="F" />
                </BezierPath>
                <PointOnFigure Name="Q" X="2" Y="4" Parameter="0.5">
                  <Dependency Name="DEF" />
                </PointOnFigure>
                <PointByCoordinates Name="R" X="Q.Parameter" Y="-3">
                  <Dependency Name="Q" />
                </PointByCoordinates>
              </Figures>
            </Drawing>
            """);
        var side = (Segment)Find(onLine, "AB");
        var p = (PointByCoordinates)Find(onLine, "P");
        Require(p.Exists && side.CanEdit("Length"), "The point at (B.Parameter, 3) is nowhere, or AB has no length to type.");
        Require(!side.CanEdit("FixLength"), "AB offers Fix length while P reads B.Parameter.");
        side.FixLength();
        Require(Find(onLine, "B") is PointOnFigure && p.Exists, "Fix length took away B.Parameter.");
        ((IRenamableExpressions)p).ExpressionTexts = new[] { "B.X", "3" };
        ((IRenamableExpressions)p).RebindExpressions();
        Require(side.CanEdit("FixLength"), "AB offers no Fix length while P reads only B.X.");
        side.FixLength();
        Require(Find(onLine, "B") is TranslatedPoint && p.Exists, "Fix length with P on B.X: B is a " + Find(onLine, "B").GetType().Name);

        var q = (PointOnFigure)Find(onLine, "Q");
        var r = (PointByCoordinates)Find(onLine, "R");
        Require(q.Exists && r.Exists, "The point on the path or the point at (Q.Parameter, -3) is nowhere.");
        Require(!BezierPath.CanBecomeAnchor(q), "Q offers Convert to path anchor while R reads Q.Parameter.");
        ((IRenamableExpressions)r).ExpressionTexts = new[] { "Q.X", "-3" };
        ((IRenamableExpressions)r).RebindExpressions();
        Require(BezierPath.CanBecomeAnchor(q), "Q offers no Convert to path anchor while R reads only Q.X.");
        onLine.Figures.CheckConsistency();
    }

    /// <summary>
    /// What defines a figure is no row of a selection of several: not the X and Y of points
    /// by coordinates, a function, a label's text, the parts of a line's or a circle's
    /// equation (one text made them all one figure; the X row threw, the function row took
    /// "sin(" into both graphs and their file). How they look and Locked still are. Behind
    /// that, the editors check text against each figure of a row of several, should one be
    /// offered again: what compiles for each goes into each, what doesn't stays in the box.
    /// </summary>
    static void SharedExpressionRows()
    {
        var drawing = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <PointByCoordinates Name="A" X="-2" Y="1" />
                <PointByCoordinates Name="B" X="2" Y="-1" />
                <FunctionGraph Name="f" Function="x" />
                <FunctionGraph Name="g" Function="-x" />
                <Label Name="first" Text="one" X="0" Y="3" />
                <Label Name="second" Text="two" X="0" Y="4" />
                <CircleByEquation Name="c" X="0" Y="0" R="1" />
                <CircleByEquation Name="d" X="1" Y="0" R="2" />
                <LineByEquation Name="h" A="1" B="0" C="0" />
                <LineByEquation Name="k" A="0" B="1" C="0" />
              </Figures>
            </Drawing>
            """);
        foreach (var pair in new[] { ("A", "B"), ("f", "g"), ("first", "second"), ("c", "d"), ("h", "k") })
        {
            var rows = new CompositePropertyProvider(new IncludeByDefaultValueDiscoveryStrategy(), new object[] { Find(drawing, pair.Item1), Find(drawing, pair.Item2) })
                .GetProperties()
                .ToList();
            var definitions = rows.Where(row => row.Type == typeof(DrawingExpression) || row.Name is "FunctionText" or "Text" or "Equation").Select(row => row.Name).ToList();
            Require(definitions.Count == 0, $"{pair.Item1} and {pair.Item2} selected have rows {string.Join(", ", definitions)}.");
            Require(rows.Any(row => row.Name == nameof(IFigure.Locked)), $"{pair.Item1} and {pair.Item2} selected have no Locked row.");
        }

        var a = (PointByCoordinates)Find(drawing, "A");
        var b = (PointByCoordinates)Find(drawing, "B");
        var coordinates = new ExpressionEditor { Value = new CompositeValueProvider(new IValueProvider[] { a.XExpression, b.XExpression }) };
        Require(Commit(coordinates, "sin(") != null && a.XExpression.Text == "-2" && b.XExpression.Text == "2", "Text that doesn't compile went into the points.");
        Require(Commit(coordinates, "4") == null, "The common X was refused.");
        NearPoint(a.Coordinates, new Point(4, 1));
        NearPoint(b.Coordinates, new Point(4, -1));

        var f = (FunctionGraph)Find(drawing, "f");
        var g = (FunctionGraph)Find(drawing, "g");
        var functions = new FunctionEditor
        {
            Value = new CompositeValueProvider(new IValueProvider[]
            {
                PropertyDiscoveryStrategy.CreateValueProvider(f, nameof(FunctionGraph.FunctionText)),
                PropertyDiscoveryStrategy.CreateValueProvider(g, nameof(FunctionGraph.FunctionText))
            })
        };
        Require(Commit(functions, "sin(") != null && f.FunctionText == "x" && g.FunctionText == "-x", "A function that doesn't compile went into the graphs: " + f.FunctionText);
        Require(Commit(functions, "2 * x") == null && f.FunctionText == "2 * x" && g.FunctionText == "2 * x", "The common function was refused.");
        Near(g.Function(3), expected: 6);

        static string Commit(StringEditor editor, string text)
        {
            editor.TextBox.Text = text;
            editor.TextBox.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter });
            return string.IsNullOrEmpty(editor.ErrorText) ? null : editor.ErrorText;
        }
    }

    /// <summary>
    /// A point on a segment whose Parameter a graph reads (P.Parameter * x) offers no Free
    /// point: a free point has none. Let go all the same, the graph is drawn no more at once,
    /// as it isn't when the file is read again (the old function, holding the point that
    /// left, went on drawing it), and undo brings it back.
    /// </summary>
    static void ReleaseKeepsFunctionsWhole()
    {
        var drawing = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <FreePoint Name="A" X="-2" Y="-2" />
                <FreePoint Name="B" X="2" Y="-2" />
                <Segment Name="AB">
                  <Dependency Name="A" />
                  <Dependency Name="B" />
                </Segment>
                <PointOnFigure Name="P" X="0" Y="-2" Parameter="0.5">
                  <Dependency Name="AB" />
                </PointOnFigure>
                <FunctionGraph Name="f" Function="P.Parameter * x">
                  <Dependency Name="P" />
                </FunctionGraph>
                <FreePoint Name="C" X="1" Y="-4" />
                <FreePoint Name="D" X="1" Y="0" />
                <Segment Name="CD">
                  <Dependency Name="C" />
                  <Dependency Name="D" />
                </Segment>
              </Figures>
            </Drawing>
            """);
        var point = (PointOnFigure)Find(drawing, "P");
        Require(!PointSnapping.CanRelease(point) && !point.CanEdit(nameof(PointOnFigure.Release)), "A point whose Parameter a graph reads offers Free point.");

        // nor does Alt let it go: dragged with Alt onto where CD crosses AB, it stays on AB
        // (as a point on a figure is a FreePoint too, it was snapped into the intersection)
        using (var window = new TestWindow(drawing.Canvas))
        {
            var system = drawing.CoordinateSystem;
            drawing.Behavior = new Dragger();
            Drag(drawing, system.ToPhysical(new Point(0, -2)), system.ToPhysical(new Point(1, -2)), KeyModifiers.Alt);
            var moved = Find(drawing, "P");
            Require(moved is PointOnFigure, "Alt let go of the point: it is a " + moved.GetType().Name + " now.");
            NearPoint(((IPoint)moved).Coordinates, new Point(1, -2));
            Require(((FunctionGraph)Find(drawing, "f")).Function != null, "The graph lost its function in the Alt-drag.");
            drawing.ActionManager.Undo();
            point = (PointOnFigure)Find(drawing, "P");
        }

        PointSnapping.Release(point);
        var graph = (FunctionGraph)Find(drawing, "f");
        Require(graph.Function == null, "The graph still has the function that read the point that left.");
        drawing.ActionManager.Undo();
        Require(graph.Function != null, "Undo of letting the point go left the graph without its function.");
        Near(graph.Function(2), expected: 1);
        drawing.Figures.CheckConsistency();
    }

    static void AxisLines()
    {
        var drawing = NewDrawing();
        drawing.CoordinateGrid.Visible = true;
        using var window = new TestWindow(drawing.Canvas);
        var system = drawing.CoordinateSystem;
        var xAxis = drawing.GetAxisLine(AxisDirection.X);
        var yAxis = drawing.GetAxisLine(AxisDirection.Y);
        int Listed() => drawing.Figures.OfType<AxisLine>().Count();

        // offered on hover, not in the drawing yet
        var placement = PointPlacement.Find(drawing, new Point(2, 0.01), snapToMidpoint: false);
        Require(placement.Kind == PointPlacementKind.OnFigure && placement.Sources[0] == xAxis, "Near the x-axis: " + placement.Kind);
        Require(Listed() == 0, "Hovering put axes into the drawing.");
        string empty = drawing.SaveAsText();

        // a point on the x-axis brings both axes in
        drawing.Behavior = new FreePointCreator();
        Click(drawing, system.ToPhysical(new Point(2, 0)));
        var onX = drawing.Figures.OfType<PointOnFigure>().Single();
        Require(onX.Dependencies.Single() == xAxis, "The point is not on the x-axis.");
        Require(Listed() == 2 && drawing.Figures.Contains(yAxis), "Both axes did not come in.");
        drawing.Figures.CheckConsistency();
        string withX = drawing.SaveAsText();
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == empty && Listed() == 0, "Undo of the point left the axes.");
        drawing.ActionManager.Redo();
        Require(drawing.SaveAsText() == withX, "Redo of the point on the axis.");

        // a second point, on the y-axis: no more axes
        Click(drawing, system.ToPhysical(new Point(0, 3)));
        var onY = drawing.Figures.OfType<PointOnFigure>().Single(p => p != onX);
        Require(onY.Dependencies.Single() == yAxis && Listed() == 2, "The point on the y-axis.");

        // where a circle crosses the axis
        var center = AddPoint(drawing, x: 1, y: 1);
        var through = AddPoint(drawing, x: 1, y: 4);
        var circle = Factory.CreateCircle(drawing, new IFigure[] { center, through });
        Actions.Add(drawing, circle);
        var crossing = PointPlacement.Find(drawing, new Point(1 + System.Math.Sqrt(8), 0), snapToMidpoint: false);
        Require(crossing.Kind == PointPlacementKind.Intersection && crossing.Sources.Contains(xAxis), "The circle and the x-axis: " + crossing.Kind);

        // a parallel to the y-axis, a reflection in the x-axis
        drawing.Behavior = new ParallelLineCreator();
        Click(drawing, system.ToPhysical(new Point(0, -2)));
        Click(drawing, system.ToPhysical(new Point(5, 5)));
        Require(drawing.Figures.OfType<ParallelLine>().Single().Dependencies.Contains(yAxis), "No parallel to the y-axis.");
        drawing.Behavior = new ReflectionCreator();
        Click(drawing, system.ToPhysical(new Point(1, 4)));
        Click(drawing, system.ToPhysical(new Point(-3, 0)));
        var image = drawing.Figures.OfType<IPoint>().SingleOrDefault(p => p.Dependencies.Contains(xAxis) && p.Dependencies.Contains(through));
        Require(image != null, "No reflection in the x-axis.");
        Near(image.Coordinates.Y, expected: -4);
        Require(Listed() == 2, Listed() + " axis lines.");
        drawing.Figures.CheckConsistency();

        // a file: read back the same, one axis of each
        string saved = drawing.SaveAsText();
        var reloaded = ReadLgf(saved);
        Require(reloaded.LoadErrors == null && reloaded.SaveAsText() == saved, "The axes did not round trip.");
        Require(reloaded.Figures.OfType<AxisLine>().Count() == 2
            && reloaded.Figures.Contains(reloaded.GetAxisLine(AxisDirection.X)), "The file's axes are not the drawing's own.");

        // a paste into the same drawing builds on its axis; into another brings both
        string copied = DrawingSerializer.WriteUsingXmlWriter(writer =>
            new DrawingSerializer().WriteFiguresWithStyles(drawing, new IFigure[] { xAxis, onX }, writer));
        drawing.PasteFromText(copied);
        Require(Listed() == 2, "A paste made a second axis.");
        Require(drawing.Figures.OfType<PointOnFigure>().Count(p => p.Dependencies.Contains(xAxis)) == 2, "The copy is not on the drawing's x-axis.");
        drawing.ActionManager.Undo();
        var other = NewDrawing();
        string otherEmpty = other.SaveAsText();
        other.PasteFromText(copied);
        Require(other.Figures.OfType<AxisLine>().Count() == 2, "The paste into another drawing.");
        other.Figures.CheckConsistency();
        other.ActionManager.Undo();
        Require(other.SaveAsText() == otherEmpty, "Undo of the paste left an axis.");

        // the axes go together, when nothing is built on either
        var fresh = ReadLgf(withX);
        fresh.Figures.CheckConsistency();
        var point = fresh.Figures.OfType<PointOnFigure>().Single();
        Actions.Remove(point);
        Require(!fresh.Figures.OfType<AxisLine>().Any(), "The axes stayed after their last point.");
        fresh.ActionManager.Undo();
        Require(fresh.Figures.OfType<AxisLine>().Count() == 2 && fresh.SaveAsText() == withX, "Undo of the delete.");
        Actions.Remove(fresh.GetAxisLine(AxisDirection.X));
        Require(!fresh.Figures.OfType<AxisLine>().Any() && !fresh.Figures.Contains(point), "Deleting the x-axis left figures.");
        fresh.ActionManager.Undo();
        Require(fresh.SaveAsText() == withX, "Undo of deleting an axis.");

        // the grid hidden: no clicks on the axes, what is built on them stays
        drawing.CoordinateGrid.Visible = false;
        Require(PointPlacement.Find(drawing, new Point(-4, 0.01), snapToMidpoint: false).Kind == PointPlacementKind.Free, "A hidden axis took a click.");
        drawing.Recalculate();
        Require(onX.Exists && onY.Exists, "The points on the axes went with the grid.");
    }

    static void Hover(Drawing drawing, Point at)
    {
        var canvas = drawing.Canvas;
        using var pointer = new Pointer(Pointer.GetNextFreeId(), PointerType.Mouse, isPrimary: true);
        canvas.RaiseEvent(new PointerEventArgs(
            InputElement.PointerMovedEvent,
            canvas,
            pointer,
            canvas,
            at,
            timestamp: 0,
            new PointerPointProperties(RawInputModifiers.None, PointerUpdateKind.Other),
            KeyModifiers.None));
    }

    static void ChoiceAmongOverlaps()
    {
        var drawing = NewDrawing();
        using var window = new TestWindow(drawing.Canvas);
        var system = drawing.CoordinateSystem;
        string status = null;
        drawing.ChoiceStatus += text => status = text;

        // segment AB along the x axis and line CD up the y axis, crossing at (0, 0)
        var segment = Factory.CreateSegment(drawing, AddPoint(drawing, x: -1, y: 0), AddPoint(drawing, x: 5, y: 0));
        Actions.Add(drawing, segment);
        var line = Factory.CreateLineTwoPoints(drawing, new IFigure[] { AddPoint(drawing, x: 0, y: -1), AddPoint(drawing, x: 0, y: 3) });
        Actions.Add(drawing, line);
        var crossing = system.ToPhysical(new Point(0, 0));

        // a test window draws no frames: a tool handles one move until the next press, so
        // each hover is a fresh tool's
        FreePointCreator HoverWithPointTool()
        {
            var tool = new FreePointCreator();
            drawing.Behavior = tool;
            Hover(drawing, crossing);
            return tool;
        }

        // the first option is what a click took before there was a choice
        HoverWithPointTool();
        Require(status != null && status.StartsWith("Intersection of ") && status.Contains("(1 of 3)"), "The status at a crossing: " + status);
        Click(drawing, crossing);
        Require(drawing.Figures.OfType<IntersectionPoint>().Count() == 1, "A click at the crossing made no intersection.");
        drawing.ActionManager.Undo();

        // Tab: a point on one of the two, then on the other, then round to the crossing
        var tool = HoverWithPointTool();
        Require(tool.StepChoice(backwards: false), "Tab chose nothing.");
        Require(status.Contains("(2 of 3)"), "The status after Tab: " + status);
        Click(drawing, crossing);
        var first = drawing.Figures.OfType<PointOnFigure>().SingleOrDefault();
        Require(first != null, "The second option made no point on a figure.");
        // (the new point is all a click there means now)
        HoverWithPointTool();
        Require(status == null, "A click did not end the choice: " + status);
        drawing.ActionManager.Undo();
        tool = HoverWithPointTool();
        tool.StepChoice(backwards: false);
        tool.StepChoice(backwards: false);
        Click(drawing, crossing);
        var second = drawing.Figures.OfType<PointOnFigure>().SingleOrDefault();
        Require(second != null && second.Dependencies[0] != first.Dependencies[0], "The third option was not on the other figure.");
        drawing.ActionManager.Undo();
        tool = HoverWithPointTool();
        tool.StepChoice(backwards: true);
        Require(status.Contains("(3 of 3)"), "Shift+Tab did not go back round: " + status);

        // a tool that wants a line: the other of two along each other
        var along = Factory.CreateLineTwoPoints(drawing, new IFigure[] { AddPoint(drawing, x: -3, y: 0), AddPoint(drawing, x: 8, y: 0) });
        Actions.Add(drawing, along);
        var parallel = new ParallelLineCreator();
        drawing.Behavior = parallel;
        var onBoth = system.ToPhysical(new Point(2, 0));
        Hover(drawing, onBoth);
        Require(parallel.StepChoice(backwards: false), "No choice between a segment and a line along it.");
        Click(drawing, onBoth);
        Click(drawing, system.ToPhysical(new Point(3, 4)));
        var made = drawing.Figures.OfType<ParallelLine>().SingleOrDefault();
        Require(made != null, "No parallel was made.");
        Require(made.Dependencies[0] == segment, "Tab took " + made.Dependencies[0] + ", not the segment under the line.");

        drawing.Behavior = new Dragger();
        Hover(drawing, system.ToPhysical(new Point(-6, -6)));
        Require(status == null, "The choice's status stayed: " + status);

        // the Drag tool: the status names what is under the cursor, one thing or a choice,
        // and a click takes the one chosen with Tab
        var alone = AddPoint(drawing, x: -5, y: 5);
        var dragger = new Dragger();
        drawing.Behavior = dragger;
        Hover(drawing, system.ToPhysical(alone.Coordinates));
        Require(status == "Point " + alone.Name, "The status over a point alone: " + status);
        dragger = new Dragger();
        drawing.Behavior = dragger;
        Hover(drawing, onBoth);
        // (the line, the newer, first: a press takes it)
        Require(status != null && status.StartsWith("Line ") && status.Contains("(1 of 2)"), "The status over the segment on the line: " + status);
        Require(dragger.StepChoice(backwards: false) && status.StartsWith("Segment "), "Tab under the Drag tool: " + status);
        Click(drawing, onBoth);
        Require(segment.Selected && !along.Selected, "The click did not select the segment chosen with Tab.");
        drawing.Figures.CheckConsistency();
    }

    /// <summary>
    /// Triangle ABC with two sides as segments: one click inside the angle at A, next to the
    /// vertex, gives the Angle Bisector its three points, and gives Line at Angle, Rotate and
    /// Translate a measurement of the angle (mark and number) to take the angle from, in the
    /// same undo step
    /// </summary>
    static void AngleAtVertexShortcuts()
    {
        var drawing = NewDrawing();
        using var window = new TestWindow(drawing.Canvas);
        var system = drawing.CoordinateSystem;
        var a = AddPoint(drawing, x: 0, y: 0);
        var b = AddPoint(drawing, x: 4, y: 0);
        var c = AddPoint(drawing, x: 0, y: 3);
        var ab = Factory.CreateSegment(drawing, a, b);
        Actions.Add(drawing, ab);
        Actions.Add(drawing, Factory.CreateSegment(drawing, a, c));

        // inside the angle at A, within the mark's reach but off the point
        var insideA = system.ToPhysical(a.Coordinates) + new Point(10, -10);
        int Measurements() => drawing.Figures.OfType<AngleMeasurement>().Count();

        drawing.Behavior = new AngleBisectorCreator();
        Click(drawing, insideA);
        var bisector = drawing.Figures.OfType<AngleBisector>().SingleOrDefault();
        Require(
            bisector != null && bisector.Dependencies[0] == a && bisector.Dependencies.Contains(b) && bisector.Dependencies.Contains(c),
            "The bisector's click inside the angle.");
        Require(Measurements() == 0, "The bisector measured the angle.");
        drawing.ActionManager.Undo();

        drawing.Behavior = new LineAtAngleCreator();
        Click(drawing, insideA);
        Click(drawing, system.ToPhysical(b.Coordinates));
        var line = drawing.Figures.OfType<LineAtAngle>().SingleOrDefault();
        Require(line != null && line.AngleSource is AngleMeasurement && Measurements() == 1, "Line at Angle from the angle at A.");
        Require(drawing.Figures.OfType<AngleArc>().Count() == 1, "The angle's mark with its number.");
        drawing.ActionManager.Undo();
        Require(Measurements() == 0 && !drawing.Figures.OfType<LineAtAngle>().Any(), "Undo of the line left the measurement.");

        drawing.Behavior = new RotationCreator();
        Click(drawing, system.ToPhysical(c.Coordinates));
        Click(drawing, system.ToPhysical(b.Coordinates));
        Click(drawing, insideA);
        var rotated = drawing.Figures.OfType<RotatedPoint>().SingleOrDefault();
        Require(rotated != null && rotated.AngleSource is AngleMeasurement && Measurements() == 1, "Rotate by the angle at A.");
        drawing.ActionManager.Undo();

        drawing.Behavior = new TranslationCreator();
        Click(drawing, system.ToPhysical(c.Coordinates));
        Click(drawing, system.ToPhysical(new Point(2, 0)));
        Click(drawing, insideA);
        var translated = drawing.Figures.OfType<TranslatedPoint>().SingleOrDefault();
        Require(
            translated != null && translated.DistanceSource == ab && translated.DirectionSource is AngleMeasurement && Measurements() == 1,
            "Translate in the direction of the angle at A.");
        drawing.ActionManager.Undo();
        Require(Measurements() == 0, "Undo of the translation left the measurement.");
        drawing.Figures.CheckConsistency();
    }

    /// <summary>
    /// An angle mark on the line style is hit on its arc only; given a filled shape style it
    /// fills the angle under the arc, which a click inside then takes; undo takes it away
    /// </summary>
    static void FilledAngleMark()
    {
        var drawing = NewDrawing();
        using var window = new TestWindow(drawing.Canvas);
        var a = AddPoint(drawing, x: 0, y: 0);
        var b = AddPoint(drawing, x: 4, y: 0);
        var c = AddPoint(drawing, x: 0, y: 3);
        var arc = Factory.CreateAngleArc(drawing, new IFigure[] { a, b, c });
        Actions.Add(drawing, arc);

        // inside the sector, well off the arc
        var inside = new Point(arc.Radius * 0.4, arc.Radius * 0.4);
        Require(arc.HitTest(inside) == null, "An unfilled mark took a click inside it.");

        var style = drawing.StyleManager.GetSupportedStyles(arc).OfType<ShapeStyle>().First(s => s.IsFilled);
        Set(drawing, arc, nameof(FigureBase.Style), style);
        Require(arc.HitTest(inside) == arc, "A filled mark did not take a click inside it.");
        Require(
            drawing.Canvas.Children.OfType<Avalonia.Controls.Shapes.Path>().Any(p => p.Fill != null && p.IsVisible && !p.IsHitTestVisible),
            "No fill under the arc.");
        drawing.ActionManager.Undo();
        Require(arc.HitTest(inside) == null, "Undo of the style left the fill hit.");
        drawing.Figures.CheckConsistency();
    }

    /// <summary>
    /// One Sweep on the mark, the number, the bisector and an arc: the four choices give the
    /// two regions between the sides, the mark and the number keep each other in step (undo
    /// too), a bisector of a measurement follows the measurement, a reflection mirrors the
    /// choice, and it is saved and read back
    /// </summary>
    static void AngleSweeps()
    {
        var drawing = NewDrawing();
        var a = AddPoint(drawing, x: 0, y: 0);
        var b = AddPoint(drawing, x: 4, y: 0);
        var c = AddPoint(drawing, x: 0, y: -3);

        // counterclockwise from B to C at A is 270 degrees: the smaller angle goes clockwise
        var mark = Factory.CreateAngleArc(drawing, new IFigure[] { a, b, c });
        var number = Factory.CreateAngleMeasurement(drawing, new IFigure[] { a, b, c });
        Actions.Add(drawing, mark);
        Actions.Add(drawing, number);
        Require(number.Sweep == AngleSweep.Smaller && mark.Sweep == AngleSweep.Smaller, "A new angle is not the smaller one.");
        Near(number.Measure, 90);
        Near(mark.Measure, 90);
        Require(mark.IsClockwise, "The smaller angle from B to C goes counterclockwise.");

        var expected = new[]
        {
            (AngleSweep.Counterclockwise, 270.0),
            (AngleSweep.Clockwise, 90.0),
            (AngleSweep.Larger, 270.0),
            (AngleSweep.Smaller, 90.0)
        };
        foreach (var (sweep, measure) in expected)
        {
            Set(drawing, number, nameof(AngleMeasurement.Sweep), sweep);
            Require(mark.Sweep == sweep, sweep + ": the mark did not follow the number.");
            Near(number.Measure, measure);
            Near(mark.Measure, measure);
            Require(mark.IsClockwise == (measure == 90), sweep + ": the mark goes the wrong way round.");
        }

        drawing.ActionManager.Undo();
        Require(number.Sweep == AngleSweep.Larger && mark.Sweep == AngleSweep.Larger, "Undo of the sweep left the pair apart.");
        Set(drawing, mark, nameof(AngleArc.Sweep), AngleSweep.Clockwise);
        Require(number.Sweep == AngleSweep.Clockwise, "The number did not follow the mark.");

        // the bisector of the smaller angle points into it, below the x axis
        var bisector = Factory.CreateAngleBisector(drawing, new IFigure[] { a, b, c });
        Actions.Add(drawing, bisector);
        var direction = bisector.Coordinates.P2.Minus(bisector.Coordinates.P1);
        Require(direction.X > 0 && direction.Y < 0, "The bisector of the smaller angle points to " + direction);
        Near(bisector.Angle, 90);
        Set(drawing, bisector, nameof(AngleBisector.Sweep), AngleSweep.Larger);
        direction = bisector.Coordinates.P2.Minus(bisector.Coordinates.P1);
        Require(direction.X < 0 && direction.Y > 0, "The bisector of the larger angle points to " + direction);
        Near(bisector.Angle, 270);

        // a bisector of the number halves the angle the number says
        var onNumber = Factory.CreateAngleBisector(drawing, new IFigure[] { number });
        Actions.Add(drawing, onNumber);
        Require(!onNumber.CanEdit(nameof(AngleBisector.Sweep)), "A bisector of a measurement offers a sweep of its own.");
        Require(onNumber.Sweep == AngleSweep.Clockwise, "A bisector of a measurement does not take its sweep.");
        Near(onNumber.Angle, 90);
        Set(drawing, number, nameof(AngleMeasurement.Sweep), AngleSweep.Larger);
        Require(onNumber.Sweep == AngleSweep.Larger, "A bisector of a measurement did not follow its sweep.");
        Near(onNumber.Angle, 270);
        direction = onNumber.Coordinates.P2.Minus(onNumber.Coordinates.P1);
        Require(direction.X < 0 && direction.Y > 0, "The bisector of the measurement's larger angle points to " + direction);

        // an arc goes counterclockwise from its start unless its sweep says otherwise
        var arc = Factory.CreateArc(drawing, new IFigure[] { a, b, c });
        Actions.Add(drawing, arc);
        Require(arc.Sweep == AngleSweep.Counterclockwise, "A new arc is not counterclockwise.");
        Near(arc.Angle, 3 * System.Math.PI / 2);
        Near(arc.Length, 6 * System.Math.PI);
        Set(drawing, arc, nameof(CircleArc.Sweep), AngleSweep.Smaller);
        Require(arc.IsClockwise, "The smaller arc from B to C goes counterclockwise.");
        Near(arc.Angle, System.Math.PI / 2);
        Near(arc.Length, 2 * System.Math.PI);
        var sector = Factory.CreateCircleSector(drawing, arc.Dependencies);
        EllipseArc.Replace(arc, sector);
        Require(sector.Sweep == AngleSweep.Smaller, "The sector did not take the arc's sweep.");
        Near(sector.Area, 4 * System.Math.PI);

        // a reflection in a line mirrors a fixed choice and keeps a conditional one (the
        // images are made and let go of again: the drawing stays as it is for the file)
        var mirror = Factory.CreateSegment(drawing, a, b);
        Actions.Add(drawing, mirror);
        Set(drawing, mark, nameof(AngleArc.Sweep), AngleSweep.Clockwise);
        var markImages = Transformer.CreateReflectedFigure(drawing, mark, mirror);
        var reflectedMark = markImages.OfType<AngleArc>().Single();
        Require(reflectedMark.Sweep == AngleSweep.Counterclockwise, "The mirror image of a clockwise mark is " + reflectedMark.Sweep);
        var bisectorImages = Transformer.CreateReflectedFigure(drawing, bisector, mirror);
        var reflectedBisector = bisectorImages.OfType<AngleBisector>().Single();
        Require(reflectedBisector.Sweep == AngleSweep.Larger, "The mirror image of the larger angle's bisector is " + reflectedBisector.Sweep);
        foreach (var image in markImages.Concat(bisectorImages))
        {
            image.UnregisterFromDependencies();
        }

        Set(drawing, number, nameof(AngleMeasurement.Sweep), AngleSweep.Larger);
        var saved = drawing.SaveAsText();
        Require(saved.Contains("Sweep=\"Larger\"") && !saved.Contains("Clockwise=") && !saved.Contains("Interior="), "The sweep is not what the file says.");
        var loaded = ReadLgf(saved);
        Require(loaded.LoadErrors == null, "Reload: " + loaded.LoadErrors);
        Require(
            loaded.Figures.OfType<AngleMeasurement>().Single().Sweep == AngleSweep.Larger
                && loaded.Figures.OfType<AngleArc>().Single().Sweep == AngleSweep.Larger
                && loaded.Figures.OfType<AngleBisector>().All(f => f.Sweep == AngleSweep.Larger)
                && loaded.Figures.OfType<CircleSector>().Single().Sweep == AngleSweep.Smaller,
            "The sweeps did not come back from the file.");
        drawing.Figures.CheckConsistency();
    }

    /// <summary>
    /// One click on an arc measures its central angle the way round the arc goes, and gives
    /// a bisector of it; one click on an angle's mark gives a bisector of the angle as it is
    /// shown, its sweep included
    /// </summary>
    static void AngleFromClickedFigure()
    {
        var drawing = NewDrawing();
        using var window = new TestWindow(drawing.Canvas);
        var system = drawing.CoordinateSystem;
        var a = AddPoint(drawing, x: 0, y: 0);
        var b = AddPoint(drawing, x: 4, y: 0);
        var c = AddPoint(drawing, x: 0, y: -3);

        // counterclockwise from B round to C: 270 degrees, through the west
        var arc = Factory.CreateArc(drawing, new IFigure[] { a, b, c });
        Actions.Add(drawing, arc);
        drawing.Behavior = new AngleMeasurementCreator();
        Click(drawing, system.ToPhysical(new Point(-4, 0)));
        var number = drawing.Figures.OfType<AngleMeasurement>().SingleOrDefault();
        var mark = drawing.Figures.OfType<AngleArc>().SingleOrDefault();
        Require(number != null && mark != null && number.Dependencies.SequenceEqual(new IFigure[] { a, b, c }), "The Angle tool did not take the arc's points.");
        Require(number.Sweep == AngleSweep.Counterclockwise && mark.Sweep == AngleSweep.Counterclockwise, "The measurement did not take the arc's sweep.");
        Near(number.Measure, 270);
        drawing.ActionManager.Undo();
        Require(!drawing.Figures.OfType<AngleMeasurement>().Any() && !drawing.Figures.OfType<AngleArc>().Any(), "Undo left the measurement.");

        // the smaller arc, clockwise from B down to C: the bisector halves that one
        Set(drawing, arc, nameof(CircleArc.Sweep), AngleSweep.Smaller);
        drawing.Behavior = new AngleBisectorCreator();
        Click(drawing, system.ToPhysical(new Point(4 * System.Math.Cos(-System.Math.PI / 4), 4 * System.Math.Sin(-System.Math.PI / 4))));
        var bisector = drawing.Figures.OfType<AngleBisector>().SingleOrDefault();
        Require(bisector != null && bisector.Dependencies.SequenceEqual(new IFigure[] { a, b, c }), "The bisector did not take the arc's points.");
        Require(bisector.Sweep == AngleSweep.Smaller, "The bisector did not take the arc's sweep.");
        Near(bisector.Angle, 90);
        var direction = bisector.Coordinates.P2.Minus(bisector.Coordinates.P1);
        Require(direction.X > 0 && direction.Y < 0, "The bisector of the smaller arc points to " + direction);
        drawing.ActionManager.Undo();

        // an angle shown as the larger one: a click on its mark halves the larger one
        var angleMark = Factory.CreateAngleArc(drawing, new IFigure[] { a, b, c });
        var angleNumber = Factory.CreateAngleMeasurement(drawing, new IFigure[] { a, b, c });
        Actions.Add(drawing, angleMark);
        Actions.Add(drawing, angleNumber);
        Set(drawing, angleNumber, nameof(AngleMeasurement.Sweep), AngleSweep.Larger);
        Set(drawing, angleMark, nameof(AngleArc.Size), 60.0);
        Click(drawing, system.ToPhysical(a.Coordinates) + new Point(0, -60));
        bisector = drawing.Figures.OfType<AngleBisector>().SingleOrDefault();
        Require(bisector != null && bisector.Dependencies.SequenceEqual(new IFigure[] { a, b, c }), "The bisector did not take the mark's points.");
        Require(bisector.Sweep == AngleSweep.Larger, "The bisector did not take the angle's sweep.");
        Near(bisector.Angle, 270);
        direction = bisector.Coordinates.P2.Minus(bisector.Coordinates.P1);
        Require(direction.X < 0 && direction.Y > 0, "The bisector of the larger angle points to " + direction);
        drawing.Figures.CheckConsistency();
    }

    static void ArcConversionStyles()
    {
        var drawing = NewDrawing();
        var a = AddPoint(drawing, x: 0, y: 0);
        var b = AddPoint(drawing, x: 3, y: 0);
        var c = AddPoint(drawing, x: 0, y: 3);
        var arc = Factory.CreateArc(drawing, new IFigure[] { a, b, c });
        Actions.Add(drawing, arc);
        Require(arc.Style?.Name == StyleManager.LineStyleName, "A new arc is on " + arc.Style?.Name);

        // a new sector is the outlined shape, filled; a new circle stays on the line
        var sector = Factory.CreateCircleSector(drawing, new IFigure[] { a, b, c });
        Actions.Add(drawing, sector);
        Require(sector.Style?.Name == StyleManager.OutlinedShapeStyleName, "A new sector is on " + sector.Style?.Name);
        Require(sector.Style is ShapeStyle { IsFilled: true }, "A new sector is not filled.");
        var circle = Factory.CreateCircle(drawing, new IFigure[] { a, b });
        Actions.Add(drawing, circle);
        Require(circle.Style?.Name == StyleManager.LineStyleName, "A new circle is on " + circle.Style?.Name);
        Actions.Remove(circle);
        Actions.Remove(sector);

        // the sector made of a default arc is the outlined shape; made an arc again, the line
        var before = drawing.SaveAsText();
        arc.ConvertToSector();
        sector = drawing.Figures.OfType<CircleSector>().Single();
        Require(sector.Style?.Name == StyleManager.OutlinedShapeStyleName, "The sector of a default arc is on " + sector.Style?.Name);
        sector.ConvertToArc();
        arc = drawing.Figures.OfType<CircleArc>().Single();
        Require(arc.Style?.Name == StyleManager.LineStyleName, "The arc of a default sector is on " + arc.Style?.Name);

        // the hue follows: a thick red arc gives the red outline, a gradient blue sector the blue line
        Set(drawing, arc, nameof(FigureBase.Style), drawing.StyleManager["ThickRedLine"]);
        arc.ConvertToCircleSegment();
        var segment = drawing.Figures.OfType<CircleSegment>().Single();
        Require(segment.Style?.Name == "RedOutline", "The segment of a red arc is on " + segment.Style?.Name);

        // from one filled kind to another the style stays
        Set(drawing, segment, nameof(FigureBase.Style), drawing.StyleManager["GradientBlueOutline"]);
        segment.ConvertToSector();
        sector = drawing.Figures.OfType<CircleSector>().Single();
        Require(sector.Style?.Name == "GradientBlueOutline", "The sector of a segment is on " + sector.Style?.Name);
        sector.ConvertToArc();
        arc = drawing.Figures.OfType<CircleArc>().Single();
        Require(arc.Style?.Name == "BlueLine", "The arc of a blue sector is on " + arc.Style?.Name);

        // a style of the user's own stays where its kind fits the new figure, else the default
        var ownLine = new LineStyle() { Name = "mine", Color = Avalonia.Media.Colors.Red };
        drawing.StyleManager.Add(ownLine);
        Set(drawing, arc, nameof(FigureBase.Style), ownLine);
        arc.ConvertToSector();
        sector = drawing.Figures.OfType<CircleSector>().Single();
        Require(sector.Style == ownLine, "The sector of an arc in a style of its own is on " + sector.Style?.Name);
        var ownShape = new ShapeStyle() { Name = "mine too", Color = Avalonia.Media.Colors.Blue };
        drawing.StyleManager.Add(ownShape);
        Set(drawing, sector, nameof(FigureBase.Style), ownShape);
        sector.ConvertToArc();
        arc = drawing.Figures.OfType<CircleArc>().Single();
        Require(arc.Style?.Name == StyleManager.LineStyleName, "The arc of a sector in a shape style of its own is on " + arc.Style?.Name);

        // and undone, the arc is the first one again, on the line
        while (drawing.ActionManager.CanUndo && drawing.SaveAsText() != before)
        {
            drawing.ActionManager.Undo();
        }

        Require(drawing.SaveAsText() == before, "Undo did not bring the arc back as it was.");
        Require(drawing.Figures.OfType<CircleArc>().Single().Style?.Name == StyleManager.LineStyleName, "The arc undone is not on the line.");
        drawing.Figures.CheckConsistency();
    }

    static void ArcPanelConverts()
    {
        var drawing = NewDrawing();
        using var window = new TestWindow(drawing.Canvas);
        var system = drawing.CoordinateSystem;
        object shown = null;
        drawing.DisplayProperties += (_, args) => shown = args.Object;

        // three clicks make the arc, and the panel for it comes up
        var tool = new CircleArcCreator();
        drawing.Behavior = tool;
        Click(drawing, system.ToPhysical(new Point(0, 0)));
        Click(drawing, system.ToPhysical(new Point(2, 0)));
        Click(drawing, system.ToPhysical(new Point(0, 2)));
        var arc = drawing.Figures.OfType<CircleArc>().Single();
        Require(shown is ArcPanel panel && panel.ToString() == arc.Title, "After an arc the panel shows " + shown);
        var verbs = ((ArcPanel)shown).GetMethods().Select(method => method.Name).ToArray();
        Require(verbs.SequenceEqual(new[] { "ConvertToSector", "ConvertToSegment", "OK" }), "The panel's buttons: " + string.Join(", ", verbs));
        var before = drawing.SaveAsText();

        // the verb: a filled sector in the arc's place, one undo step, and the panel again
        // for the sector with the verbs to the other two kinds
        ((ArcPanel)shown).ConvertToSector();
        var sector = drawing.Figures.OfType<CircleSector>().Single();
        Require(!drawing.Figures.Contains(arc), "The arc is still there next to its sector.");
        Require(shown is ArcPanel sectorPanel && sectorPanel.ToString() == sector.Title, "After the verb the panel shows " + shown);
        verbs = ((ArcPanel)shown).GetMethods().Select(method => method.Name).ToArray();
        Require(verbs.SequenceEqual(new[] { "ConvertToArc", "ConvertToSegment", "OK" }), "The sector panel's buttons: " + string.Join(", ", verbs));
        Require(sector.Style?.Name == StyleManager.OutlinedShapeStyleName, "The sector is on " + sector.Style?.Name);
        Require(!drawing.IsRecordingTransaction, "The verb left a transaction open.");
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the verb did not bring the arc back.");
        drawing.ActionManager.Redo();
        sector = drawing.Figures.OfType<CircleSector>().Single();
        Require(sector.Style?.Name == StyleManager.OutlinedShapeStyleName, "Redo lost the sector's fill.");

        // a verb from the figure's own grid ends in the new figure's grid, which takes the
        // selection over; its panel has the verbs to the other two kinds, and a segment whose
        // area is measured offers no way to an arc
        sector.Selected = true;
        sector.ConvertToCircleSegment();
        var circleSegment = drawing.Figures.OfType<CircleSegment>().Single();
        Require(circleSegment.Selected, "The segment did not take the sector's selection.");
        Require(drawing.GetSelectedFigures().SequenceEqual(new IFigure[] { circleSegment }), "The selection is not the segment alone.");
        Require(shown == circleSegment, "After the grid's verb the panel shows " + shown);
        var segmentPanel = new ArcPanel(circleSegment);
        verbs = segmentPanel.GetMethods().Select(method => method.Name).ToArray();
        Require(verbs.SequenceEqual(new[] { "ConvertToArc", "ConvertToSector", "OK" }), "The segment panel's buttons: " + string.Join(", ", verbs));
        var area = Factory.CreateAreaMeasurement(drawing, new IFigure[] { circleSegment });
        Actions.Add(drawing, area);
        verbs = segmentPanel.GetMethods().Select(method => method.Name).ToArray();
        Require(verbs.SequenceEqual(new[] { "ConvertToSector", "OK" }), "A measured segment's panel: " + string.Join(", ", verbs));
        Actions.Remove(area);
        segmentPanel.ConvertToArc();
        arc = drawing.Figures.OfType<CircleArc>().Single();
        Require(shown is ArcPanel arcPanel && arcPanel.ToString() == arc.Title, "Back to an arc the panel shows " + shown);
        Require(arc.Style?.Name == StyleManager.LineStyleName, "The arc is on " + arc.Style?.Name);

        // OK only closes the panel; an elliptical arc gets the panel too, and its verb makes an elliptical segment
        ((ArcPanel)shown).OK();
        Require(shown == null, "OK left the panel up: " + shown);
        drawing.Behavior = new EllipseArcCreator();
        Click(drawing, system.ToPhysical(new Point(5, 0)));
        Click(drawing, system.ToPhysical(new Point(8, 0)));
        Click(drawing, system.ToPhysical(new Point(5, 1)));
        Click(drawing, system.ToPhysical(new Point(8, 1)));
        Click(drawing, system.ToPhysical(new Point(4, 1)));
        var ellipseArc = drawing.Figures.OfType<EllipseArc>().Single();
        Require(shown is ArcPanel ellipsePanel && ellipsePanel.ToString() == ellipseArc.Title, "After an elliptical arc the panel shows " + shown);
        ((ArcPanel)shown).ConvertToSegment();
        var segment = drawing.Figures.OfType<EllipseSegment>().Single();
        Require(shown is ArcPanel ellipseSegmentPanel && ellipseSegmentPanel.ToString() == segment.Title, "After the elliptical verb the panel shows " + shown);
        Require(segment.Style?.Name == StyleManager.OutlinedShapeStyleName, "The elliptical segment is on " + segment.Style?.Name);
        drawing.Figures.CheckConsistency();
    }

    static void PastePlainText()
    {
        var drawing = NewDrawing();
        drawing.PasteFromText("Hello, <b>world</b>");
        drawing.PasteFromText("<html><body>a page</body></html>");
        Require(!drawing.ActionManager.CanUndo, "Pasting text made an undo step.");
    }

    static void HiddenNameShowsAgain()
    {
        var drawing = NewDrawing();
        var point = AddPoint(drawing, x: 1, y: 2);
        Set(drawing, point, nameof(PointBase.ShowName), value: true);
        point.Label.Visible = false;
        Set(drawing, point, nameof(PointBase.ShowName), value: false);
        Set(drawing, point, nameof(PointBase.ShowName), value: true);
        Require(point.Label.Visible, "Show name brought the name back hidden.");
        foreach (var type in new[] { typeof(PointLabel), typeof(FigureLabel) })
        {
            var attribute = (PropertyGridVisibleAttribute)Attribute.GetCustomAttribute(type.GetProperty(nameof(IFigure.Visible)), typeof(PropertyGridVisibleAttribute));
            Require(attribute is { Visible: false }, type.Name + " has a Visible row of its own.");
        }
    }

    /// <summary>
    /// Define figure on a construction made of expressions (as the Catenary is): the tool
    /// asks for its inputs in the order they were clicked, and what it makes is built on
    /// the figures it is given, not on the ones it was defined on.
    /// </summary>
    static void DefinedToolOnExpressions()
    {
        var drawing = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <FreePoint Name="A" X="-3" Y="-1" />
                <FreePoint Name="B" X="1" Y="-2" />
                <Slider Name="s" X="-4" Y="3" Value="2" />
                <PointByCoordinates Name="Middle" Visible="false" X="(A.X + B.X) / 2" Y="(A.Y + B.Y) / 2" />
                <PointByCoordinates Name="Up" X="Middle.X" Y="Middle.Y + s" />
              </Figures>
            </Drawing>
            """);
        using var window = new TestWindow(drawing.Canvas);
        UserDefinedTool tool = null;
        Action<Behavior> created = behavior => tool = (UserDefinedTool)behavior;
        Behavior.NewBehaviorCreated += created;
        try
        {
            var definer = new MacroDefiner();
            drawing.Behavior = definer;
            Click(drawing, drawing.CoordinateSystem.ToPhysical(new Point(-3, 3)));
            Click(drawing, drawing.CoordinateSystem.ToPhysical(new Point(-3, -1)));
            Click(drawing, drawing.CoordinateSystem.ToPhysical(new Point(1, -2)));
            ((MacroDefiner.SelectInputsDialog)definer.PropertyBag).OK();
            Click(drawing, drawing.CoordinateSystem.ToPhysical(new Point(-1, 0.5)));
            ((MacroDefiner.SelectResultsDialog)definer.PropertyBag).CreateTool();
        }
        finally
        {
            Behavior.NewBehaviorCreated -= created;
        }

        Require(tool != null, "Define figure made no tool.");
        Require(drawing.Behavior == tool, "Define figure didn't hand over to the new tool.");
        Require(tool.Name == "Up", "The tool isn't named after the figure it makes: " + tool.Name);
        var inputs = tool.RootElement.Element("Inputs").Elements().Select(e => e.Attribute("Name").Value);
        Require(string.Join(" ", inputs) == "s A B", "The inputs are not in the order clicked: " + string.Join(" ", inputs));
        Require(tool.HintText == "Click a slider (s), a point (A), then a point (B).", "The tool's hint: " + tool.HintText);
        var icon = tool.RootElement.Element("Icon");
        Require(icon != null && icon.Elements("Dot").Any(d => d.ReadBool("Made", defaultValue: false)), "The tool's icon doesn't show the point it makes: " + icon);
        var first = AddPoint(drawing, x: 2, y: 0);
        var second = AddPoint(drawing, x: 4, y: 2);
        int count = drawing.Figures.Count;
        drawing.Behavior = tool;
        Click(drawing, drawing.CoordinateSystem.ToPhysical(new Point(-3, 3)));
        Click(drawing, drawing.CoordinateSystem.ToPhysical(new Point(2, 0)));
        Click(drawing, drawing.CoordinateSystem.ToPhysical(new Point(4, 2)));
        var made = drawing.Figures.Skip(count).OfType<PointByCoordinates>().ToArray();
        Require(made.Length == 2, made.Length + " points made.");
        var originals = new[] { Find(drawing, "A"), Find(drawing, "B"), Find(drawing, "Middle") };
        Require(!made.Any(p => p.Dependencies.Any(originals.Contains)), "A figure made is built on the figures the tool was defined on.");
        Require(made[0].Name.StartsWith("PointByCoordinates"), "The hidden helper took a point's letter: " + made[0].Name);
        Near(made[1].Coordinates.X, expected: 3);
        Near(made[1].Coordinates.Y, expected: 3);
        drawing.Figures.CheckConsistency();
    }

    /// <summary>
    /// A defined tool is stored when made (with the styles its figures name and a version),
    /// read back as at the next start, works in another drawing with its look, and its
    /// document goes with its button. In the in-memory settings store: nothing on disk.
    /// </summary>
    static void StoredToolsRoundTrip()
    {
        var previous = ToolStorage.Instance;
        var tools = new List<UserDefinedTool>();
        Action<Behavior> created = behavior => tools.Add((UserDefinedTool)behavior);
        Behavior.NewBehaviorCreated += created;
        try
        {
            ToolStorage.Instance = new StoredTools();
            var drawing = ReadLgf("""
                <Drawing Version="1">
                  <Styles>
                    <PointStyle Name="Big" Size="20" Fill="#FFFF0000" />
                  </Styles>
                  <Figures>
                    <FreePoint Name="A" X="-3" Y="-1" />
                    <FreePoint Name="B" X="1" Y="-2" />
                    <PointByCoordinates Name="Mid" Style="Big" X="(A.X + B.X) / 2" Y="(A.Y + B.Y) / 2" />
                    <Segment Name="AB">
                      <Dependency Name="A" />
                      <Dependency Name="B" />
                    </Segment>
                  </Figures>
                </Drawing>
                """);
            var inputs = new List<IFigure> { Find(drawing, "A"), Find(drawing, "B") };
            var results = new List<IFigure> { Find(drawing, "Mid"), Find(drawing, "AB") };
            var defined = UserDefinedTool.AddFromString(MacroSerializer.WriteMacroToString(inputs, results, "Mid"));
            var keys = SettingsStore.Current.GetDocumentKeys("Tools");
            Require(keys.Count == 1, keys.Count + " tools stored.");
            var text = SettingsStore.Current.GetDocument("Tools", keys[0]);
            var macro = XElement.Parse(text);
            Require(macro.Attribute("Version")?.Value == "1", "The stored tool has no version.");
            Require(macro.Element("Styles")?.Elements().Any(s => s.Attribute("Name")?.Value == "Big") == true, "The stored tool doesn't carry its style.");
            Require(UserDefinedTool.Read(text.Replace("Version=\"1\"", "Version=\"99\""), out string problem) == null && problem.Contains("newer"), "A tool of a newer version was read: " + problem);

            // the next start
            var reading = new StoredTools();
            ToolStorage.Instance = reading;
            reading.Load();
            var loaded = tools.Last();
            Require(loaded != defined && loaded.Name == "Mid 2", "The stored tool was not read back: " + loaded.Name);

            var other = NewDrawing();
            AddPoint(other, x: 2, y: 0).Name = "P";
            AddPoint(other, x: 4, y: 2).Name = "Q";
            using var window = new TestWindow(other.Canvas);
            other.Behavior = loaded;
            Click(other, other.CoordinateSystem.ToPhysical(new Point(2, 0)));
            Click(other, other.CoordinateSystem.ToPhysical(new Point(4, 2)));
            var made = other.Figures.OfType<PointByCoordinates>().SingleOrDefault();
            Require(made != null, "The stored tool made nothing.");
            Near(made.Coordinates.X, expected: 3);
            Near(made.Coordinates.Y, expected: 1);
            Require(made.Style?.Name == "Big" && other.StyleManager["Big"] is PointStyle big && big.Size == 20, "The made point lost its look: " + made.Style?.Name);
            var segment = other.Figures.OfType<Segment>().SingleOrDefault();
            Require(segment?.Name == "PQ", "The made segment isn't named after its points: " + segment?.Name);

            Behavior.Delete(loaded);
            Require(SettingsStore.Current.GetDocumentKeys("Tools").Count == 0, "Deleting the tool's button left its document.");
            Behavior.Delete(defined);
        }
        finally
        {
            Behavior.NewBehaviorCreated -= created;
            ToolStorage.Instance = previous;
        }
    }

    /// <summary>
    /// A show/hide box loses what is deleted of what it holds and goes when nothing is left,
    /// also when what it holds goes in one deletion (A and AB, A deleted: the removal threw
    /// half way, and undo gave the box back without A). Undo and redo give each drawing back.
    /// </summary>
    static void ShowHideBoxDeletion()
    {
        foreach (var (holds, deleted, left) in new (string Holds, string Deleted, string Left)[]
        {
            ("A AB", "A", null),
            ("A C", "A", "C"),
            ("AB BC", "B", null),
            ("AB C", "B", "C")
        })
        {
            var dependencies = string.Concat(holds.Split(' ').Select(name => $"<Dependency Name=\"{name}\" />"));
            var drawing = ReadLgf($"""
                <Drawing Version="1">
                  <Figures>
                    <FreePoint Name="A" X="0" Y="0" />
                    <FreePoint Name="B" X="3" Y="0" />
                    <FreePoint Name="C" X="3" Y="3" />
                    <Segment Name="AB">
                      <Dependency Name="A" />
                      <Dependency Name="B" />
                    </Segment>
                    <Segment Name="BC">
                      <Dependency Name="B" />
                      <Dependency Name="C" />
                    </Segment>
                    <ShowHideControl Name="box" Show="true" Text="Show" X="-4" Y="3">{dependencies}</ShowHideControl>
                  </Figures>
                </Drawing>
                """);
            string before = drawing.SaveAsText();
            drawing.Delete(new[] { Find(drawing, deleted) });
            string after = drawing.SaveAsText();
            var box = drawing.Figures.OfType<ShowHideControl>().SingleOrDefault();
            string kept = box == null ? null : string.Join(" ", box.Dependencies.Select(figure => figure.Name));
            Require(kept == left, $"A box holding {holds}, {deleted} deleted, holds {kept ?? "(gone)"}.");
            drawing.Figures.CheckConsistency();
            drawing.ActionManager.Undo();
            Require(drawing.SaveAsText() == before, $"Undo of deleting {deleted} from a box holding {holds} changed the drawing.");
            drawing.Figures.CheckConsistency();
            drawing.ActionManager.Redo();
            Require(drawing.SaveAsText() == after, $"Redo of deleting {deleted} from a box holding {holds} changed the drawing.");
            drawing.Figures.CheckConsistency();
        }
    }

    /// <summary>
    /// A click on a show/hide box in the middle of a construction (a segment's first end
    /// placed) leaves the box as it was: anything recorded then joins the construction's
    /// transaction, and shown without being recorded, the segment stayed hidden after Escape
    /// and after undo of the new segment. A click outside a construction is an undo step.
    /// </summary>
    static void ShowHideBoxDuringConstruction()
    {
        var drawing = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <FreePoint Name="A" X="0" Y="0" />
                <FreePoint Name="B" X="3" Y="0" />
                <Segment Name="AB">
                  <Dependency Name="A" />
                  <Dependency Name="B" />
                </Segment>
                <ShowHideControl Name="box" Show="true" Text="Show segment" X="-4" Y="3">
                  <Dependency Name="AB" />
                </ShowHideControl>
              </Figures>
            </Drawing>
            """);
        using var window = new TestWindow(drawing.Canvas);
        var box = (ShowHideControl)Find(drawing, "box");
        var segment = Find(drawing, "AB");
        string before = drawing.SaveAsText();
        var tool = new SegmentCreator();
        drawing.Behavior = tool;
        new FigureCreator.Dialog(tool) { X = "-2", Y = "4" }.AddPoint();
        Require(drawing.IsRecordingTransaction, "The segment's first end opened no construction.");

        // what a click on the box does to it
        box.Checkbox.IsChecked = false;
        Require(box.Checkbox.IsChecked == true && segment.Visible, "A click on the box in the middle of a construction hid AB.");
        new FigureCreator.Dialog(tool) { X = "2", Y = "4" }.AddPoint();
        Require(!drawing.IsRecordingTransaction && drawing.Figures.OfType<Segment>().Count() == 2, "The new segment was not made.");
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the new segment left " + drawing.SaveAsText());

        drawing.Behavior = new Dragger();
        box.Checkbox.IsChecked = false;
        Require(!segment.Visible, "A click on the box outside a construction didn't hide AB.");
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before && segment.Visible, "Undo of the click left " + drawing.SaveAsText());
        drawing.Figures.CheckConsistency();
    }

    /// <summary>
    /// A box's Edit figures in a fresh editor, the Selection tab open and the Misc tab
    /// remembering no tool: the ribbon, following the tool, opened Misc, which clicked its first
    /// tool and stopped the editing; started again, the tool had forgotten the box, and the
    /// next click on the paper made a second box. Through the real ribbon: a bare drawing has
    /// no tab to click anything.
    /// </summary>
    static void EditFiguresFromAnotherTab()
    {
        string startupPath = MainView.StartupPath;
        MainView.StartupPath = "/drawing";
        var view = new MainView();
        var window = new Window { Content = view, Width = 1100, Height = 850 };
        window.Show();
        try
        {
            Dispatcher.UIThread.RunJobs();
            view.OpenDrawing("box.lgf", Encoding.UTF8.GetBytes("""
                <Drawing Version="1">
                  <Figures>
                    <FreePoint Name="A" X="-2" Y="0" />
                    <FreePoint Name="B" X="0" Y="0" />
                    <FreePoint Name="C" X="2" Y="0" />
                    <ShowHideControl Name="box" Show="true" Text="Show A" X="-4" Y="3">
                      <Dependency Name="A" />
                    </ShowHideControl>
                  </Figures>
                </Drawing>
                """));
            Dispatcher.UIThread.RunJobs();
            var host = (DrawingHost)typeof(MainView)
                .GetProperty("DrawingHost", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .GetValue(view);
            var drawing = host.CurrentDrawing;
            var box = (ShowHideControl)Find(drawing, "box");
            string before = drawing.SaveAsText();

            box.EditFigures();
            Dispatcher.UIThread.RunJobs();
            Require(drawing.Behavior is ShowHideCreator && drawing.Behavior.PropertyBag is ShowHideCreator.EditPanel, "Edit figures shows " + drawing.Behavior.PropertyBag);
            Require(Find(drawing, "A").Selected && !Find(drawing, "C").Selected, "Edit figures didn't pick the box's figure A alone.");
            // as a click on C's row of the Figure List does
            ((IFigurePicker)drawing.Behavior).TogglePick(Find(drawing, "C"));
            ((ShowHideCreator.EditPanel)drawing.Behavior.PropertyBag).OK();
            Dispatcher.UIThread.RunJobs();
            Require(drawing.Figures.OfType<ShowHideControl>().Count() == 1, "Editing the box made another box.");
            Require(string.Join(" ", box.Dependencies.Select(figure => figure.Name)) == "A C", "The box holds " + string.Join(" ", box.Dependencies.Select(figure => figure.Name)));
            drawing.ActionManager.Undo();
            Require(drawing.SaveAsText() == before, "Undo of the edit left " + drawing.SaveAsText());
            drawing.Figures.CheckConsistency();
        }
        finally
        {
            window.Close();
            MainView.StartupPath = startupPath;

            // the view reports every first-chance exception from now on, the ones later cases
            // throw on purpose too
            var report = typeof(MainView)
                .GetMethod("CurrentDomain_FirstChanceException", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                .CreateDelegate<EventHandler<System.Runtime.ExceptionServices.FirstChanceExceptionEventArgs>>(view);
            AppDomain.CurrentDomain.FirstChanceException -= report;
        }
    }

    /// <summary>
    /// The end of a segment with its length and direction fixed, one of them freed by the
    /// point's own box, dragged, and fixed again: undo of the three steps gives the drawing back
    /// as it was. Fixing again took the Number kept aside when the quantity was freed and wrote
    /// the new value into it, the value the first step's undo put back: the end went to where it
    /// had been dragged, and stayed there in the file.
    /// </summary>
    static void FreedomRefixUndo()
    {
        foreach (var (free, moved) in new[] { (nameof(TranslatedPoint.FreeDistance), new Point(5, 0)), (nameof(TranslatedPoint.FreeDirection), new Point(0, 3)) })
        {
            var drawing = ReadLgf("""
                <Drawing Version="1">
                  <Figures>
                    <FreePoint Name="A" X="0" Y="0" />
                    <FreePoint Name="B" X="3" Y="0" />
                    <Segment Name="AB">
                      <Dependency Name="A" />
                      <Dependency Name="B" />
                    </Segment>
                  </Figures>
                </Drawing>
                """);
            using var window = new TestWindow(drawing.Canvas);
            var system = drawing.CoordinateSystem;
            var segment = (Segment)Find(drawing, "AB");
            segment.FixLength();
            segment.FixDirection();
            var end = (TranslatedPoint)Find(drawing, "B");
            string fixedBoth = drawing.SaveAsText();

            Set(drawing, end, free, value: true);
            drawing.Behavior = new Dragger();
            Drag(drawing, system.ToPhysical(end.Coordinates), system.ToPhysical(moved));
            NearPoint(end.Coordinates, moved);
            Set(drawing, end, free, value: false);
            string refixed = drawing.SaveAsText();

            for (int i = 0; i < 3; i++)
            {
                drawing.ActionManager.Undo();
            }

            Require(drawing.SaveAsText() == fixedBoth, free + ", a drag and fixing again, all undone, left " + drawing.SaveAsText());
            NearPoint(end.Coordinates, new Point(3, 0));
            for (int i = 0; i < 3; i++)
            {
                drawing.ActionManager.Redo();
            }

            Require(drawing.SaveAsText() == refixed, free + ": redo of the three steps changed the drawing.");
            drawing.Figures.CheckConsistency();
        }
    }

    /// <summary>A narrower range clamps the value; undo puts the range and the value back</summary>
    static void SliderRangeUndo()
    {
        var drawing = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <Slider Name="a" X="-3" Y="-2" Value="5" Maximum="10" />
              </Figures>
            </Drawing>
            """);
        var slider = (DynamicGeometry.Slider)Find(drawing, "a");
        string before = drawing.SaveAsText();
        foreach (var (property, value) in new[] { ("Maximum", 2.0), ("Minimum", 7.0), ("Step", 4.0) })
        {
            Set(drawing, slider, property, value);
            string after = drawing.SaveAsText();
            Require(slider.Value != 5, property + " didn't clamp the value.");
            drawing.ActionManager.Undo();
            Near(slider.Value, expected: 5);
            Require(drawing.SaveAsText() == before, "Undo of " + property + " changed the drawing: " + drawing.SaveAsText());
            drawing.ActionManager.Redo();
            Require(drawing.SaveAsText() == after, "Redo of " + property + " changed the drawing.");
            drawing.ActionManager.Undo();
        }
    }

    /// <summary>A vector brought to front or sent to back stays so in the file (its own WriteXml left Z out)</summary>
    static void CompositeZRoundTrips()
    {
        var drawing = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <FreePoint Name="A" X="0" Y="0" />
                <FreePoint Name="B" X="2" Y="0" />
                <Vector Name="AB" Z="-1">
                  <Dependency Name="A" />
                  <Dependency Name="B" />
                </Vector>
              </Figures>
            </Drawing>
            """);
        Require(Find(drawing, "AB").Z == -1, "A vector read without its Z.");
        var reloaded = ReadLgf(drawing.SaveAsText());
        Require(Find(reloaded, "AB").Z == -1, "A vector saved without its Z.");
    }

    /// <summary>
    /// A tool defined on a circle whose radius is a segment asks for a segment, not any line.
    /// And a figure that can't come in (made by such a tool stored before, which took a line)
    /// takes the construction back: the drawing is as it was, and the tool can be put down.
    /// </summary>
    static void DefinedToolOnSegmentLength()
    {
        var drawing = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <FreePoint Name="A" X="-4" Y="-3" />
                <FreePoint Name="B" X="-1" Y="-3" />
                <Segment Name="AB">
                  <Dependency Name="A" />
                  <Dependency Name="B" />
                </Segment>
                <FreePoint Name="C" X="-3" Y="1" />
                <CircleByRadius Name="c">
                  <Dependency Name="AB" />
                  <Dependency Name="C" />
                </CircleByRadius>
                <FreePoint Name="D" X="1" Y="-2" />
                <FreePoint Name="E" X="4" Y="1" />
                <LineTwoPoints Name="DE">
                  <Dependency Name="D" />
                  <Dependency Name="E" />
                </LineTwoPoints>
                <FreePoint Name="F" X="2" Y="3" />
              </Figures>
            </Drawing>
            """);
        var macro = MacroSerializer.WriteMacroToString(
            new List<IFigure> { Find(drawing, "C"), Find(drawing, "AB") },
            new List<IFigure> { Find(drawing, "c") },
            "Radius");
        var types = XElement.Parse(macro).Element("Inputs").Elements().Select(input => input.Attribute("Type").Value);
        Require(string.Join(" ", types) == "IPoint Segment", "The tool asks for: " + string.Join(" ", types));

        // a result's expressions read a number off an input that not every figure of its kind
        // has: u.X of a vector (given a segment, the point at (CD.X, CD.Y) was nowhere),
        // Q.Parameter of a point on a figure; a point's X and Y every point has
        var expressions = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <FreePoint Name="A" X="-4" Y="-2" />
                <FreePoint Name="B" X="-1" Y="-1" />
                <Vector Name="u">
                  <Dependency Name="A" />
                  <Dependency Name="B" />
                </Vector>
                <PointByCoordinates Name="P" X="u.X" Y="u.Y">
                  <Dependency Name="u" />
                </PointByCoordinates>
                <FreePoint Name="C" X="0" Y="3" />
                <FreePoint Name="D" X="4" Y="3" />
                <Segment Name="CD">
                  <Dependency Name="C" />
                  <Dependency Name="D" />
                </Segment>
                <PointOnFigure Name="Q" X="2" Y="3" Parameter="0.5">
                  <Dependency Name="CD" />
                </PointOnFigure>
                <PointByCoordinates Name="R" X="Q.X + 1" Y="Q.Y">
                  <Dependency Name="Q" />
                </PointByCoordinates>
                <PointByCoordinates Name="S" X="Q.Parameter" Y="0">
                  <Dependency Name="Q" />
                </PointByCoordinates>
              </Figures>
            </Drawing>
            """);
        foreach (var (input, result, expected) in new[] { ("u", "P", "Vector"), ("Q", "R", "IPoint"), ("Q", "S", "PointOnFigure") })
        {
            var written = MacroSerializer.WriteMacroToString(
                new List<IFigure> { Find(expressions, input) },
                new List<IFigure> { Find(expressions, result) },
                "Tool");
            var type = XElement.Parse(written).Element("Inputs").Elements().Single().Attribute("Type").Value;
            Require(type == expected, $"A tool on {input} making {result} asks for {type}, not {expected}.");
        }

        var old = UserDefinedTool.AddFromString(macro.Replace("Type=\"Segment\"", "Type=\"ILine\""));
        try
        {
            using var window = new TestWindow(drawing.Canvas);
            var system = drawing.CoordinateSystem;
            string before = drawing.SaveAsText();
            drawing.Behavior = old;
            Click(drawing, system.ToPhysical(new Point(2, 3)));
            try
            {
                Click(drawing, system.ToPhysical(new Point(2.5, -0.5)));
            }
            catch (InvalidOperationException)
            {
                // "d is left out: it is built on 2 figures and needs a point as the 1st", as
                // the circle came in: what matters is what it leaves behind
            }

            Require(drawing.Figures.OfType<CircleByRadius>().Count() == 1, "The circle that could not come in stayed in the drawing.");
            Require(drawing.SaveAsText() == before, "The failed construction changed the drawing.");
            drawing.Behavior = new Dragger();
            Require(drawing.SaveAsText() == before, "Putting the tool down changed the drawing.");
            drawing.Figures.CheckConsistency();
        }
        finally
        {
            Behavior.Delete(old);
        }
    }

    /// <summary>
    /// A point inside a shape has no X or Y to type, as a point on a figure hasn't: a place
    /// typed outside went to the nearest place inside, the other coordinate with it, and
    /// undo put back only the one typed. It is dragged only.
    /// </summary>
    static void PointInsideCoordinatesReadOnly()
    {
        var drawing = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <FreePoint Name="O" X="0" Y="0" />
                <FreePoint Name="P" X="2" Y="0" />
                <Circle Name="c">
                  <Dependency Name="O" />
                  <Dependency Name="P" />
                </Circle>
                <PointInRegion Name="C" X="1" Y="1" OffsetX="1" OffsetY="1">
                  <Dependency Name="c" />
                </PointInRegion>
              </Figures>
            </Drawing>
            """);
        var point = Find(drawing, "C");
        foreach (var name in new[] { "X", "Y" })
        {
            Require(!PropertyDiscoveryStrategy.CreateValueProvider(point, name).CanSetValue, "A point inside a shape has a " + name + " to type.");
        }

        Require(PropertyDiscoveryStrategy.CreateValueProvider(Find(drawing, "O"), "X").CanSetValue, "A free point has no X to type.");
    }

    /// <summary>
    /// Alt does nothing to a point dragged with the rest of a selection, which goes as one
    /// piece: with Alt the point pressed was released and moved alone, off its segment
    /// </summary>
    static void AltDragOfSelection()
    {
        var drawing = ReadLgf("""
            <Drawing Version="1">
              <Figures>
                <FreePoint Name="A" X="-2" Y="-2" />
                <FreePoint Name="B" X="2" Y="2" />
                <Segment Name="AB">
                  <Dependency Name="A" />
                  <Dependency Name="B" />
                </Segment>
                <PointOnFigure Name="C" X="0" Y="0" Parameter="0.5">
                  <Dependency Name="AB" />
                </PointOnFigure>
                <FreePoint Name="D" X="3" Y="-2" />
                <Segment Name="CD">
                  <Dependency Name="C" />
                  <Dependency Name="D" />
                </Segment>
              </Figures>
            </Drawing>
            """);
        using var window = new TestWindow(drawing.Canvas);
        var system = drawing.CoordinateSystem;
        drawing.Behavior = new Dragger();
        foreach (var figure in drawing.Figures)
        {
            figure.Selected = true;
        }

        drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());
        Drag(drawing, system.ToPhysical(new Point(0, 0)), system.ToPhysical(new Point(1, 1)), KeyModifiers.Alt);
        Require(Find(drawing, "C") is PointOnFigure, "Alt released the point dragged with the selection.");
        NearPoint(((IPoint)Find(drawing, "A")).Coordinates, new Point(-1, -1));
        NearPoint(((IPoint)Find(drawing, "D")).Coordinates, new Point(4, -1));
        NearPoint(((IPoint)Find(drawing, "C")).Coordinates, new Point(1, 1));
        drawing.Figures.CheckConsistency();
    }

    /// <summary>
    /// While a point is dragged, the status says what Shift and Alt do to it (Option on a
    /// Mac); at the drop the Drag tool's own hint is back.
    /// </summary>
    static void DragModifierHints()
    {
        var drawing = NewDrawing();
        var free = AddPoint(drawing, x: -3, y: 1);
        var first = AddPoint(drawing, x: 0, y: -3);
        var second = AddPoint(drawing, x: 0, y: 3);
        var segment = Factory.CreateSegment(drawing, first, second);
        Actions.Add(drawing, segment);
        var onSegment = Factory.CreatePointOnFigure(drawing, segment, new Point(0, 0));
        Actions.Add(drawing, onSegment);
        using var window = new TestWindow(drawing.Canvas);
        var dragger = new Dragger();
        drawing.Behavior = dragger;
        string status = null;
        drawing.Status += text => status = text;

        // the status while the button is still down, then after the release
        (string During, string After) Drag(IPoint point, Point to)
        {
            var canvas = drawing.Canvas;
            var from = drawing.CoordinateSystem.ToPhysical(point.Coordinates);
            var target = drawing.CoordinateSystem.ToPhysical(to);
            using var pointer = new Pointer(Pointer.GetNextFreeId(), PointerType.Mouse, isPrimary: true);
            canvas.RaiseEvent(new PointerPressedEventArgs(
                canvas,
                pointer,
                canvas,
                from,
                timestamp: 0,
                new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.LeftButtonPressed),
                KeyModifiers.None,
                clickCount: 1));
            canvas.RaiseEvent(new PointerEventArgs(
                InputElement.PointerMovedEvent,
                canvas,
                pointer,
                canvas,
                target,
                timestamp: 1,
                new PointerPointProperties(RawInputModifiers.LeftMouseButton, PointerUpdateKind.Other),
                KeyModifiers.None));
            var during = status;
            canvas.RaiseEvent(new PointerReleasedEventArgs(
                canvas,
                pointer,
                canvas,
                target,
                timestamp: 2,
                new PointerPointProperties(RawInputModifiers.None, PointerUpdateKind.LeftButtonReleased),
                KeyModifiers.None,
                MouseButton.Left));
            return (during, status);
        }

        // the keys as a PC's keyboard names them, then as a Mac's, whatever this runs on
        bool wasMac = KeyNames.IsMac;
        KeyNames.IsMac = false;
        try
        {
            var dragged = Drag(free, new Point(-2, 2));
            Require(dragged.During == "Hold Shift to snap to grid. Hold Alt to snap to a figure.", "Dragging a free point said: " + dragged.During);
            Require(dragged.After == dragger.HintText, "After the drop the status said: " + dragged.After);
            dragged = Drag(onSegment, new Point(0, 1));
            Require(dragged.During == "Hold Alt to detach the point.", "Dragging a point on a segment said: " + dragged.During);
            KeyNames.IsMac = true;
            dragged = Drag(free, new Point(-3, 1));
            Require(dragged.During == "Hold Shift to snap to grid. Hold Option to snap to a figure.", "On a Mac it said: " + dragged.During);
        }
        finally
        {
            KeyNames.IsMac = wasMac;
        }
    }

    static void FiguresWithoutValue()
    {
        // an empty viewport (saved from a window without a size) keeps the view
        var drawing = ReadLgf("""
            <Drawing Version="1">
              <Viewport Left="0" Top="0" Right="0" Bottom="0" />
              <Figures>
                <FreePoint Name="A" X="1" Y="2"/>
                <CircleByEquation Name="c1" X="A.X" Y="A.Y" R="-1"/>
                <CircleByEquation Name="c2" X="A.X" Y="A.Y" R="2"/>
                <LineByEquation Name="l1" A="0" B="0" C="1"/>
                <LineByEquation Name="l2" A="1" B="0" C="-3"/>
                <PointByCoordinates Name="P" X="A.X + 1" Y="sqrt(0 - 1)"/>
              </Figures>
            </Drawing>
            """);
        Require(drawing.LoadErrors == null, "The drawing did not load: " + drawing.LoadErrors);
        drawing.Recalculate();
        Require(!Find(drawing, "c1").Exists, "A circle of radius -1 exists.");
        Require(Find(drawing, "c2").Exists, "A circle of radius 2 doesn't exist.");
        Require(!Find(drawing, "l1").Exists, "The line 0 = 1 exists.");
        Require(Find(drawing, "l2").Exists, "The line x = 3 doesn't exist.");
        Require(!Find(drawing, "P").Exists, "A point at y = sqrt(-1) exists.");
    }

    static void PartialLoad()
    {
        var drawing = ReadLgf("""
            <Drawing Version="1"><Figures>
              <FreePoint Name="A" X="1" Y="2"/>
              <UnsupportedFigure Name="bad"/>
              <Segment Name="missing"><Dependency Name="A"/><Dependency Name="bad"/></Segment>
              <FreePoint Name="B" X="3" Y="4"/>
              <Segment Name="AB"><Dependency Name="A"/><Dependency Name="B"/></Segment>
            </Figures></Drawing>
            """);
        Require(drawing.LoadErrors != null, "Missing figures were not reported.");
        Require(Find(drawing, "AB") is Segment, "Valid figures after a bad one were dropped.");
    }

    // a segment on one point asked for its second as it was worked out and threw from Recalculate
    static void TooFewDependenciesLoad()
    {
        var drawing = ReadLgf("""
            <Drawing Version="1"><Figures>
              <FreePoint Name="A" X="1" Y="2"/>
              <FreePoint Name="B" X="3" Y="4"/>
              <Segment Name="short"><Dependency Name="A"/></Segment>
              <MidPoint Name="M"><Dependency Name="short"/><Dependency Name="B"/></MidPoint>
              <PerpendicularLine Name="h"><Dependency Name="A"/><Dependency Name="B"/></PerpendicularLine>
              <Segment Name="AB"><Dependency Name="A"/><Dependency Name="B"/></Segment>
            </Figures></Drawing>
            """);
        Require(drawing.LoadErrors != null && drawing.LoadErrors.Contains("short is left out"), "A segment on one point was not reported: " + drawing.LoadErrors);
        Require(drawing.LoadErrors.Contains("h is left out") && drawing.LoadErrors.Contains("needs a line"), "A perpendicular to a point was not reported: " + drawing.LoadErrors);
        Require(!drawing.Figures.Any(figure => figure.Name is "short" or "M" or "h"), "An invalid figure, or one built on it, stayed in the drawing.");
        Require(Find(drawing, "AB") is Segment && Find(drawing, "A") is FreePoint, "Valid figures were dropped with the invalid ones.");
        drawing.Figures.CheckConsistency();
    }

    static void AbstractFigureLoad()
    {
        var drawing = ReadLgf("""
            <Drawing Version="1"><Figures>
              <CircleBase Name="bad"/>
              <FreePoint Name="A" X="1" Y="2"/>
            </Figures></Drawing>
            """);
        Require(drawing.LoadErrors != null, "An abstract figure was not reported.");
        Require(Find(drawing, "A") is FreePoint, "An abstract figure prevented valid figures loading.");
    }

    // a character's size has another default (24) than a shape's (10): each is left out of
    // the file only at its own default
    static void EmojiSizeRoundTrips()
    {
        var drawing = ReadLgf("""
            <Drawing Version="1">
              <Styles>
                <PointStyle Name="SmallStar" Character="★" Size="10" Fill="#FFFFFFFF" />
                <PointStyle Name="BigStar" Character="★" Size="24" Fill="#FFFFFFFF" />
                <PointStyle Name="SmallDot" Size="10" Fill="#FFFFFFFF" />
                <PointStyle Name="BigDot" Size="24" Fill="#FFFFFFFF" />
              </Styles>
              <Figures>
                <FreePoint Name="A" Style="SmallStar" X="0" Y="0"/>
                <FreePoint Name="B" Style="BigStar" X="1" Y="0"/>
                <FreePoint Name="C" Style="SmallDot" X="2" Y="0"/>
                <FreePoint Name="D" Style="BigDot" X="3" Y="0"/>
              </Figures>
            </Drawing>
            """);
        var reloaded = ReadLgf(drawing.SaveAsText());
        foreach (var (name, size) in new[] { ("A", 10.0), ("B", 24.0), ("C", 10.0), ("D", 24.0) })
        {
            var style = Find(reloaded, name).Style as PointStyle;
            Require(style != null && style.Size == size, $"Point {name} came back at size {style?.Size}, not {size}.");
        }
    }

    static void GalleryRoundTrips()
    {
        var directory = System.IO.Path.Combine(FindRepository(), "Main", "Avalonia", "LiveGeometry", "Gallery", "Drawings");
        int count = 0;
        foreach (var path in Directory.GetFiles(directory, "*.lgf"))
        {
            var drawing = ReadLgf(File.ReadAllText(path));
            Require(drawing.LoadErrors == null, path + ": " + drawing.LoadErrors);
            var saved = drawing.SaveAsText();
            var loaded = ReadLgf(saved);
            Require(loaded.LoadErrors == null, path + ": reload: " + loaded.LoadErrors);
            Require(loaded.Figures.Count == drawing.Figures.Count, path + ": figure count changed.");
            count++;
        }

        Require(count > 0, "No gallery files found.");
        Console.WriteLine("  Gallery: " + count + " drawings loaded and reloaded.");
    }

    static string FindRepository()
    {
        for (var directory = new DirectoryInfo(Directory.GetCurrentDirectory()); directory != null; directory = directory.Parent)
        {
            if (Directory.Exists(System.IO.Path.Combine(directory.FullName, "Main", "Avalonia")))
            {
                return directory.FullName;
            }
        }

        throw new InvalidOperationException("Run this tool from the repository.");
    }

    static void Near(double actual, double expected)
    {
        Require(double.IsFinite(actual) && System.Math.Abs(actual - expected) <= 1e-10, $"Expected {expected:R}, got {actual:R}.");
    }

    static void Require(bool condition, string message)
    {
        if (!condition)
        {
            throw new InvalidOperationException(message);
        }
    }
}
