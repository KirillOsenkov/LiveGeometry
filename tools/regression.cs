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
            ("Polygon parts survive redo", PolygonPartsSurviveRedo),
            ("Point-on-circle parameter survives undo", PointParameterSurvivesUndo),
            ("Length edits across figure kinds", LengthEditsAcrossKinds),
            ("Length panel preserves endpoint on undo", LengthPanelUndo),
            ("Point functions reject dependency cycles", PointFunctionsRejectCycles),
            ("Expression strategies agree", ExpressionStrategiesAgree),
            ("Labels reject dependency cycles", LabelsRejectCycles),
            ("Clearing label text clears display", ClearLabelText),
            ("Functions reject dependency cycles", FunctionsRejectCycles),
            ("Repeated function references survive replacement", FunctionReplacement),
            ("Construction cancellation and undo", ConstructionUndo),
            ("Dragging and Alt snapping undo", DraggingUndo),
            ("Alt dragging a point onto its neighbor joins them", AltDragJoinsNeighbor),
            ("Dragging a point says what Shift and Alt do", DragModifierHints),
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
            ("The middle of a vector gives its midpoint", VectorMidpoint),
            ("A dashed vector has a dashed shaft", DashedVector),
            ("A regular polygon's sides and vertices are selected by themselves", RegularPolygonPartSelection),
            ("A regular polygon's sides and vertices keep their styles", RegularPolygonPartStyles),
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
            ("A hidden name shows again with Show name", HiddenNameShowsAgain),
            ("Figures without a value don't exist", FiguresWithoutValue),
            ("A tool defined on expressions builds on its inputs", DefinedToolOnExpressions),
            ("Defined tools are stored, read back and deleted", StoredToolsRoundTrip),
            ("An emoji of a shape's default size keeps its size", EmojiSizeRoundTrips),
            ("LGF partial load", PartialLoad),
            ("LGF rejects abstract figures", AbstractFigureLoad),
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

    static void PolygonPartsSurviveRedo()
    {
        var drawing = NewDrawing();
        var center = AddPoint(drawing, x: 0, y: 0);
        var vertex = AddPoint(drawing, x: 3, y: 0);
        var polygon = Factory.CreateRegularPolygon(drawing, new IFigure[] { center, vertex });
        Actions.Add(drawing, polygon);
        Set(drawing, polygon, nameof(RegularPolygon.NumberOfSides), value: 6);
        var part = (IPoint)polygon.GetPart("Vertex6");
        var line = Factory.CreateSegment(drawing, center, part);
        Actions.Add(drawing, line);
        var side = polygon.GetPart("Side6");
        var onSide = Factory.CreatePointOnFigure(drawing, side, parameter: 0.4);
        Actions.Add(drawing, onSide);
        string after = drawing.SaveAsText();
        drawing.ActionManager.Undo();
        drawing.ActionManager.Undo();
        drawing.ActionManager.Undo();
        drawing.ActionManager.Redo();
        drawing.ActionManager.Redo();
        drawing.ActionManager.Redo();
        Require(ReferenceEquals(polygon.GetPart("Vertex6"), line.Dependencies[1]), "Redo left the segment on a retired vertex.");
        Require(ReferenceEquals(polygon.GetPart("Side6"), onSide.Dependencies[0]), "Redo left the point on a retired side.");
        Require(drawing.SaveAsText() == after, "Redo changed the file.");
        drawing.Figures.CheckConsistency();
    }

    static void LengthEditsAcrossKinds()
    {
        foreach (string kind in new[] { "segment", "circle", "vector", "polygon", "fixed", "onLine" })
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
                "polygon" => Factory.CreateRegularPolygon(drawing, new IFigure[] { first, second }),
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
        var value = new LengthPanel(segment).GetProperties().Single(property => property.Name == nameof(Segment.Length));
        Actions.SetProperty(drawing.ActionManager, value, value: 0.0);
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Panel length lost the endpoint.");
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
            (new TriangleCreator(), 3), (new RegularPolygonCreator(), 2), (new SquareCreator(), 2),
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

        // a caption for a name, where no expression names the figure; refused where one does
        var free = AddPoint(drawing, x: 5, y: 5);
        Require(Rename(free, "Drag me!") == null && free.Name == "Drag me!", "A caption was refused for a name.");
        Require(Rename(other, "Drag me too") != null && other.Name == "B", "A name the label can't say was taken.");

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

    static RegularPolygon AddHexagon(Drawing drawing)
    {
        var center = AddPoint(drawing, x: 0, y: 0);
        var vertex = AddPoint(drawing, x: 3, y: 0);
        var polygon = Factory.CreateRegularPolygon(drawing, new IFigure[] { center, vertex });
        Actions.Add(drawing, polygon);
        Set(drawing, polygon, nameof(RegularPolygon.NumberOfSides), value: 6);
        return polygon;
    }

    static void RegularPolygonPartSelection()
    {
        var drawing = NewDrawing();
        var polygon = AddHexagon(drawing);
        using var window = new TestWindow(drawing.Canvas);
        drawing.Behavior = new Dragger();
        var system = drawing.CoordinateSystem;
        var side = (Segment)polygon.GetPart("Side2");
        var vertex = (IPoint)polygon.GetPart("Vertex3");

        Click(drawing, system.ToPhysical(side.Coordinates.Midpoint));
        Require(drawing.GetSelectedFigures().SequenceEqual(new IFigure[] { side }), "A click on a side selected " + string.Join(", ", drawing.GetSelectedFigures()));
        Require(!polygon.Selected && !polygon.GetPart("Side1").Selected, "A click on a side selected the polygon.");

        Click(drawing, system.ToPhysical(vertex.Coordinates));
        Require(drawing.GetSelectedFigures().SequenceEqual(new IFigure[] { vertex }), "A click on a vertex selected " + string.Join(", ", drawing.GetSelectedFigures()));
        Require(!side.Selected, "The side stayed selected.");

        // the inside is the polygon, all of it
        Click(drawing, system.ToPhysical(new Point(1, 0.8)));
        Require(drawing.GetSelectedFigures().SequenceEqual(new IFigure[] { polygon }), "A click inside selected " + string.Join(", ", drawing.GetSelectedFigures()));
        Require(polygon.SelectableParts.All(part => part.Selected), "The polygon's parts don't show it selected.");

        // a side's page: its style and the polygon's side length, not a segment's rows
        var rows = ((ICustomPropertyProvider)side).GetProperties().Select(row => row.Name).ToList();
        Require(rows.SequenceEqual(new[] { nameof(RegularPolygon.Length), nameof(FigureBase.StyleDisplay) }), "A side's rows: " + string.Join(", ", rows));

        // Delete with a side selected takes the polygon, and undo brings it back
        Click(drawing, system.ToPhysical(side.Coordinates.Midpoint));
        string before = drawing.SaveAsText();
        drawing.DeleteSelection();
        Require(!drawing.Figures.Contains(polygon), "Delete left the polygon.");
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the deletion changed the drawing.");
        drawing.Figures.ClearSelection();
        Require(!drawing.GetSelectedFigures().Any(), "Clearing the selection left " + string.Join(", ", drawing.GetSelectedFigures()));
        drawing.Figures.CheckConsistency();
    }

    static void RegularPolygonPartStyles()
    {
        var drawing = NewDrawing();
        var polygon = AddHexagon(drawing);
        var red = drawing.StyleManager["RedLine"];
        var blue = drawing.StyleManager["BlueLine"];
        var redPoint = drawing.StyleManager["RedPoint"];
        Require(red != null && blue != null && redPoint != null, "Palette styles missing.");
        string before = drawing.SaveAsText();

        var side2 = polygon.GetPart("Side2");
        Set(drawing, side2, nameof(FigureBase.StyleDisplay), blue);
        string oneSide = drawing.SaveAsText();
        Require(oneSide != before, "A side's style was not saved.");

        // all the sides at once, then all the vertices; undo gives side 2 its own back
        Actions.SetProperty(drawing.ActionManager, polygon.SideStyles, red);
        Actions.SetProperty(drawing.ActionManager, polygon.VertexStyles, redPoint);
        Require(polygon.SelectableParts.OfType<Segment>().All(side => side.Style == red), "Not every side took the style.");
        Require(polygon.SelectableParts.OfType<IPoint>().All(vertex => vertex.Style == redPoint), "Not every vertex took the style.");
        string all = drawing.SaveAsText();
        drawing.ActionManager.Undo();
        drawing.ActionManager.Undo();
        Require(side2.Style == blue, "Undo did not give side 2 its style back.");
        Require(drawing.SaveAsText() == oneSide, "Undo of the sides' style changed the drawing.");
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the side's style changed the drawing.");
        drawing.ActionManager.Redo();
        drawing.ActionManager.Redo();
        drawing.ActionManager.Redo();
        Require(drawing.SaveAsText() == all, "Redo changed the drawing.");

        // a side more takes the sides' style; fewer and more again keeps an odd one's
        Set(drawing, polygon.GetPart("Side6"), nameof(FigureBase.StyleDisplay), blue);
        Set(drawing, polygon, nameof(RegularPolygon.NumberOfSides), value: 7);
        Require(polygon.GetPart("Side7").Style == red && polygon.GetPart("Vertex7").Style == redPoint, "A new side or vertex did not take the others' style.");
        Set(drawing, polygon, nameof(RegularPolygon.NumberOfSides), value: 5);
        drawing.ActionManager.Undo();
        Require(polygon.GetPart("Side6").Style == blue, "Undo of fewer sides lost a side's style.");

        // what is saved comes back
        string saved = drawing.SaveAsText();
        var reloaded = ReadLgf(saved);
        var copy = reloaded.Figures.OfType<RegularPolygon>().Single();
        Require(copy.GetPart("Side6").Style?.Name == "BlueLine", "Side 6 came back as " + copy.GetPart("Side6").Style?.Name);
        Require(copy.GetPart("Side1").Style?.Name == "RedLine", "Side 1 came back as " + copy.GetPart("Side1").Style?.Name);
        Require(copy.GetPart("Vertex4").Style?.Name == "RedPoint", "Vertex 4 came back as " + copy.GetPart("Vertex4").Style?.Name);
        Require(reloaded.SaveAsText() == saved, "The polygon's styles did not round trip.");
        drawing.Figures.CheckConsistency();
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
        string after = drawing.SaveAsText();
        drawing.ActionManager.Undo();
        Require(drawing.SaveAsText() == before, "Undo of the square changed the drawing.");
        drawing.ActionManager.Redo();
        Require(drawing.SaveAsText() == after, "Redo of the square changed the drawing.");
        drawing.Figures.CheckConsistency();
        Require(ReadLgf(after).LoadErrors == null, "The drawing with the square did not load.");
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
