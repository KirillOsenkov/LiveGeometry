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
            ("Labels reject dependency cycles", LabelsRejectCycles),
            ("Clearing label text clears display", ClearLabelText),
            ("Functions reject dependency cycles", FunctionsRejectCycles),
            ("Repeated function references survive replacement", FunctionReplacement),
            ("Construction cancellation and undo", ConstructionUndo),
            ("Dragging and Alt snapping undo", DraggingUndo),
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
            ("Paste into another drawing keeps the look", PasteBringsStyles),
            ("Pasting plain text is no error", PastePlainText),
            ("A hidden name shows again with Show name", HiddenNameShowsAgain),
            ("Figures without a value don't exist", FiguresWithoutValue),
            ("A tool defined on expressions builds on its inputs", DefinedToolOnExpressions),
            ("Defined tools are stored, read back and deleted", StoredToolsRoundTrip),
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

        var dragged = Drag(free, new Point(-2, 2));
        Require(dragged.During == "Hold Shift to snap to grid. Hold Alt to snap to a figure.", "Dragging a free point said: " + dragged.During);
        Require(dragged.After == dragger.HintText, "After the drop the status said: " + dragged.After);
        dragged = Drag(onSegment, new Point(0, 1));
        Require(dragged.During == "Hold Alt to detach the point.", "Dragging a point on a segment said: " + dragged.During);
        bool wasMac = KeyNames.IsMac;
        KeyNames.IsMac = true;
        try
        {
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

    static void GalleryRoundTrips()
    {
        var directory = System.IO.Path.Combine(FindRepository(), @"Main\Avalonia\LiveGeometry\Gallery\Drawings");
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
            if (Directory.Exists(System.IO.Path.Combine(directory.FullName, @"Main\Avalonia")))
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
