#:property Nullable=disable
#:property PublishAot=false

// tangram - writes the "Tangram" gallery drawing: the seven pieces, each a polygon on a free
// pivot point and a direction point kept on a circle of radius 1 around the pivot, the other
// corners worked out from the two, so that dragging the piece moves it and dragging the
// small knob on its edge turns it. Scattered around a dashed square they came out of; a box
// shows the solution.
//
//   dotnet tools/tangram.cs -- <out.lgf>

using System.Globalization;
using System.Text;

if (args.Length < 1)
{
    Console.WriteLine("usage: tangram <out.lgf>");
    return 1;
}

var invariant = CultureInfo.InvariantCulture;
var text = new StringBuilder();

void Write(string line)
{
    text.Append(line).Append("\r\n");
}

string Format(double value)
{
    return value.ToString("0.####", invariant);
}

// the pieces in the 4 x 4 square, first corner first (the pivot), then where each starts:
// its pivot and how far it is turned, in degrees, and its default style - the palette's
// outlined gradients, one hue each (tangrams have no colors of their own: the old sets are
// black card or plain wood)
var pieces = new[]
{
    new Piece("Big1", new[] { (0.0, 0.0), (4.0, 0.0), (2.0, 2.0) }, (-6.5, 1.5), 45, "GradientRedOutline"),
    new Piece("Big2", new[] { (0.0, 0.0), (2.0, 2.0), (0.0, 4.0) }, (5.0, -3.6), -45, "GradientBlueOutline"),
    new Piece("Medium", new[] { (2.0, 4.0), (4.0, 4.0), (4.0, 2.0) }, (5.6, 2.2), 90, "GradientGreenOutline"),
    new Piece("Square", new[] { (1.0, 3.0), (2.0, 2.0), (3.0, 3.0), (2.0, 4.0) }, (-5.4, -3.2), 0, "GradientBrownOutline"),
    new Piece("Small1", new[] { (0.0, 4.0), (1.0, 3.0), (2.0, 4.0) }, (-0.6, 4.4), -45, "GradientPurpleOutline"),
    new Piece("Small2", new[] { (2.0, 2.0), (3.0, 1.0), (3.0, 3.0) }, (-1.6, -5.0), 180, "GradientCyanOutline"),
    new Piece("Parallelogram", new[] { (3.0, 1.0), (4.0, 0.0), (4.0, 2.0), (3.0, 3.0) }, (2.6, 4.9), 225, "GradientOrangeOutline"),
};

Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
Write("<Drawing Version=\"1\" Creator=\"LiveGeometry.App\">");
Write("  <Viewport Left=\"-9\" Top=\"7\" Right=\"9\" Bottom=\"-7\" />");
Write("  <Styles>");
Write("    <ShapeStyle Name=\"Target\" IsFilled=\"false\" Color=\"#FF8A94A6\" StrokeWidth=\"1.5\" Dash=\"Dash\" />");
Write("    <ShapeStyle Name=\"Solution\" IsFilled=\"true\" Fill=\"#30808080\" Color=\"#A0808080\" StrokeWidth=\"1\" Dash=\"Dash\" />");
Write("    <PointStyle Name=\"Pivot\" Size=\"9\" Fill=\"#FFFFFFFF\" Color=\"#80000000\" StrokeWidth=\"1.5\" />");
Write("    <PointStyle Name=\"Knob\" Size=\"9\" Fill=\"#FFFFE066\" Color=\"#FF806000\" StrokeWidth=\"1.5\" />");
Write("  </Styles>");
Write("  <Figures>");
Write("    <Number Name=\"One\" Value=\"1\" />");

// the square the pieces came out of, in the middle, and the solution: the pieces where they
// were in it
foreach (var (name, x, y) in new[] { ("TargetA", -2.0, -2.0), ("TargetB", 2.0, -2.0), ("TargetC", 2.0, 2.0), ("TargetD", -2.0, 2.0) })
{
    Write($"    <PointByCoordinates Name=\"{name}\" Visible=\"false\" X=\"{Format(x)}\" Y=\"{Format(y)}\" />");
}

Write("    <Polygon Name=\"Target\" Style=\"Target\">");
Write("      <Dependency Name=\"TargetA\" />");
Write("      <Dependency Name=\"TargetB\" />");
Write("      <Dependency Name=\"TargetC\" />");
Write("      <Dependency Name=\"TargetD\" />");
Write("    </Polygon>");
var solutionNames = new List<string>();
foreach (var piece in pieces)
{
    var names = new List<string>();
    for (int i = 0; i < piece.Corners.Length; i++)
    {
        string name = $"Solved{piece.Name}{i}";
        Write($"    <PointByCoordinates Name=\"{name}\" Visible=\"false\" X=\"{Format(piece.Corners[i].X - 2)}\" Y=\"{Format(piece.Corners[i].Y - 2)}\" />");
        names.Add(name);
    }

    string polygon = $"Solved{piece.Name}";
    Write($"    <Polygon Name=\"{polygon}\" Visible=\"false\" Style=\"Solution\">");
    foreach (var name in names)
    {
        Write($"      <Dependency Name=\"{name}\" />");
    }

    Write("    </Polygon>");
    solutionNames.Add(polygon);
}

// the pieces: pivot P, direction D on the unit circle around P, the piece's angle A the
// knob's rounded to the nearest 45 degrees (every corner of the solved square is then on
// the grid, so that Shift-dragging a piece there lands it exactly), corner = P + u (cos A,
// sin A) + v (-sin A, cos A)
foreach (var piece in pieces)
{
    var pivot = piece.Corners[0];
    var first = piece.Corners[1];
    double length = Math.Sqrt((first.X - pivot.X) * (first.X - pivot.X) + (first.Y - pivot.Y) * (first.Y - pivot.Y));
    var e = ((first.X - pivot.X) / length, (first.Y - pivot.Y) / length);
    double turn = piece.Turn * Math.PI / 180;
    string p = piece.Name + "P";
    string d = piece.Name + "D";
    Write($"    <FreePoint Name=\"{p}\" Style=\"Pivot\" X=\"{Format(piece.Start.X)}\" Y=\"{Format(piece.Start.Y)}\" />");
    Write($"    <CircleByRadius Name=\"{piece.Name}Dial\" Visible=\"false\">");
    Write("      <Dependency Name=\"One\" />");
    Write($"      <Dependency Name=\"{p}\" />");
    Write("    </CircleByRadius>");
    Write($"    <PointOnFigure Name=\"{d}\" Style=\"Knob\" X=\"{Format(piece.Start.X + Math.Cos(turn))}\" Y=\"{Format(piece.Start.Y + Math.Sin(turn))}\" Parameter=\"{Format(turn)}\">");
    Write($"      <Dependency Name=\"{piece.Name}Dial\" />");
    Write("    </PointOnFigure>");
    string a = piece.Name + "A";
    Write($"    <PointByCoordinates Name=\"{a}\" Visible=\"false\" X=\"round(xangle({p}, {d}) / (pi / 4)) * (pi / 4)\" Y=\"0\">");
    Write($"      <Dependency Name=\"{p}\" />");
    Write($"      <Dependency Name=\"{d}\" />");
    Write("    </PointByCoordinates>");
    var names = new List<string> { p };
    for (int i = 1; i < piece.Corners.Length; i++)
    {
        var corner = piece.Corners[i];
        double dx = corner.X - pivot.X;
        double dy = corner.Y - pivot.Y;
        double u = dx * e.Item1 + dy * e.Item2;
        double v = -dx * e.Item2 + dy * e.Item1;
        string name = $"{piece.Name}{i}";
        Write($"    <PointByCoordinates Name=\"{name}\" Visible=\"false\" X=\"{p}.X + {Format(u)} * cos({a}.X) - {Format(v)} * sin({a}.X)\" Y=\"{p}.Y + {Format(u)} * sin({a}.X) + {Format(v)} * cos({a}.X)\">");
        Write($"      <Dependency Name=\"{p}\" />");
        Write($"      <Dependency Name=\"{a}\" />");
        Write("    </PointByCoordinates>");
        names.Add(name);
    }

    Write($"    <Polygon Name=\"{piece.Name}\" Style=\"{piece.Style}\">");
    foreach (var name in names)
    {
        Write($"      <Dependency Name=\"{name}\" />");
    }

    Write("    </Polygon>");
}

Write("    <ShowHideControl Name=\"SolutionBox\" Style=\"GalleryText\" Show=\"false\" Text=\"Solution\" Pin=\"TopLeft\" OffsetX=\"16\" OffsetY=\"16\">");
foreach (var name in solutionNames)
{
    Write($"      <Dependency Name=\"{name}\" />");
}

Write("    </ShowHideControl>");
Write("    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"Tangram\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"Seven pieces: five triangles, a square and a parallelogram. They were cut from the dashed square in the middle. Can you put them back? It is harder than it looks.\\n\\nDrag a piece to move it; hold Shift while you drag and it snaps to the grid. Drag the small yellow knob on its edge to turn it, in steps of 45 degrees. If you give up, tick the box.\\n\\nThe tangram came from China about 200 years ago and spread around the world as a craze. With the same seven pieces people have made thousands of figures: cats, swans, runners, houses, letters. Make one of your own.\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("  </Figures>");
text.Append("</Drawing>");
File.WriteAllText(args[0], text.ToString(), new UTF8Encoding(false));
Console.WriteLine("wrote " + args[0]);
return 0;

record Piece(string Name, (double X, double Y)[] Corners, (double X, double Y) Start, double Turn, string Style);
