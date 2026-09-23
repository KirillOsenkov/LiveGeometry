#:property Nullable=disable
#:property PublishAot=false

// fibonacci - writes the "Fibonacci Spiral" gallery drawing: the squares 13, 8, 5, 3, 2, 1, 1
// tiled into a 21 x 13 rectangle, a quarter circle in each, a number in each. The two free
// points A (bottom) and B (top) are the left edge of the big square; every corner is A plus
// a whole number of steps along AB/13 and its perpendicular, so dragging A or B turns and
// resizes the whole tiling and nothing can come apart.
//
//   dotnet tools/fibonacci.cs -- <out.lgf>

using System.Globalization;
using System.Text;

if (args.Length < 1)
{
    Console.WriteLine("usage: fibonacci <out.lgf>");
    return 1;
}

var invariant = CultureInfo.InvariantCulture;

// (left, bottom, side), the arc's center corner, its begin corner and its end corner, in
// grid units of the 21 x 13 rectangle. An arc is drawn counterclockwise from begin to end.
// The spiral itself winds clockwise from the big square inward, so each arc is listed
// backwards: the end of one is the begin of the next, and the tangents agree there (the
// center of each arc is the corner on the inside of the turn, the one facing the next
// square)
var squares = new[]
{
    new Square(13, 0, 0, (13, 0), (13, 13), (0, 0), "#FFC06DFF", "#FF860095"),
    new Square(8, 13, 5, (13, 5), (21, 5), (13, 13), "#FF0099FF", "#FF2A0EB8"),
    new Square(5, 16, 0, (16, 5), (16, 0), (21, 5), "#FF00FFE5", "#FF009196"),
    new Square(3, 13, 0, (16, 3), (13, 3), (16, 0), "#FF11FF00", "#FF047700"),
    new Square(2, 13, 3, (15, 3), (15, 5), (13, 3), "#FFFFF700", "#FF777303"),
    new Square(1, 15, 4, (15, 4), (16, 4), (15, 5), "#FFFF8800", "#FF885109"),
    new Square(1, 15, 3, (15, 4), (15, 3), (16, 4), "#FFFF6D6D", "#FF8E0606"),
};

var styles = new StringBuilder();
var pointFigures = new StringBuilder();
var figures = new StringBuilder();
var points = new HashSet<string>();

// a corner of the grid as a point: A + x steps to the right + y steps up, where a step is
// AB/13 and "right" is "up" turned clockwise
string Corner((int X, int Y) corner)
{
    string name = "P" + corner.X + "_" + corner.Y;
    if (points.Add(name))
    {
        pointFigures.AppendLine(string.Format(invariant,
            "    <PointByCoordinates Name=\"{0}\" Visible=\"false\" Style=\"DependentPointStyle\" X=\"A.X + {1} * (B.Y - A.Y) / 13 + {2} * (B.X - A.X) / 13\" Y=\"A.Y - {1} * (B.X - A.X) / 13 + {2} * (B.Y - A.Y) / 13\">",
            name, corner.X, corner.Y));
        pointFigures.AppendLine("      <Dependency Name=\"A\" />");
        pointFigures.AppendLine("      <Dependency Name=\"B\" />");
        pointFigures.AppendLine("    </PointByCoordinates>");
    }

    return name;
}

int index = 0;
foreach (var square in squares)
{
    index++;
    string fill = "Square" + index;
    string arc = "Arc" + index;
    styles.AppendLine("    <ShapeStyle Fill=\"" + square.Fill + "\" IsFilled=\"true\" Color=\"#FF000000\" StrokeWidth=\"1\" Dash=\"Solid\" Name=\"" + fill + "\" />");
    styles.AppendLine("    <LineStyle Color=\"" + square.Arc + "\" StrokeWidth=\"5\" Dash=\"Solid\" Name=\"" + arc + "\" />");

    var corners = new[]
    {
        (square.Left, square.Bottom),
        (square.Left + square.Side, square.Bottom),
        (square.Left + square.Side, square.Bottom + square.Side),
        (square.Left, square.Bottom + square.Side)
    };
    figures.AppendLine("    <Polygon Name=\"" + fill + "\" Style=\"" + fill + "\">");
    foreach (var corner in corners.Concat(new[] { corners[0] }))
    {
        figures.AppendLine("      <Dependency Name=\"" + Corner(corner) + "\" />");
    }

    figures.AppendLine("    </Polygon>");
}

// the arcs over the squares, so that the polygons' outlines don't cross them
index = 0;
foreach (var square in squares)
{
    index++;
    figures.AppendLine("    <CircleArc Name=\"Arc" + index + "\" Style=\"Arc" + index + "\">");
    figures.AppendLine("      <Dependency Name=\"" + Corner(square.Center) + "\" />");
    figures.AppendLine("      <Dependency Name=\"" + Corner(square.Begin) + "\" />");
    figures.AppendLine("      <Dependency Name=\"" + Corner(square.End) + "\" />");
    figures.AppendLine("    </CircleArc>");
}

// the number in each square: a hidden point in the middle, named with the number, with its
// name shown; the second 1 is "1 " so that the names stay distinct
index = 0;
foreach (var square in squares)
{
    index++;
    string name = square.Side.ToString(invariant) + (index == squares.Length ? " " : "");
    double x = square.Left + square.Side / 2.0;
    double y = square.Bottom + square.Side / 2.0;
    figures.AppendLine(string.Format(invariant,
        "    <PointByCoordinates Name=\"{0}\" Visible=\"false\" Style=\"DependentPointStyle\" X=\"A.X + {1} * (B.Y - A.Y) / 13 + {2} * (B.X - A.X) / 13\" Y=\"A.Y - {1} * (B.X - A.X) / 13 + {2} * (B.Y - A.Y) / 13\">",
        name, x, y));
    figures.AppendLine("      <Dependency Name=\"A\" />");
    figures.AppendLine("      <Dependency Name=\"B\" />");
    figures.AppendLine("    </PointByCoordinates>");
    double width = square.Side.ToString().Length * 8;
    figures.AppendLine(string.Format(invariant,
        "    <PointLabel Name=\"Number{0}\" IsHitTestVisible=\"false\" Style=\"Number\" OffsetX=\"{1}\" OffsetY=\"-10\" ShowName=\"true\" ShowCoordinates=\"false\">",
        index, -width / 2));
    figures.AppendLine("      <Dependency Name=\"" + name + "\" />");
    figures.AppendLine("    </PointLabel>");
}

var text = new StringBuilder();
text.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
text.AppendLine("<Drawing Version=\"1\">");
text.AppendLine("  <Viewport Left=\"-6\" Top=\"4\" Right=\"6.5\" Bottom=\"-4\" Grid=\"false\" />");
text.AppendLine("  <Styles>");
text.AppendLine("    <PointStyle Size=\"10\" Fill=\"#FFFFFF64\" IsFilled=\"true\" Color=\"#64000000\" StrokeWidth=\"1\" Dash=\"Solid\" Name=\"FreePoint\" />");
text.AppendLine("    <PointStyle Size=\"10\" Fill=\"#FF7CE38B\" IsFilled=\"true\" Color=\"#64000000\" StrokeWidth=\"1\" Dash=\"Solid\" Name=\"PointOnFigure\" />");
text.AppendLine("    <PointStyle Size=\"8\" Fill=\"#FF6FD3F7\" IsFilled=\"true\" Color=\"#64000000\" StrokeWidth=\"1\" Dash=\"Solid\" Name=\"IntersectionPoint\" />");
text.AppendLine("    <PointStyle Size=\"8\" Fill=\"#FFFFB45A\" IsFilled=\"true\" Color=\"#64000000\" StrokeWidth=\"1\" Dash=\"Solid\" Name=\"Midpoint\" />");
text.AppendLine("    <PointStyle Size=\"8\" Fill=\"#FFD0D0D0\" IsFilled=\"true\" Color=\"#64000000\" StrokeWidth=\"1\" Dash=\"Solid\" Name=\"DependentPointStyle\" />");
text.AppendLine("    <TextStyle FontSize=\"15\" Color=\"#FF2B3038\" FontFamily=\"Segoe UI\" Bold=\"false\" Italic=\"false\" Underline=\"false\" Name=\"Number\" />");
text.Append(styles);
text.AppendLine("    <TextStyle FontSize=\"30\" Color=\"#FF1F4E8C\" FontFamily=\"Segoe UI\" Bold=\"true\" Italic=\"false\" Underline=\"false\" Name=\"GalleryTitle\" />");
text.AppendLine("    <TextStyle FontSize=\"16\" Color=\"#FF2B3038\" FontFamily=\"Segoe UI\" Bold=\"false\" Italic=\"false\" Underline=\"false\" Name=\"GalleryText\" />");
text.AppendLine("    <LineStyle Color=\"#FFE0362B\" StrokeWidth=\"2.5\" Name=\"GalleryLocus\" />");
text.AppendLine("  </Styles>");
text.AppendLine("  <Figures>");
text.AppendLine("    <FreePoint Name=\"A\" Style=\"FreePoint\" X=\"-5\" Y=\"-3.25\" />");
text.AppendLine("    <FreePoint Name=\"B\" Style=\"FreePoint\" X=\"-5\" Y=\"3.25\" />");
text.Append(pointFigures);
text.Append(figures);
text.AppendLine("    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"Fibonacci Spiral\" DecimalsToShow=\"2\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"57.5\" WrapWidth=\"400\" Backdrop=\"true\" />");
text.AppendLine("    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"1, 1, 2, 3, 5, 8, 13: each Fibonacci number is the sum of the two before it. Squares with these sides fit together into a rectangle - every new square is exactly as wide as the two before it side by side.\\n\\nA quarter circle in each square, corner to corner, makes a spiral. Every quarter turn it grows by about the same factor, close to 1.618, the golden ratio. A snail shell grows by a steady factor too, which is why it looks alike.\\n\\nCount the spirals of seeds on a sunflower head or of scales on a pinecone: you get two neighbors from this list.\\n\\nDrag the yellow points to turn and resize it.\" DecimalsToShow=\"2\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"105.5\" WrapWidth=\"400\" Backdrop=\"true\" />");
text.AppendLine("  </Figures>");
text.Append("</Drawing>");
File.WriteAllText(args[0], text.ToString().Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(false));
Console.WriteLine("wrote " + args[0]);
return 0;

record Square(int Side, int Left, int Bottom, (int X, int Y) Center, (int X, int Y) Begin, (int X, int Y) End, string Fill, string Arc);
