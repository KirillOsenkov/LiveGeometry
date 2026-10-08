#:property Nullable=disable
#:property PublishAot=false

// squares17 - writes the "Seventeen Squares" gallery drawing: John Bidwell's packing of 17
// unit squares into a square of side 4.6755..., the smallest anyone has found. The two free
// points A and B are the bottom side of the box; every corner of every square is A plus x
// steps along AB/s and y steps along its perpendicular (a hidden point U is A plus one
// step), so dragging A or B turns and resizes the whole packing and nothing can come apart.
// A show/hide box draws the 4 by 4 box that 16 squares fill exactly.
//
// The centers and angles are those of the witness file of the Squares project
// (github.com/jlevy/squares, packing/witnesses/known-best/n-017.yaml), cut to 20 digits.
//
//   dotnet tools/squares17.cs -- <out.lgf>

using System.Globalization;
using System.Text;

if (args.Length < 1)
{
    Console.WriteLine("usage: squares17 <out.lgf>");
    return 1;
}

var invariant = CultureInfo.InvariantCulture;
const double side = 4.67553009360455095163;

// center x, center y, angle in degrees
var squares = new (double X, double Y, double Angle)[]
{
    (0.5, 0.5, 0),
    (1.5, 0.5, 0),
    (0.5, 1.5, 0),
    (0.5, 4.17553009360455095163, 0),
    (4.17553009360455095163, 0.5, 0),
    (3.17553009360455095163, 0.5, 0),
    (4.17553009360455095163, 1.5, 0),
    (4.17553009360455095163, 4.17553009360455095163, 0),
    (2.34732482651028509718, 4.17553009360455095163, 0),
    (4.17553009360455095163, 2.61346013251256404552, 0),
    (0.70420215565751755964, 2.70420215565751755964, 39.8049589797677950558),
    (1.43604349399822481338, 3.38804343326336292085, 39.8049589797677950558),
    (1.55673716317087917381, 2.11293587184035188566, 39.8049589797677950558),
    (2.28857850151158642755, 2.79677714944619724687, 39.8049589797677950558),
    (2.29528644618152335793, 1.42668354766673931183, 39.8049589797677950558),
    (3.02712778452223061166, 2.11052482527258467305, 39.8049589797677950558),
    (3.24507371265696607273, 3.37249386149385352755, 53.3762136165534106798),
};

string StyleOf(double angle) => angle == 0 ? "GradientBlueOutline"
    : angle < 45 ? "GradientOrangeOutline"
    : "GradientRedOutline";

// a number as the expression scanner reads it: no exponent, a negative one in parentheses
string Number(double value)
{
    string text = System.Math.Round(value, 15).ToString("0.###############", invariant);
    if (text == "-0")
    {
        text = "0";
    }

    return text.StartsWith('-') ? "(" + text + ")" : text;
}

var points = new StringBuilder();
var figures = new StringBuilder();

// a point at (x, y) in the units of the packing
void AddPoint(string name, double x, double y)
{
    points.AppendLine(string.Format(invariant,
        "    <PointByCoordinates Name=\"{0}\" Visible=\"false\" X=\"A.X + {1} * (U.X - A.X) - {2} * (U.Y - A.Y)\" Y=\"A.Y + {1} * (U.Y - A.Y) + {2} * (U.X - A.X)\">",
        name, Number(x), Number(y)));
    points.AppendLine("      <Dependency Name=\"A\" />");
    points.AppendLine("      <Dependency Name=\"U\" />");
    points.AppendLine("    </PointByCoordinates>");
}

void AddSegment(string name, string from, string to, string style, bool visible)
{
    figures.AppendLine("    <Segment Name=\"" + name + "\"" + (visible ? "" : " Visible=\"false\"") + " Style=\"" + style + "\">");
    figures.AppendLine("      <Dependency Name=\"" + from + "\" />");
    figures.AppendLine("      <Dependency Name=\"" + to + "\" />");
    figures.AppendLine("    </Segment>");
}

// the box: a polygon with a transparent fill under the squares, so that a drag in a gap moves
// the box rather than the view
figures.AppendLine("    <Polygon Name=\"ABCD\" Style=\"Box\">");
foreach (var corner in new[] { "A", "B", "C", "D" })
{
    figures.AppendLine("      <Dependency Name=\"" + corner + "\" />");
}

figures.AppendLine("    </Polygon>");

int index = 0;
foreach (var square in squares)
{
    index++;
    double radians = square.Angle * System.Math.PI / 180;
    double cos = System.Math.Cos(radians);
    double sin = System.Math.Sin(radians);
    var corners = new List<string>();
    foreach (var (dx, dy) in new[] { (-0.5, -0.5), (0.5, -0.5), (0.5, 0.5), (-0.5, 0.5) })
    {
        string name = "S" + index + "_" + (corners.Count + 1);
        AddPoint(name, square.X + dx * cos - dy * sin, square.Y + dx * sin + dy * cos);
        corners.Add(name);
    }

    figures.AppendLine("    <Polygon Name=\"Square" + index + "\" Style=\"" + StyleOf(square.Angle) + "\">");
    foreach (var corner in corners)
    {
        figures.AppendLine("      <Dependency Name=\"" + corner + "\" />");
    }

    figures.AppendLine("    </Polygon>");
}

AddSegment("AB", "A", "B", "ThickLine", visible: true);
AddSegment("BC", "B", "C", "ThickLine", visible: true);
AddSegment("CD", "C", "D", "ThickLine", visible: true);
AddSegment("AD", "D", "A", "ThickLine", visible: true);

// the 4 by 4 box from A, of which only the two sides inside the big box are drawn
AddPoint("F1", 4, 0);
AddPoint("F2", 4, 4);
AddPoint("F3", 0, 4);
AddSegment("Four1", "F1", "F2", "DashedRedLine", visible: false);
AddSegment("Four2", "F2", "F3", "DashedRedLine", visible: false);

const double half = 2.4;
var text = new StringBuilder();
text.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
text.AppendLine("<Drawing Version=\"1\">");
text.AppendLine(string.Format(invariant, "  <Viewport Left=\"{0}\" Top=\"{1}\" Right=\"{1}\" Bottom=\"{0}\" />", -half - 0.5, half + 0.5));
text.AppendLine("  <Styles>");
text.AppendLine("    <ShapeStyle Name=\"Box\" Fill=\"#00FFFFFF\" Color=\"#00FFFFFF\" />");
text.AppendLine("  </Styles>");
text.AppendLine("  <Figures>");
text.AppendLine(string.Format(invariant, "    <FreePoint Name=\"A\" X=\"{0}\" Y=\"{0}\" />", -side / 2));
text.AppendLine(string.Format(invariant, "    <FreePoint Name=\"B\" X=\"{0}\" Y=\"{1}\" />", side / 2, -side / 2));
text.AppendLine("    <PointByCoordinates Name=\"U\" Visible=\"false\" X=\"A.X + (B.X - A.X) / " + Number(side) + "\" Y=\"A.Y + (B.Y - A.Y) / " + Number(side) + "\">");
text.AppendLine("      <Dependency Name=\"A\" />");
text.AppendLine("      <Dependency Name=\"B\" />");
text.AppendLine("    </PointByCoordinates>");
text.AppendLine("    <PointByCoordinates Name=\"C\" Visible=\"false\" X=\"B.X - (B.Y - A.Y)\" Y=\"B.Y + (B.X - A.X)\">");
text.AppendLine("      <Dependency Name=\"A\" />");
text.AppendLine("      <Dependency Name=\"B\" />");
text.AppendLine("    </PointByCoordinates>");
text.AppendLine("    <PointByCoordinates Name=\"D\" Visible=\"false\" X=\"A.X - (B.Y - A.Y)\" Y=\"A.Y + (B.X - A.X)\">");
text.AppendLine("      <Dependency Name=\"A\" />");
text.AppendLine("      <Dependency Name=\"B\" />");
text.AppendLine("    </PointByCoordinates>");
text.Append(points);
text.Append(figures);
text.AppendLine(string.Format(invariant,
    "    <ShowHideControl Name=\"Hint\" Style=\"GalleryText\" Show=\"false\" Text=\"Show the 4 × 4 box\" X=\"{0}\" Y=\"{1}\">",
    -side / 2, -side / 2 - 0.25));
text.AppendLine("      <Dependency Name=\"Four1\" />");
text.AppendLine("      <Dependency Name=\"Four2\" />");
text.AppendLine("    </ShowHideControl>");
text.AppendLine("    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"Seventeen Squares\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"57.5\" WrapWidth=\"400\" Backdrop=\"true\" />");
text.AppendLine("    <Label Name=\"Description\" Style=\"GalleryText\" Text=\""
    + "How small a box can hold 17 squares of side 1, without overlapping? Sixteen fill a 4 × 4 box exactly. "
    + "Add one more and the neat rows have to give way.\\n\\n"
    + "This is the best anyone has found, by John Bidwell in the 1990s: a box 4.6755 wide, with squares tilted at two different angles and gaps all over. "
    + "It looks wrong, as if a little shaking would let everything settle into something tidier, but every tidier arrangement needs a bigger box.\\n\\n"
    + "Is it the best possible? Nobody knows yet. A computer-assisted proof from 2026 shows that a box narrower than 4.66 can never work, which leaves a gap of less than 0.02.\\n\\n"
    + "Drag the yellow corners to turn and resize the box."
    + "\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"105.5\" WrapWidth=\"400\" Backdrop=\"true\" />");
text.AppendLine("  </Figures>");
text.Append("</Drawing>");
File.WriteAllText(args[0], text.ToString().Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(false));
Console.WriteLine("wrote " + args[0]);
return 0;
