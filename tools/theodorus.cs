#:property Nullable=disable
#:property PublishAot=false

// theodorus - writes the "Spiral of Theodorus" gallery drawing: sixteen right triangles, each
// standing on the last one's hypotenuse with a leg of 1, so the hypotenuses are the square
// roots of 2, 3, 4... 17. The two free points A (the center) and B (the first corner) are the
// first leg; every other corner is worked out from them, so dragging either turns and resizes
// the whole spiral.
//
//   dotnet tools/theodorus.cs -- <out.lgf>

using System.Globalization;
using System.Text;

if (args.Length < 1)
{
    Console.WriteLine("usage: theodorus <out.lgf>");
    return 1;
}

const int Triangles = 16;
var invariant = CultureInfo.InvariantCulture;
var text = new StringBuilder();

void Write(string line)
{
    text.Append(line).Append("\r\n");
}

string Format(double value)
{
    return value.ToString("0.######", invariant);
}

// the corners in the frame where A is the origin and B is (1, 0): each next corner is the
// last one plus the unit vector at right angles to it
var corners = new List<(double X, double Y)> { (1, 0) };
for (int k = 1; k < Triangles; k++)
{
    var (x, y) = corners[k - 1];
    double length = Math.Sqrt(x * x + y * y);
    corners.Add((x - y / length, y + x / length));
}

// one more for the last hypotenuse
{
    var (x, y) = corners[Triangles - 1];
    double length = Math.Sqrt(x * x + y * y);
    corners.Add((x - y / length, y + x / length));
}

Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
Write("<Drawing Version=\"1\" Creator=\"LiveGeometry.App\">");
Write("  <Viewport Left=\"-5\" Top=\"4.5\" Right=\"5\" Bottom=\"-4.5\" />");
Write("  <Styles>");
for (int k = 0; k < Triangles; k++)
{
    double hue = 360.0 * k / Triangles;
    Write($"    <ShapeStyle Name=\"Wedge{k + 1}\" IsFilled=\"true\" Fill=\"{Hsl(hue, 0.75, 0.78)}\" Color=\"#FF2B3038\" StrokeWidth=\"1.5\">");
    Write($"      <Dark Fill=\"{Hsl(hue, 0.55, 0.42)}\" Color=\"#FFE4E8F0\" />");
    Write("    </ShapeStyle>");
}

Write("    <TextStyle Name=\"Root\" FontSize=\"13\" Color=\"#FF2B3038\" FontFamily=\"Segoe UI\">");
Write("      <Dark Color=\"#FFE4E8F0\" />");
Write("    </TextStyle>");
Write("  </Styles>");
Write("  <Figures>");
Write("    <FreePoint Name=\"A\" X=\"0.3\" Y=\"-0.4\" />");
Write("    <FreePoint Name=\"B\" X=\"1.3\" Y=\"-0.4\" />");
for (int k = 1; k < corners.Count; k++)
{
    var (x, y) = corners[k];
    Write($"    <PointByCoordinates Name=\"P{k + 1}\" Visible=\"false\" X=\"A.X + {Format(x)} * (B.X - A.X) - {Format(y)} * (B.Y - A.Y)\" Y=\"A.Y + {Format(x)} * (B.Y - A.Y) + {Format(y)} * (B.X - A.X)\">");
    Write("      <Dependency Name=\"A\" />");
    Write("      <Dependency Name=\"B\" />");
    Write("    </PointByCoordinates>");
}

for (int k = 0; k < Triangles; k++)
{
    string near = k == 0 ? "B" : $"P{k + 1}";
    string far = $"P{k + 2}";
    Write($"    <Polygon Name=\"Triangle{k + 1}\" Style=\"Wedge{k + 1}\">");
    Write("      <Dependency Name=\"A\" />");
    Write($"      <Dependency Name=\"{near}\" />");
    Write($"      <Dependency Name=\"{far}\" />");
    Write("    </Polygon>");
}

// the roots along the hypotenuses: a hidden point halfway, named after the root (a trailing
// space keeps the digits off the subscript line), its name shown
for (int k = 0; k < Triangles; k++)
{
    var (x, y) = corners[k + 1];
    string name = $"√{k + 2} ";
    Write($"    <PointByCoordinates Name=\"{name}\" Visible=\"false\" X=\"A.X + {Format(x / 2)} * (B.X - A.X) - {Format(y / 2)} * (B.Y - A.Y)\" Y=\"A.Y + {Format(x / 2)} * (B.Y - A.Y) + {Format(y / 2)} * (B.X - A.X)\">");
    Write("      <Dependency Name=\"A\" />");
    Write("      <Dependency Name=\"B\" />");
    Write("    </PointByCoordinates>");
    Write($"    <PointLabel Name=\"Root{k + 2}\" IsHitTestVisible=\"false\" Style=\"Root\" OffsetX=\"-11\" OffsetY=\"-9\" ShowName=\"true\" ShowCoordinates=\"false\">");
    Write($"      <Dependency Name=\"{name}\" />");
    Write("    </PointLabel>");
}

Write("    <PointLabel Name=\"LabelA\" OffsetX=\"-16\" OffsetY=\"-6\" ShowName=\"true\" ShowCoordinates=\"false\">");
Write("      <Dependency Name=\"A\" />");
Write("    </PointLabel>");
Write("    <PointLabel Name=\"LabelB\" OffsetX=\"8\" OffsetY=\"-4\" ShowName=\"true\" ShowCoordinates=\"false\">");
Write("      <Dependency Name=\"B\" />");
Write("    </PointLabel>");
Write("    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"Spiral of Theodorus\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"Start with a right triangle whose two short sides are both 1. Its long side is √2. Now build a new right triangle on that long side, with the other short side 1 again: its long side is √3. Keep going.\\n\\nEvery long side is the square root of the next whole number, by the Pythagorean theorem, and no ruler had to measure anything. Theodorus of Cyrene drew this about 2,400 years ago, and stopped at √17, where the next triangle would start to overlap the first one.\\n\\nDrag A or B to turn and resize the spiral. Every length in it scales with AB, which is 1.\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("  </Figures>");
text.Append("</Drawing>");
File.WriteAllText(args[0], text.ToString(), new UTF8Encoding(false));
Console.WriteLine("wrote " + args[0]);
return 0;

static string Hsl(double hue, double saturation, double lightness)
{
    double c = (1 - Math.Abs(2 * lightness - 1)) * saturation;
    double x = c * (1 - Math.Abs(hue / 60 % 2 - 1));
    double m = lightness - c / 2;
    (double r, double g, double b) = hue switch
    {
        < 60 => (c, x, 0.0),
        < 120 => (x, c, 0.0),
        < 180 => (0.0, c, x),
        < 240 => (0.0, x, c),
        < 300 => (x, 0.0, c),
        _ => (c, 0.0, x),
    };
    return "#FF" + Channel(r + m) + Channel(g + m) + Channel(b + m);
}

static string Channel(double value)
{
    return ((int)Math.Round(value * 255)).ToString("X2");
}
