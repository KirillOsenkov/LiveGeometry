#:property Nullable=disable
#:property PublishAot=false

// flower - writes the "Flower of Life" gallery drawing: nineteen circles of one size, the
// center of each on the rims of its neighbors, inside a ring. Two free points, the center O
// and R on the first circle, are the only inputs; every other center is worked out from
// them, so the whole flower turns and grows with R.
//
//   dotnet tools/flower.cs -- <out.lgf>

using System.Globalization;
using System.Text;

if (args.Length < 1)
{
    Console.WriteLine("usage: flower <out.lgf>");
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
    return value.ToString("0.######", invariant);
}

// the centers, in the frame where O is the origin and R is (1, 0)
var centers = new List<(double X, double Y)> { (0, 0) };
for (int k = 0; k < 6; k++)
{
    double angle = k * Math.PI / 3;
    centers.Add((Math.Cos(angle), Math.Sin(angle)));
}

for (int k = 0; k < 6; k++)
{
    double angle = k * Math.PI / 3;
    centers.Add((2 * Math.Cos(angle), 2 * Math.Sin(angle)));
    double between = angle + Math.PI / 6;
    centers.Add((Math.Sqrt(3) * Math.Cos(between), Math.Sqrt(3) * Math.Sin(between)));
}

// a point at (x, y) in that frame as an expression over O and R
string X(double x, double y)
{
    return $"O.X + {Format(x)} * (R.X - O.X) - {Format(y)} * (R.Y - O.Y)";
}

string Y(double x, double y)
{
    return $"O.Y + {Format(x)} * (R.Y - O.Y) + {Format(y)} * (R.X - O.X)";
}

void WritePoint(string name, double x, double y)
{
    Write($"    <PointByCoordinates Name=\"{name}\" Visible=\"false\" X=\"{X(x, y)}\" Y=\"{Y(x, y)}\">");
    Write("      <Dependency Name=\"O\" />");
    Write("      <Dependency Name=\"R\" />");
    Write("    </PointByCoordinates>");
}

Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
Write("<Drawing Version=\"1\" Creator=\"LiveGeometry.App\">");
Write("  <Viewport Left=\"-4\" Top=\"3.6\" Right=\"4\" Bottom=\"-3.6\" />");
Write("  <Styles>");
Write("    <ShapeStyle Name=\"Petal\" IsFilled=\"true\" Fill=\"#2A2F7BD6\" Color=\"#FF2F7BD6\" StrokeWidth=\"1.2\">");
Write("      <Dark Fill=\"#2A62AEFF\" Color=\"#FF62AEFF\" />");
Write("    </ShapeStyle>");
Write("    <ShapeStyle Name=\"Ring\" IsFilled=\"false\" Color=\"#FF2F7BD6\" StrokeWidth=\"3\">");
Write("      <Dark Color=\"#FF62AEFF\" />");
Write("    </ShapeStyle>");
Write("    <ShapeStyle Name=\"Ring2\" IsFilled=\"false\" Color=\"#FF2F7BD6\" StrokeWidth=\"1.2\">");
Write("      <Dark Color=\"#FF62AEFF\" />");
Write("    </ShapeStyle>");
Write("  </Styles>");
Write("  <Figures>");
Write("    <FreePoint Name=\"O\" X=\"0\" Y=\"0\" />");
Write("    <FreePoint Name=\"R\" X=\"1\" Y=\"0\" />");
for (int k = 0; k < centers.Count; k++)
{
    var (x, y) = centers[k];
    string center = k == 0 ? "O" : $"C{k}";
    string rim = k == 0 ? "R" : $"R{k}";
    if (k > 0)
    {
        WritePoint(center, x, y);
        WritePoint(rim, x + 1, y);
    }

    Write($"    <Circle Name=\"Petal{k}\" Style=\"Petal\">");
    Write($"      <Dependency Name=\"{center}\" />");
    Write($"      <Dependency Name=\"{rim}\" />");
    Write("    </Circle>");
}

WritePoint("Edge", 3, 0);
Write("    <Circle Name=\"Ring\" Style=\"Ring\">");
Write("      <Dependency Name=\"O\" />");
Write("      <Dependency Name=\"Edge\" />");
Write("    </Circle>");
WritePoint("Edge2", 3.12, 0);
Write("    <Circle Name=\"Ring2\" Style=\"Ring2\">");
Write("      <Dependency Name=\"O\" />");
Write("      <Dependency Name=\"Edge2\" />");
Write("    </Circle>");
Write("    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"Flower of Life\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"Draw a circle. Draw another one of the same size with its center on the first one's rim. Where the two cross, put the center of a third. Keep going: nineteen circles later you have this.\\n\\nIt works only because six circles fit exactly around one, with no gap and no overlap, which is true because six equilateral triangles fit exactly around a point. The same fact makes honeycombs hexagonal.\\n\\nThe pattern is at least 2,600 years old: it is carved into temple pillars in Egypt and Assyria, and Leonardo da Vinci filled pages of his notebooks with it. Drag R to resize it, O to move it.\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("  </Figures>");
text.Append("</Drawing>");
File.WriteAllText(args[0], text.ToString(), new UTF8Encoding(false));
Console.WriteLine("wrote " + args[0]);
return 0;
