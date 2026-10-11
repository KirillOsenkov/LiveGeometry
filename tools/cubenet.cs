#:property Nullable=disable
#:property PublishAot=false

// cubenet - writes the "Fold the Cube" gallery drawing: the cross-shaped net of a cube, lying
// on the ground, whose four flaps fold up by the angle of the slider and whose lid folds over
// from the top flap by the same angle again, so that at 90 degrees the box is closed. Every
// corner is a point by coordinates: its place in space as an expression of the slider,
// projected onto the paper from a fixed viewpoint.
//
//   dotnet tools/cubenet.cs -- <out.lgf>

using System.Globalization;
using System.Text;

if (args.Length < 1)
{
    Console.WriteLine("usage: cubenet <out.lgf>");
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

// the viewpoint: turned this much about the vertical, seen from this high
double turn = 35 * Math.PI / 180;
double elevation = 50 * Math.PI / 180;
string cosTurn = Format(Math.Cos(turn));
string sinTurn = Format(Math.Sin(turn));
string cosElevation = Format(Math.Cos(elevation));
string sinElevation = Format(Math.Sin(elevation));

// the fold angle: the slider runs 0 to 3, a right angle at 3
const string Angle = "rad(30 * fold)";
string C = $"cos({Angle})";
string S = $"sin({Angle})";
string C2 = $"cos(2 * {Angle})";
string S2 = $"sin(2 * {Angle})";

// a point in space (each coordinate an expression) projected onto the paper
void WritePoint(string name, string x, string y, string z)
{
    string screenX = $"({x}) * {cosTurn} - ({y}) * {sinTurn}";
    string screenY = $"(({x}) * {sinTurn} + ({y}) * {cosTurn}) * {sinElevation} + ({z}) * {cosElevation}";
    Write($"    <PointByCoordinates Name=\"{name}\" Visible=\"false\" X=\"{screenX}\" Y=\"{screenY}\">");
    Write("      <Dependency Name=\"fold\" />");
    Write("    </PointByCoordinates>");
}

Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
Write("<Drawing Version=\"1\" Creator=\"LiveGeometry.App\">");
Write("  <Viewport Left=\"-6\" Top=\"5\" Right=\"6\" Bottom=\"-5\" />");
Write("  <Styles>");
var faces = new (string Name, string Fill, string DarkFill)[]
{
    ("Base", "#C0D9A066", "#C0A07040"),
    ("East", "#C0F2B48C", "#C0C07848"),
    ("West", "#C0F2B48C", "#C0C07848"),
    ("South", "#C0E8C89A", "#C0B08C58"),
    ("North", "#C0E8C89A", "#C0B08C58"),
    ("Lid", "#C0C8A070", "#C0906040"),
};
foreach (var face in faces)
{
    Write($"    <ShapeStyle Name=\"{face.Name}\" IsFilled=\"true\" Fill=\"{face.Fill}\" Color=\"#FF6B4423\" StrokeWidth=\"2\">");
    Write($"      <Dark Fill=\"{face.DarkFill}\" Color=\"#FFE0C0A0\" />");
    Write("    </ShapeStyle>");
}

Write("    <LineStyle Name=\"Ground\" Color=\"#FF8A94A6\" StrokeWidth=\"1\" Dash=\"Dash\" />");
Write("  </Styles>");
Write("  <Figures>");
Write("    <Slider Name=\"fold\" X=\"-5.5\" Y=\"-4.2\" Value=\"1\" Maximum=\"3\" />");

// the base on the ground, corners (+-1, +-1, 0)
foreach (var (name, x, y) in new[] { ("BaseSW", "-1", "-1"), ("BaseSE", "1", "-1"), ("BaseNE", "1", "1"), ("BaseNW", "-1", "1") })
{
    WritePoint(name, x, y, "0");
}

// the flaps: each hinged on a side of the base, its far corners swung up by the angle
WritePoint("EastS", $"1 + 2 * {C}", "-1", $"2 * {S}");
WritePoint("EastN", $"1 + 2 * {C}", "1", $"2 * {S}");
WritePoint("WestS", $"-1 - 2 * {C}", "-1", $"2 * {S}");
WritePoint("WestN", $"-1 - 2 * {C}", "1", $"2 * {S}");
WritePoint("SouthW", "-1", $"-1 - 2 * {C}", $"2 * {S}");
WritePoint("SouthE", "1", $"-1 - 2 * {C}", $"2 * {S}");
WritePoint("NorthW", "-1", $"1 + 2 * {C}", $"2 * {S}");
WritePoint("NorthE", "1", $"1 + 2 * {C}", $"2 * {S}");

// the lid, hinged on the north flap's far edge, swung over by the angle again
WritePoint("LidW", "-1", $"1 + 2 * {C} + 2 * {C2}", $"2 * {S} + 2 * {S2}");
WritePoint("LidE", "1", $"1 + 2 * {C} + 2 * {C2}", $"2 * {S} + 2 * {S2}");

void WriteFace(string name, params string[] corners)
{
    Write($"    <Polygon Name=\"{name}Face\" Style=\"{name}\">");
    foreach (var corner in corners)
    {
        Write($"      <Dependency Name=\"{corner}\" />");
    }

    Write("    </Polygon>");
}

WriteFace("Base", "BaseSW", "BaseSE", "BaseNE", "BaseNW");
WriteFace("West", "BaseSW", "BaseNW", "WestN", "WestS");
WriteFace("South", "BaseSW", "BaseSE", "SouthE", "SouthW");
WriteFace("North", "BaseNW", "BaseNE", "NorthE", "NorthW");
WriteFace("Lid", "NorthW", "NorthE", "LidE", "LidW");
WriteFace("East", "BaseSE", "BaseNE", "EastN", "EastS");

Write("    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"Fold the Cube\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"Six squares in a cross, lying flat. Drag the slider and the four sides fold up, the lid folds over, and the cross becomes a box.\\n\\nA flat shape that folds into a solid is called a net. The cross is the famous net of the cube, but it isn't the only one: there are exactly eleven different ways to lay six squares out so that they fold into a cube. Could you draw another?\\n\\nThe paper is see-through so you can watch the far sides. Fold it all the way to see the cube closed, then back down to flat.\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("  </Figures>");
text.Append("</Drawing>");
File.WriteAllText(args[0], text.ToString(), new UTF8Encoding(false));
Console.WriteLine("wrote " + args[0]);
return 0;
