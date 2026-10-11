#:property Nullable=disable
#:property PublishAot=false

// spindie - writes the "Spin the Die" gallery drawing: a glass die whose eight corners and
// twenty-one pips are points by coordinates turned by two sliders (about the vertical axis,
// then tipped), projected straight onto the paper. The six faces are translucent polygons,
// so the far ones show through and no face needs to be painted over another.
//
//   dotnet tools/spindie.cs -- <out.lgf>

using System.Globalization;
using System.Text;

if (args.Length < 1)
{
    Console.WriteLine("usage: spindie <out.lgf>");
    return 1;
}

// the die's half side, in units; the sliders are in degrees, so the die is big
const double Half = 70;
const double PipInset = 0.55;
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

// a point (x, y, z) in the die's own frame, turned by the sliders and projected: first
// about the vertical axis by "turn", then tipped towards the viewer by "tilt"
string ProjectX(double x, double y, double z)
{
    return $"{Format(Half * x)} * cos(rad(turn)) + {Format(Half * z)} * sin(rad(turn))";
}

string ProjectY(double x, double y, double z)
{
    return $"{Format(Half * y)} * cos(rad(tilt)) - ({Format(-Half * x)} * sin(rad(turn)) + {Format(Half * z)} * cos(rad(turn))) * sin(rad(tilt))";
}

void WritePoint(string name, string style, double x, double y, double z)
{
    string visible = style == null ? " Visible=\"false\"" : $" Style=\"{style}\"";
    Write($"    <PointByCoordinates Name=\"{name}\"{visible} X=\"{ProjectX(x, y, z)}\" Y=\"{ProjectY(x, y, z)}\">");
    Write("      <Dependency Name=\"turn\" />");
    Write("      <Dependency Name=\"tilt\" />");
    Write("    </PointByCoordinates>");
}

Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
Write("<Drawing Version=\"1\" Creator=\"LiveGeometry.App\">");
Write("  <Viewport Left=\"-200\" Top=\"150\" Right=\"200\" Bottom=\"-200\" />");
Write("  <Styles>");
// a face's tint and, deeper in the same hue, its pips: through the glass the pips of two
// faces lie over each other, and the color says which face each belongs to
var faces = new[]
{
    new Face("One", (0, 0, 1), 1, "#A0F28C8C", "#A0C04040", "#FFB81E1E", "#FFFF8080"),
    new Face("Six", (0, 0, -1), 6, "#A08CB8F2", "#A04060C0", "#FF1E4FB8", "#FF80A8FF"),
    new Face("Two", (0, 1, 0), 2, "#A0F2E08C", "#A0C0A030", "#FFA08A00", "#FFFFE060"),
    new Face("Five", (0, -1, 0), 5, "#A0A0E0A0", "#A040A050", "#FF1E8A3C", "#FF80E8A0"),
    new Face("Three", (1, 0, 0), 3, "#A0D0A0F0", "#A08050C0", "#FF7A2EB8", "#FFD0A0FF"),
    new Face("Four", (-1, 0, 0), 4, "#A0F2B88C", "#A0C07030", "#FFC05A10", "#FFFFB070"),
};
foreach (var face in faces)
{
    Write($"    <ShapeStyle Name=\"{face.Name}\" IsFilled=\"true\" Fill=\"{face.Fill}\" Color=\"#C0303030\" StrokeWidth=\"1.5\">");
    Write($"      <Dark Fill=\"{face.DarkFill}\" Color=\"#C0E0E0E0\" />");
    Write("    </ShapeStyle>");
    Write($"    <PointStyle Name=\"Pip{face.Name}\" Size=\"11\" Fill=\"{face.Pip}\" Color=\"#60000000\" StrokeWidth=\"1\">");
    Write($"      <Dark Fill=\"{face.DarkPip}\" Color=\"#60FFFFFF\" />");
    Write("    </PointStyle>");
}
Write("  </Styles>");
Write("  <Figures>");
Write("    <Slider Name=\"turn\" X=\"-190\" Y=\"-150\" Value=\"35\" Maximum=\"360\" />");
Write("    <Slider Name=\"tilt\" X=\"-190\" Y=\"-185\" Value=\"25\" Maximum=\"90\" />");

// the corners
var corners = new Dictionary<(int, int, int), string>();
foreach (int x in new[] { -1, 1 })
{
    foreach (int y in new[] { -1, 1 })
    {
        foreach (int z in new[] { -1, 1 })
        {
            string name = $"Corner{(x > 0 ? "R" : "L")}{(y > 0 ? "T" : "B")}{(z > 0 ? "F" : "K")}";
            corners[(x, y, z)] = name;
            WritePoint(name, null, x, y, z);
        }
    }
}

// the faces: the four corners on the side the normal points to, around the face
foreach (var face in faces)
{
    var (nx, ny, nz) = face.Normal;
    var around = new List<string>();
    foreach (var (a, b) in new[] { (-1, -1), (1, -1), (1, 1), (-1, 1) })
    {
        var corner = nx != 0 ? (nx, a, b) : ny != 0 ? (a, ny, b) : (a, b, nz);
        around.Add(corners[corner]);
    }

    Write($"    <Polygon Name=\"Face{face.Name}\" Style=\"{face.Name}\">");
    foreach (var corner in around)
    {
        Write($"      <Dependency Name=\"{corner}\" />");
    }

    Write("    </Polygon>");
}

// the pips, a little in front of each face so that they draw over it
foreach (var face in faces)
{
    var (nx, ny, nz) = face.Normal;
    int index = 0;
    foreach (var (a, b) in Pips(face.Pips))
    {
        double u = a * PipInset;
        double v = b * PipInset;
        var (x, y, z) = nx != 0 ? (nx * 1.001, u, v) : ny != 0 ? (u, ny * 1.001, v) : (u, v, nz * 1.001);
        WritePoint($"Pip{face.Name}{index}", "Pip" + face.Name, x, y, z);
        index++;
    }
}

Write("    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"Spin the Die\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"A glass die. Drag the sliders to turn it around and tip it over, and look at it from every side.\\n\\nThe die is really twenty-nine points, eight corners and twenty-one pips, each with three coordinates. The sliders turn them in space, and the picture is just the first two coordinates of each, with the third one thrown away. That is all a 3D picture on a flat screen ever is.\\n\\nEach face has its own color, pips included, so you can tell them apart through the glass. Turn it so you can see the 1 and the 6 at the same time: they are on opposite faces, and so are 2 and 5, and 3 and 4. Every opposite pair adds up to 7. Dice have been made this way for thousands of years.\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("  </Figures>");
text.Append("</Drawing>");
File.WriteAllText(args[0], text.ToString(), new UTF8Encoding(false));
Console.WriteLine("wrote " + args[0]);
return 0;

static IEnumerable<(int, int)> Pips(int count)
{
    switch (count)
    {
        case 1:
            return new[] { (0, 0) };
        case 2:
            return new[] { (-1, -1), (1, 1) };
        case 3:
            return new[] { (-1, -1), (0, 0), (1, 1) };
        case 4:
            return new[] { (-1, -1), (1, 1), (-1, 1), (1, -1) };
        case 5:
            return new[] { (-1, -1), (1, 1), (-1, 1), (1, -1), (0, 0) };
        default:
            return new[] { (-1, -1), (-1, 0), (-1, 1), (1, -1), (1, 0), (1, 1) };
    }
}

record Face(string Name, (int X, int Y, int Z) Normal, int Pips, string Fill, string DarkFill, string Pip, string DarkPip);
