#:property Nullable=disable
#:property PublishAot=false

// gears - writes the "Gears" gallery drawing: three gears that mesh, each a polygon whose
// vertices are points by coordinates turned by the slider "spin" (turns of the first gear),
// the second gear the other way round and slower by the ratio of the teeth, the third the
// first way again. The phases are worked out so that a tooth of one always sits in a gap of
// the next.
//
//   dotnet tools/gears.cs -- <out.lgf>

using System.Globalization;
using System.Text;

if (args.Length < 1)
{
    Console.WriteLine("usage: gears <out.lgf>");
    return 1;
}

// the module: the pitch radius is Module * teeth / 2
const double Module = 0.3;
const double Addendum = Module;
const double Dedendum = Module * 1.15;

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

// the gears, each meshing with the one before: teeth, the direction (degrees) from the
// previous gear's center to this one's, the default style (the palette's outlined
// gradients, which the file needn't carry)
var gears = new List<Gear>
{
    new Gear("Gold", 12, 0, "GradientBrownOutline"),
    new Gear("Green", 20, 0, "GradientGreenOutline"),
    new Gear("Purple", 8, -50, "GradientPurpleOutline"),
};

// place them, and work out the phases
gears[0].Center = (0, 0);
gears[0].Phase = 0;
gears[0].Sign = 1;
for (int g = 1; g < gears.Count; g++)
{
    var previous = gears[g - 1];
    var gear = gears[g];
    double direction = gear.Direction * Math.PI / 180;
    double distance = previous.Radius + gear.Radius;
    gear.Center = (previous.Center.X + distance * Math.Cos(direction), previous.Center.Y + distance * Math.Sin(direction));
    gear.Sign = -previous.Sign;

    // where the previous gear is in its tooth cycle along the direction (0 = a tooth's
    // middle, 0.5 = a gap's). Along the line where the two rims touch, the previous gear's
    // cycle runs one way and this gear's the other, so this one must be half a tooth
    // further the other way round
    double previousCycle = (gear.Direction - previous.Phase) * previous.Teeth / 360;
    gear.Phase = gear.Direction + 180 - (0.5 - previousCycle) * 360 / gear.Teeth;
}

Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
Write("<Drawing Version=\"1\" Creator=\"LiveGeometry.App\">");
Write("  <Viewport Left=\"-4\" Top=\"4.5\" Right=\"10\" Bottom=\"-5\" />");
Write("  <Styles>");
Write("    <ShapeStyle Name=\"Axle\" IsFilled=\"true\" Fill=\"#FFFFFFFF\" Color=\"#80000000\" StrokeWidth=\"2\">");
Write("      <Dark Fill=\"#FF2A2E38\" Color=\"#80FFFFFF\" />");
Write("    </ShapeStyle>");
Write("    <PointStyle Name=\"Mark\" Size=\"12\" Fill=\"#FFFFFFFF\" Color=\"#80000000\" StrokeWidth=\"1.5\">");
Write("      <Dark Fill=\"#FF2A2E38\" Color=\"#80FFFFFF\" />");
Write("    </PointStyle>");
Write("  </Styles>");
Write("  <Figures>");
Write("    <Slider Name=\"spin\" X=\"-3.5\" Y=\"-4.6\" Value=\"1\" />");

foreach (var gear in gears)
{
    // the angle of the gear as an expression: its phase plus its share of the spin (the
    // slider runs 3 units per turn of the first gear, so that it is long enough to drag)
    double speed = gear.Sign * 120.0 * gears[0].Teeth / gear.Teeth;
    string angle = $"{Format(gear.Phase)} + {Format(speed)} * spin";
    double halfPitch = 180.0 / gear.Teeth;
    var vertices = new List<string>();
    for (int tooth = 0; tooth < gear.Teeth; tooth++)
    {
        double middle = 360.0 * tooth / gear.Teeth;
        var corners = new[]
        {
            (middle - 0.5 * halfPitch, gear.Radius - Dedendum),
            (middle - 0.25 * halfPitch, gear.Radius + Addendum),
            (middle + 0.25 * halfPitch, gear.Radius + Addendum),
            (middle + 0.5 * halfPitch, gear.Radius - Dedendum),
        };
        for (int c = 0; c < corners.Length; c++)
        {
            string name = $"{gear.Name}{tooth}_{c}";
            var (offset, radius) = corners[c];
            Write($"    <PointByCoordinates Name=\"{name}\" Visible=\"false\" X=\"{Format(gear.Center.X)} + {Format(radius)} * cos(rad({angle} + {Format(offset)}))\" Y=\"{Format(gear.Center.Y)} + {Format(radius)} * sin(rad({angle} + {Format(offset)}))\">");
            Write("      <Dependency Name=\"spin\" />");
            Write("    </PointByCoordinates>");
            vertices.Add(name);
        }
    }

    Write($"    <Polygon Name=\"{gear.Name}Gear\" Style=\"{gear.Style}\">");
    foreach (var vertex in vertices)
    {
        Write($"      <Dependency Name=\"{vertex}\" />");
    }

    Write("    </Polygon>");

    // the axle, and a mark that turns with the gear so that the turning shows
    Write($"    <PointByCoordinates Name=\"{gear.Name}Axle\" Visible=\"false\" X=\"{Format(gear.Center.X)}\" Y=\"{Format(gear.Center.Y)}\" />");
    Write($"    <PointByCoordinates Name=\"{gear.Name}AxleRim\" Visible=\"false\" X=\"{Format(gear.Center.X + gear.Radius * 0.22)}\" Y=\"{Format(gear.Center.Y)}\" />");
    Write($"    <Circle Name=\"{gear.Name}Hole\" Style=\"Axle\">");
    Write($"      <Dependency Name=\"{gear.Name}Axle\" />");
    Write($"      <Dependency Name=\"{gear.Name}AxleRim\" />");
    Write("    </Circle>");
    Write($"    <PointByCoordinates Name=\"{gear.Name}Mark\" Style=\"Mark\" X=\"{Format(gear.Center.X)} + {Format(gear.Radius * 0.6)} * cos(rad({angle}))\" Y=\"{Format(gear.Center.Y)} + {Format(gear.Radius * 0.6)} * sin(rad({angle}))\">");
    Write("      <Dependency Name=\"spin\" />");
    Write("    </PointByCoordinates>");
}

Write("    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"Gears\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"Drag the slider to turn the yellow gear. Its teeth push the green gear the other way, and the green gear pushes the purple one back the first way: every gear in a chain turns against its neighbors.\\n\\nThe yellow gear has 12 teeth and the green one 20. The yellow gear has made [spin / 3] turns, so the green one has made [12 / 20 * spin / 3]: a big gear turns slower than a small one, by exactly the ratio of their teeth. The purple gear, with 8 teeth, has made [12 / 8 * spin / 3] turns. Watch the white dots.\\n\\nThat is how a bicycle works. Pedal a 48-tooth chainring driving a 12-tooth sprocket and the back wheel spins four times for every turn of your feet.\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\">");
Write("      <Dependency Name=\"spin\" />");
Write("    </Label>");
Write("  </Figures>");
text.Append("</Drawing>");
File.WriteAllText(args[0], text.ToString(), new UTF8Encoding(false));
Console.WriteLine("wrote " + args[0]);
return 0;

class Gear
{
    public Gear(string name, int teeth, double direction, string style)
    {
        Name = name;
        Teeth = teeth;
        Direction = direction;
        Style = style;
    }

    public string Name { get; }

    public int Teeth { get; }

    /// <summary>Degrees from the previous gear's center to this one's</summary>
    public double Direction { get; }

    /// <summary>The name of a default shape style</summary>
    public string Style { get; }

    public double Radius => Module * Teeth / 2;

    public (double X, double Y) Center { get; set; }

    /// <summary>Degrees the gear is turned by at spin 0</summary>
    public double Phase { get; set; }

    /// <summary>1 with the first gear, -1 against it</summary>
    public int Sign { get; set; }

    const double Module = 0.3;
}
