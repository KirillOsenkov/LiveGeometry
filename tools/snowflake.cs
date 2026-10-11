#:property Nullable=disable
#:property PublishAot=false

// snowflake - writes the "Design a Snowflake" gallery drawing: one twelfth of a snowflake is
// a closed Bezier path with automatic handles over a few free points, between the arm's
// axis and the line 30 degrees from it. Its mirror image in the axis completes one arm, and
// five rotations by 60 degrees make the other arms: twelve paths over reflected and rotated
// points, which smooth themselves the same way as the original, since Hobby's handles
// don't care which way a figure is turned.
//
//   dotnet tools/snowflake.cs -- <out.lgf>

using System.Globalization;
using System.Text;

if (args.Length < 1)
{
    Console.WriteLine("usage: snowflake <out.lgf>");
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

// the half arm, from the tip back to the junction with the next arm; T stays on the axis
// and J on the 30 degree line, the rest are free
var free = new[]
{
    ("A1", (2.55, 0.42)),
    ("A2", (2.05, 0.2)),
    ("A3", (1.75, 0.78)),
    ("A4", (1.2, 0.38)),
};
string[] shape = { "T", "A1", "A2", "A3", "A4", "J" };

Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
Write("<Drawing Version=\"1\" Creator=\"LiveGeometry.App\">");
Write("  <Viewport Left=\"-4.2\" Top=\"4\" Right=\"4.2\" Bottom=\"-4\">");
Write("    <Background>");
Write("      <LinearGradientBrush StartPoint=\"0,0\" EndPoint=\"0,1\">");
Write("        <GradientStop Color=\"#FF0E2A5C\" Offset=\"0\" />");
Write("        <GradientStop Color=\"#FF1A4A8C\" Offset=\"1\" />");
Write("      </LinearGradientBrush>");
Write("    </Background>");
Write("  </Viewport>");
Write("  <Styles>");
Write("    <TextStyle Name=\"GalleryTitle\" FontSize=\"30\" Color=\"#FFE8F4FF\" FontFamily=\"Segoe UI\" Bold=\"true\">");
Write("      <Dark Color=\"#FFE8F4FF\" />");
Write("    </TextStyle>");
Write("    <TextStyle Name=\"GalleryText\" FontSize=\"16\" Color=\"#FFD6E6FA\" FontFamily=\"Segoe UI\">");
Write("      <Dark Color=\"#FFD6E6FA\" />");
Write("    </TextStyle>");
Write("    <ShapeStyle Name=\"Ice\" IsFilled=\"true\" Fill=\"#E6FFFFFF\" Color=\"#00000000\" StrokeWidth=\"0.5\" />");
Write("    <LineStyle Name=\"IceEdge\" Color=\"#FFBFE6FF\" StrokeWidth=\"1\" />");
Write("    <LineStyle Name=\"Guide\" Color=\"#50FFFFFF\" StrokeWidth=\"1\" Dash=\"Dash\" />");
Write("    <PointStyle Name=\"Node\" Size=\"11\" Fill=\"#FF4FA3FF\" Color=\"#FFFFFFFF\" StrokeWidth=\"2\" />");
Write("  </Styles>");
Write("  <Figures>");
Write("    <PointByCoordinates Name=\"O\" Visible=\"false\" X=\"0\" Y=\"0\" />");
Write("    <PointByCoordinates Name=\"AxisEnd\" Visible=\"false\" X=\"1\" Y=\"0\" />");
Write("    <Ray Name=\"Axis\" Style=\"Guide\">");
Write("      <Dependency Name=\"O\" />");
Write("      <Dependency Name=\"AxisEnd\" />");
Write("    </Ray>");
Write("    <LineTwoPoints Name=\"Mirror\" Visible=\"false\">");
Write("      <Dependency Name=\"O\" />");
Write("      <Dependency Name=\"AxisEnd\" />");
Write("    </LineTwoPoints>");
Write($"    <PointByCoordinates Name=\"SideEnd\" Visible=\"false\" X=\"{Format(Math.Cos(Math.PI / 6))}\" Y=\"{Format(Math.Sin(Math.PI / 6))}\" />");
Write("    <Ray Name=\"SideRay\" Style=\"Guide\">");
Write("      <Dependency Name=\"O\" />");
Write("      <Dependency Name=\"SideEnd\" />");
Write("    </Ray>");
Write("    <PointOnFigure Name=\"T\" Style=\"Node\" X=\"3.2\" Y=\"0\" Parameter=\"3.2\">");
Write("      <Dependency Name=\"Axis\" />");
Write("    </PointOnFigure>");
foreach (var (name, (x, y)) in free)
{
    Write($"    <FreePoint Name=\"{name}\" Style=\"Node\" X=\"{Format(x)}\" Y=\"{Format(y)}\" />");
}

Write($"    <PointOnFigure Name=\"J\" Style=\"Node\" X=\"{Format(0.9 * Math.Cos(Math.PI / 6))}\" Y=\"{Format(0.9 * Math.Sin(Math.PI / 6))}\" Parameter=\"0.9\">");
Write("      <Dependency Name=\"SideRay\" />");
Write("    </PointOnFigure>");
foreach (int turn in new[] { 60, 120, 180, 240, 300 })
{
    Write($"    <Number Name=\"Turn{turn}\" Value=\"{turn}\" />");
}

// the mirror images, then the rotations of both halves
foreach (var name in shape)
{
    if (name == "T")
    {
        continue;
    }

    Write($"    <ReflectedPoint Name=\"{name}m\" Visible=\"false\">");
    Write($"      <Dependency Name=\"{name}\" />");
    Write("      <Dependency Name=\"Mirror\" />");
    Write("    </ReflectedPoint>");
}

string path = "L " + string.Join(" ", Enumerable.Repeat("C a a", shape.Length - 1)) + " L";

void WritePath(string name, IEnumerable<string> points)
{
    Write($"    <BezierPath Name=\"{name}\" Style=\"Ice\" Closed=\"true\" Filled=\"true\" Path=\"{path}\">");
    Write("      <Sides Style=\"IceEdge\" />");
    Write("      <Dependency Name=\"O\" />");
    foreach (var point in points)
    {
        Write($"      <Dependency Name=\"{point}\" />");
    }

    Write("    </BezierPath>");
}

string Mirrored(string name)
{
    return name == "T" ? name : name + "m";
}

WritePath("Half", shape);
WritePath("HalfM", shape.Select(Mirrored));
// (the tip is on the mirror, so both halves of an arm share its rotated copy)
var written = new HashSet<string>();
foreach (int turn in new[] { 60, 120, 180, 240, 300 })
{
    foreach (bool mirrored in new[] { false, true })
    {
        var points = new List<string>();
        foreach (var name in shape)
        {
            string source = mirrored ? Mirrored(name) : name;
            string rotated = $"{source}r{turn}";
            if (written.Add(rotated))
            {
                Write($"    <RotatedPoint Name=\"{rotated}\" Visible=\"false\">");
                Write($"      <Dependency Name=\"{source}\" />");
                Write("      <Dependency Name=\"O\" />");
                Write($"      <Dependency Name=\"Turn{turn}\" />");
                Write("    </RotatedPoint>");
            }

            points.Add(rotated);
        }

        WritePath($"Arm{turn}{(mirrored ? "M" : "")}", points);
    }
}

Write("    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"Design a Snowflake\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"Every snowflake has six arms, and every arm is the same on both sides. So a snowflake is really one small shape, drawn twelve times: mirrored, then turned five times by 60 degrees.\\n\\nDrag the blue points. You are drawing one twelfth, between the two dashed lines, and the other eleven follow. The tip slides along the arm, the last point along the line where two arms meet.\\n\\nReal snowflakes grow their six arms at once in the same cloud, in the same air, which is why the arms match. No two snowflakes take the same trip down, which is why no two flakes match. Can you make one that looks like a real snowflake? One that looks like a flower?\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("  </Figures>");
text.Append("</Drawing>");
File.WriteAllText(args[0], text.ToString(), new UTF8Encoding(false));
Console.WriteLine("wrote " + args[0]);
return 0;
