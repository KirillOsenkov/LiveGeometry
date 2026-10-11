#:property Nullable=disable
#:property PublishAot=false

// snowflake - writes the "Design a Snowflake" gallery drawing: one twelfth of a snowflake is
// a closed Bezier path with automatic handles over a few free points, between the arm's
// axis and the line 30 degrees from it. Its mirror image in the axis completes one arm, and
// five rotations by 60 degrees make the other arms: twelve paths over reflected and rotated
// points, which smooth themselves the same way as the original, since Hobby's handles
// don't care which way a figure is turned. Around it, on a snow bank and in the sky, emoji
// as free points: a snowman, a tree, a snowboarder, a cloud, sparkles and small flakes.
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

// the emoji around the flake: a style per look (several small flakes share one), and
// where each one starts
var Sprinkles = new[]
{
    ("Flake", "❄️", 26),
    ("BigFlake", "❄️", 38),
    ("TinyFlake", "❄️", 17),
    ("Snowman", "⛄", 64),
    ("Tree", "🌲", 56),
    ("Snowboarder", "🏂", 54),
    ("Cloud", "🌨️", 48),
    ("Sparkles", "✨", 28),
};
var SprinklePlaces = new[]
{
    ("Snowman", -3.4, -2.95),
    ("Tree", -1.6, -3.1),
    ("Snowboarder", 3.3, -2.85),
    ("Cloud", 3.6, 3.4),
    ("BigFlake", -3.7, 2.9),
    ("Flake", -4.1, 0.9),
    ("Flake", 4.1, 0.3),
    ("Flake", 1.0, 3.6),
    ("TinyFlake", -2.4, 3.6),
    ("TinyFlake", 2.6, 2.0),
    ("TinyFlake", -4.0, -1.3),
    ("TinyFlake", 4.2, 1.9),
    ("Sparkles", -1.1, 3.5),
    ("Sparkles", 3.9, -1.3),
    ("Sparkles", -3.4, -0.4),
};

Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
Write("<Drawing Version=\"1\" Creator=\"LiveGeometry.App\">");
Write("  <Viewport Left=\"-4.6\" Top=\"4\" Right=\"4.6\" Bottom=\"-4.2\">");
Write("    <Background>");
Write("      <LinearGradientBrush StartPoint=\"0,0\" EndPoint=\"0,1\">");
Write("        <GradientStop Color=\"#FF0E2A5C\" Offset=\"0\" />");
Write("        <GradientStop Color=\"#FF1A4A8C\" Offset=\"1\" />");
Write("      </LinearGradientBrush>");
Write("    </Background>");
Write("  </Viewport>");

// the view is the scene, not the content: the snow bank reaches far out of every window
Write("  <Scene Left=\"-4.6\" Top=\"4\" Right=\"4.6\" Bottom=\"-4.2\" />");
Write("  <Styles>");
Write("    <TextStyle Name=\"GalleryTitle\" FontSize=\"30\" Color=\"#FFE8F4FF\" FontFamily=\"Segoe UI\" Bold=\"true\">");
Write("      <Dark Color=\"#FFE8F4FF\" />");
Write("    </TextStyle>");
Write("    <TextStyle Name=\"GalleryText\" FontSize=\"16\" Color=\"#FFD6E6FA\" FontFamily=\"Segoe UI\">");
Write("      <Dark Color=\"#FFD6E6FA\" />");
Write("    </TextStyle>");
Write("    <ShapeStyle Name=\"Ice\" IsFilled=\"true\" Color=\"#00000000\" StrokeWidth=\"0.5\">");
Write("      <Fill>");
Write("        <LinearGradientBrush StartPoint=\"0,0\" EndPoint=\"1,1\">");
Write("          <GradientStop Color=\"#FFFFFFFF\" Offset=\"0\" />");
Write("          <GradientStop Color=\"#FFD2E6FF\" Offset=\"1\" />");
Write("        </LinearGradientBrush>");
Write("      </Fill>");
Write("    </ShapeStyle>");
Write("    <LineStyle Name=\"IceEdge\" Color=\"#FF4FA3FF\" StrokeWidth=\"1.5\" />");
Write("    <LineStyle Name=\"Guide\" Color=\"#C04FA3FF\" StrokeWidth=\"1.25\" Dash=\"Dash\" />");
Write("    <PointStyle Name=\"Node\" Size=\"11\" Fill=\"#FF4FA3FF\" Color=\"#FFFFFFFF\" StrokeWidth=\"2\" />");
Write("    <ShapeStyle Name=\"Snow\" IsFilled=\"true\" Color=\"#00000000\" StrokeWidth=\"0.5\">");
Write("      <Fill>");
Write("        <LinearGradientBrush StartPoint=\"0,0\" EndPoint=\"0,1\">");
Write("          <GradientStop Color=\"#FFFFFFFF\" Offset=\"0\" />");
Write("          <GradientStop Color=\"#FFC8DCF5\" Offset=\"1\" />");
Write("        </LinearGradientBrush>");
Write("      </Fill>");
Write("    </ShapeStyle>");
Write("    <LineStyle Name=\"NoLine\" Color=\"#00FFFFFF\" StrokeWidth=\"0.5\" />");
foreach (var (style, character, size) in Sprinkles.Distinct())
{
    Write($"    <PointStyle Name=\"{style}\" Character=\"{character}\" Size=\"{size}\" Fill=\"#FFFFFFFF\" />");
}

Write("  </Styles>");
Write("  <Figures>");

// a snow bank along the bottom, a path with automatic handles over fixed points, for the
// snowman and his friends to stand on
// (wavy across the view, straight far beyond it: curved through the far points it made a
// hill behind the caption)
var bank = new[] { (-40.0, -3.4), (-5.0, -3.4), (-2.6, -3.25), (0.0, -3.6), (2.6, -3.3), (5.0, -3.4), (40.0, -3.4), (40.0, -40.0), (-40.0, -40.0) };
for (int i = 0; i < bank.Length; i++)
{
    Write($"    <PointByCoordinates Name=\"Bank{i}\" Visible=\"false\" X=\"{Format(bank[i].Item1)}\" Y=\"{Format(bank[i].Item2)}\" />");
}

Write("    <BezierPath Name=\"SnowBank\" Style=\"Snow\" Closed=\"true\" Filled=\"true\" Path=\"L C a a C a a C a a C a a L L L L\">");
Write("      <Sides Style=\"NoLine\" />");
for (int i = 0; i < bank.Length; i++)
{
    Write($"      <Dependency Name=\"Bank{i}\" />");
}

Write("    </BezierPath>");

// the sprinkles: free points drawn as emoji, so they can be moved around
int sprinkle = 0;
foreach (var (style, x, y) in SprinklePlaces)
{
    sprinkle++;
    Write($"    <FreePoint Name=\"Sprinkle{sprinkle}\" Style=\"{style}\" X=\"{Format(x)}\" Y=\"{Format(y)}\" />");
}

Write("    <PointByCoordinates Name=\"O\" Visible=\"false\" X=\"0\" Y=\"0\" />");
// the guides: segments from the center to just past the tip, which the tip and the junction
// slide along (rays ran across the whole page)
const double GuideLength = 4.3;
Write($"    <PointByCoordinates Name=\"AxisEnd\" Visible=\"false\" X=\"{Format(GuideLength)}\" Y=\"0\" />");
Write("    <Segment Name=\"Axis\" Style=\"Guide\">");
Write("      <Dependency Name=\"O\" />");
Write("      <Dependency Name=\"AxisEnd\" />");
Write("    </Segment>");
Write("    <LineTwoPoints Name=\"Mirror\" Visible=\"false\">");
Write("      <Dependency Name=\"O\" />");
Write("      <Dependency Name=\"AxisEnd\" />");
Write("    </LineTwoPoints>");
Write($"    <PointByCoordinates Name=\"SideEnd\" Visible=\"false\" X=\"{Format(GuideLength * Math.Cos(Math.PI / 6))}\" Y=\"{Format(GuideLength * Math.Sin(Math.PI / 6))}\" />");
Write("    <Segment Name=\"SideRay\" Style=\"Guide\">");
Write("      <Dependency Name=\"O\" />");
Write("      <Dependency Name=\"SideEnd\" />");
Write("    </Segment>");
Write($"    <PointOnFigure Name=\"T\" Style=\"Node\" X=\"3.2\" Y=\"0\" Parameter=\"{Format(3.2 / GuideLength)}\">");
Write("      <Dependency Name=\"Axis\" />");
Write("    </PointOnFigure>");
foreach (var (name, (x, y)) in free)
{
    Write($"    <FreePoint Name=\"{name}\" Style=\"Node\" X=\"{Format(x)}\" Y=\"{Format(y)}\" />");
}

Write($"    <PointOnFigure Name=\"J\" Style=\"Node\" X=\"{Format(0.9 * Math.Cos(Math.PI / 6))}\" Y=\"{Format(0.9 * Math.Sin(Math.PI / 6))}\" Parameter=\"{Format(0.9 / GuideLength)}\">");
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
Write("    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"Every snowflake has six arms, and every arm is the same on both sides. So a snowflake is really one small shape, drawn twelve times: mirrored, then turned five times by 60 degrees.\\n\\nDrag the blue points. You are drawing one twelfth, between the two dashed lines, and the other eleven follow. The tip slides along the arm, the last point along the line where two arms meet.\\n\\nReal snowflakes grow their six arms at once in the same cloud, in the same air, which is why the arms match. No two snowflakes take the same trip down, which is why no two flakes match. Can you make one that looks like a real snowflake? One that looks like a flower? The snowman and his friends can be dragged around too.\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("  </Figures>");
text.Append("</Drawing>");
File.WriteAllText(args[0], text.ToString(), new UTF8Encoding(false));
Console.WriteLine("wrote " + args[0]);
return 0;
