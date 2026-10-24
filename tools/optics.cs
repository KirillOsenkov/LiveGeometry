#:property Nullable=disable
#:property PublishAot=false

// optics - writes the optics drawings: light bent by lenses, split by a prism, and the
// rainbow, each from rays worked out live in expressions, so that dragging a lens, a lamp
// or the sun sends the light somewhere else.
//
//   The Upside-Down Candle  - a thin lens draws a picture of a candle: three rays, the
//                             lens rule, a real picture or a magnifying glass
//   Light It with the Sun   - a burning glass: sunbeams refracted exactly at both faces of
//                             a thick lens (Snell's law), a candle that lights in the focus
//   Make a Rainbow          - a prism splits white light into seven colors, also through
//                             a reflection inside when the light can't get out
//   Inside a Raindrop       - the path of each color through a drop: in, bounce, out at 42°
//   Chase the Rainbow       - why a rainbow is round, and why you can never reach its end
//
//   dotnet tools/optics.cs -- <the gallery's Drawings folder>
//
// then LiveGeometry.Desktop.exe --rewrite <that folder>, which gives the files byte for byte.

using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

if (args.Length < 1)
{
    Console.WriteLine("usage: optics <folder>");
    return 1;
}

Directory.CreateDirectory(args[0]);
Optics.UpsideDownCandle(Path.Combine(args[0], "UpsideDownCandle.lgf"));
Optics.BurningGlass(Path.Combine(args[0], "BurningGlass.lgf"));
Optics.Prism(Path.Combine(args[0], "Prism.lgf"));
Optics.Raindrop(Path.Combine(args[0], "Raindrop.lgf"));
Optics.ChaseTheRainbow(Path.Combine(args[0], "ChaseTheRainbow.lgf"));
return 0;

/// <summary>A drawing being written: styles and figures, the dependencies of an expression read off its names</summary>
class Doc
{
    readonly List<string> header = [];
    readonly StringBuilder styles = new();
    readonly StringBuilder figures = new();
    readonly HashSet<string> names = [];
    static readonly HashSet<string> Constants = ["pi", "e", "x"];
    static readonly Regex Identifier = new(@"(?<![.\w'`])([A-Za-z_][A-Za-z0-9_']*)(\s*\()?");

    public static string F(double value)
    {
        return value.ToString("0.######", CultureInfo.InvariantCulture);
    }

    public void Header(string line)
    {
        header.Add("  " + line);
    }

    public void Style(string xml)
    {
        foreach (var line in xml.Replace("\r", "").Split('\n'))
        {
            styles.Append("    ").Append(line).Append("\r\n");
        }
    }

    public void Raw(string xml)
    {
        foreach (var line in xml.Replace("\r", "").Split('\n'))
        {
            figures.Append("    ").Append(line).Append("\r\n");
        }
    }

    void Name(string name)
    {
        if (!names.Add(name))
        {
            throw new InvalidOperationException("two figures named " + name);
        }
    }

    public bool Has(string name) => names.Contains(name);

    /// <summary>The figures an expression names, in order: every name must be a figure, a function or a constant</summary>
    public List<string> Deps(params string[] expressions)
    {
        var result = new List<string>();
        foreach (var expression in expressions)
        {
            if (Regex.IsMatch(expression, @"\d[eE][-+]?\d"))
            {
                throw new InvalidOperationException($"'{expression}' has an exponent, which the expression language doesn't read");
            }

            foreach (Match match in Identifier.Matches(expression))
            {
                var token = match.Groups[1].Value;
                if (match.Groups[2].Success || Constants.Contains(token))
                {
                    continue;
                }

                if (!names.Contains(token))
                {
                    throw new InvalidOperationException($"'{token}' in '{expression}' is no figure");
                }

                if (!result.Contains(token))
                {
                    result.Add(token);
                }
            }
        }

        return result;
    }

    void Element(string kind, string name, bool visible, string style, string attributes, IEnumerable<string> dependencies)
    {
        Name(name);
        var sb = new StringBuilder();
        sb.Append($"    <{kind} Name=\"{name}\"");
        if (!visible)
        {
            sb.Append(" Visible=\"false\"");
        }

        if (style != null)
        {
            sb.Append($" Style=\"{style}\"");
        }

        sb.Append(attributes);
        var list = dependencies.ToList();
        if (list.Count == 0)
        {
            sb.Append(" />\r\n");
        }
        else
        {
            sb.Append(">\r\n");
            foreach (var dependency in list)
            {
                sb.Append($"      <Dependency Name=\"{dependency}\" />\r\n");
            }

            sb.Append($"    </{kind}>\r\n");
        }

        figures.Append(sb);
    }

    public string Point(string name, string x, string y, string style = null, bool visible = false)
    {
        var deps = Deps(x, y);
        Element("PointByCoordinates", name, visible, style, $" X=\"{Escape(x)}\" Y=\"{Escape(y)}\"", deps);
        return name;
    }

    public string Num(string name, string formula)
    {
        var deps = Deps(formula);
        Element("Number", name, true, null, $" Formula=\"{Escape(formula)}\"", deps);
        return name;
    }

    public string Free(string name, double x, double y, string style, bool visible = true)
    {
        Element("FreePoint", name, visible, style, $" X=\"{F(x)}\" Y=\"{F(y)}\"", []);
        return name;
    }

    public string OnFigure(string name, string figure, double x, double y, double parameter, string style, bool visible = true)
    {
        Deps(figure);
        Element("PointOnFigure", name, visible, style, $" X=\"{F(x)}\" Y=\"{F(y)}\" Parameter=\"{F(parameter)}\"", [figure]);
        return name;
    }

    public string Fig(string kind, string name, string style, bool visible, params string[] dependencies)
    {
        foreach (var dependency in dependencies)
        {
            Deps(dependency);
        }

        Element(kind, name, visible, style, "", dependencies);
        return name;
    }

    public string Fig(string kind, string name, string style, string attributes, params string[] dependencies)
    {
        Element(kind, name, true, style, attributes, dependencies);
        return name;
    }

    public string Region(string name, string operation, string style, bool visible, params string[] operands)
    {
        Element("Region", name, visible, style, $" Operation=\"{operation}\"", operands);
        return name;
    }

    public void PointLabel(string name, string point, double offsetX, double offsetY, string style = null)
    {
        Element("PointLabel", name, true, style, $" OffsetX=\"{F(offsetX)}\" OffsetY=\"{F(offsetY)}\"", [point]);
    }

    public void Caption(string title, string text, params string[] dependencies)
    {
        Caption(title, text, decimals: -1, dependencies);
    }

    public void Caption(string title, string text, int decimals, params string[] dependencies)
    {
        Raw($"<Label Name=\"Title\" Style=\"Heading\" Text=\"{Escape(title)}\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
        var sb = new StringBuilder();
        var digits = decimals < 0 ? "" : $" DecimalsToShow=\"{decimals}\"";
        sb.Append($"<Label Name=\"Description\" Style=\"Explanation\"{digits} Text=\"{Escape(text)}\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\"");
        if (dependencies.Length == 0)
        {
            sb.Append(" />");
        }
        else
        {
            sb.Append('>');
            foreach (var dependency in dependencies)
            {
                Deps(dependency);
                sb.Append($"\n  <Dependency Name=\"{dependency}\" />");
            }

            sb.Append("\n</Label>");
        }

        Raw(sb.ToString());
    }

    public static string Escape(string text)
    {
        return text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("\n", "\\n");
    }

    public void Write(string path)
    {
        var text = new StringBuilder();
        text.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n");
        text.Append("<Drawing Version=\"1\">\r\n");
        foreach (var line in header)
        {
            text.Append(line).Append("\r\n");
        }

        text.Append("  <Styles>\r\n").Append(styles).Append("  </Styles>\r\n");
        text.Append("  <Figures>\r\n").Append(figures).Append("  </Figures>\r\n");
        text.Append("</Drawing>");
        File.WriteAllText(path, text.ToString(), new UTF8Encoding(false));
        Console.WriteLine($"wrote {path} ({names.Count} figures)");
    }
}

static class Optics
{
    static string F(double value) => Doc.F(value);

    /// <summary>The seven colors of the rainbow, red first, as they show on a dark paper</summary>
    static readonly (string Name, string Rgb)[] Colors =
    [
        ("Red", "FF3B3B"),
        ("Orange", "FF9A1F"),
        ("Yellow", "FFE83A"),
        ("Green", "3BE06A"),
        ("Blue", "2FA8FF"),
        ("Indigo", "6A6CFF"),
        ("Violet", "C05CFF"),
    ];

    /// <summary>The caption and point label styles on a night paper, the same under every theme</summary>
    static void NightText(Doc d)
    {
        d.Style("""
            <TextStyle Name="Text" FontSize="16" Color="#FFE8E0FF" FontFamily="Segoe UI">
              <Dark Color="#FFE8E0FF" />
            </TextStyle>
            <TextStyle Name="Heading" FontSize="30" Color="#FFFFD966" FontFamily="Segoe UI" Bold="true">
              <Dark Color="#FFFFD966" />
            </TextStyle>
            <TextStyle Name="Explanation" FontSize="16" Color="#FFE8E0FF" FontFamily="Segoe UI">
              <Dark Color="#FFE8E0FF" />
            </TextStyle>
            """);
    }

    /// <summary>A plain dark paper: plain, so that the caption gets a plate that hides the beams going under it</summary>
    static void NightPaper(Doc d, string viewport, string color = "#FF10142E")
    {
        d.Header($"<Viewport {viewport} Color=\"{color}\" />");
    }

    /// <summary>A beam of one color: a wide faint glow under a bright core, and a dashed one for lines of sight</summary>
    static void BeamStyles(Doc d, string name, string rgb, double core = 2.4, double glow = 11, string glowAlpha = "60", string coreAlpha = "FF")
    {
        d.Style($"<LineStyle Name=\"{name}Glow\" Color=\"#{glowAlpha}{rgb}\" StrokeWidth=\"{F(glow)}\" />");
        d.Style($"<LineStyle Name=\"{name}\" Color=\"#{coreAlpha}{rgb}\" StrokeWidth=\"{F(core)}\" />");
        d.Style($"<LineStyle Name=\"{name}Dash\" Color=\"#B0{rgb}\" StrokeWidth=\"1.5\" Dash=\"Dash\" />");
    }

    /// <summary>A glowing beam between two points: a segment, or a ray when it goes on without end</summary>
    static void Beam(Doc d, string kind, string name, string style, string from, string to, bool glow = true)
    {
        if (glow)
        {
            d.Fig(kind, name + "Glow", style + "Glow", true, from, to);
        }

        d.Fig(kind, name, style, true, from, to);
    }

    static void Knobs(Doc d)
    {
        d.Style("""
            <PointStyle Name="Knob" Size="13" Fill="#FFFFD966" Color="#FFFFFFFF" StrokeWidth="2" />
            <PointStyle Name="LensKnob" Shape="Diamond" Size="14" Fill="#FF7FE8FF" Color="#FFFFFFFF" StrokeWidth="2" />
            """);
    }

    static void GlassStyle(Doc d, string name = "Glass")
    {
        d.Style($"""
            <ShapeStyle Name="{name}" Color="#D0BDF2FF" StrokeWidth="2">
              <Fill>
                <LinearGradientBrush>
                  <GradientStop Color="#70D8F6FF" />
                  <GradientStop Color="#3060B8F0" Offset="0.6" />
                  <GradientStop Color="#5090E0FF" Offset="1" />
                </LinearGradientBrush>
              </Fill>
            </ShapeStyle>
            """);
    }

    // ---------------------------------------------------------------------------------
    // The Upside-Down Candle: a thin lens. Every point of the candle goes to its image,
    // L + (P - L) f / (P.X - L.X + f), a real one past the lens or a virtual one behind it.

    static readonly (double X, double Y)[] Flame =
    [
        (0, 2.2), (0.07, 2.05), (0.14, 1.88), (0.195, 1.72), (0.2, 1.6), (0.16, 1.5), (0.08, 1.44), (0, 1.42),
        (-0.08, 1.44), (-0.16, 1.5), (-0.2, 1.6), (-0.195, 1.72), (-0.14, 1.88), (-0.07, 2.05),
    ];

    static readonly (double X, double Y)[] Wax = [(-0.3, 0), (0.3, 0), (0.3, 1.25), (0.22, 1.29), (-0.3, 1.25)];

    static (double X, double Y)[] FlameCore => Flame.Select(p => (p.X * 0.5, 1.47 + (p.Y - 1.42) * 0.5)).ToArray();

    public static void UpsideDownCandle(string path)
    {
        const double H = 2.6; // half the height of the lens
        var d = new Doc();
        NightPaper(d, "Left=\"-13\" Top=\"4.5\" Right=\"13\" Bottom=\"-4.5\"");
        d.Header("<Scene Left=\"-12.5\" Top=\"4\" Right=\"12.5\" Bottom=\"-3.6\" />");
        NightText(d);
        Knobs(d);
        GlassStyle(d);
        d.Style("""
            <LineStyle Name="Axis" Color="#50B8C8FF" StrokeWidth="1" Dash="Dash" />
            <PointStyle Name="Focus" Size="9" Fill="#FFFF8FD8" Color="#FFFFFFFF" StrokeWidth="1.5" />
            <ShapeStyle Name="Wax" Color="#00000000">
              <Fill>
                <LinearGradientBrush EndPoint="1,0">
                  <GradientStop Color="#FFF7E7C4" />
                  <GradientStop Color="#FFFFF8EC" Offset="0.45" />
                  <GradientStop Color="#FFD8BE8E" Offset="1" />
                </LinearGradientBrush>
              </Fill>
            </ShapeStyle>
            <ShapeStyle Name="Flame" Color="#00000000">
              <Fill>
                <LinearGradientBrush EndPoint="0,1">
                  <GradientStop Color="#FFFF5A1F" />
                  <GradientStop Color="#FFFFB13B" Offset="0.6" />
                  <GradientStop Color="#FFFFE27A" Offset="1" />
                </LinearGradientBrush>
              </Fill>
            </ShapeStyle>
            <ShapeStyle Name="FlameCore" Fill="#FFFFFBE0" Color="#00000000" />
            <LineStyle Name="Wick" Color="#FF3A2A1A" StrokeWidth="2.5" />
            <ShapeStyle Name="WaxPicture" Color="#00000000">
              <Fill>
                <LinearGradientBrush EndPoint="1,0">
                  <GradientStop Color="#D0F7E7C4" />
                  <GradientStop Color="#D0FFF8EC" Offset="0.45" />
                  <GradientStop Color="#D0D8BE8E" Offset="1" />
                </LinearGradientBrush>
              </Fill>
            </ShapeStyle>
            <ShapeStyle Name="FlamePicture" Color="#00000000">
              <Fill>
                <LinearGradientBrush EndPoint="0,1">
                  <GradientStop Color="#E0FFE27A" />
                  <GradientStop Color="#E0FFB13B" Offset="0.4" />
                  <GradientStop Color="#E0FF5A1F" Offset="1" />
                </LinearGradientBrush>
              </Fill>
            </ShapeStyle>
            <ShapeStyle Name="WaxGhost" Fill="#50F7E7C4" Color="#C0FFF0D0" StrokeWidth="1.5" Dash="Dash" />
            <ShapeStyle Name="FlameGhost" Fill="#70FF9A3B" Color="#C0FFC070" StrokeWidth="1.5" Dash="Dash" />
            <ShapeStyle Name="Screen" Fill="#22FFFFFF" Color="#55FFFFFF" StrokeWidth="1" />
            <ShapeStyle Name="Light" Fill="#26FFB040" Color="#00000000" />
            <PointStyle Name="Eye" Character="👁️" Size="34" Fill="#FFFFFFFF" />
            """);
        BeamStyles(d, "Red", "FF6B6B");
        BeamStyles(d, "Gold", "FFD84D");
        BeamStyles(d, "Sky", "4DD2FF");

        // the lens goes anywhere (a free point, or a drag of the glass would have nothing to
        // move), its axis through its middle; the handle on its right side makes it fatter
        d.Free("L", 0, 0, "LensKnob");
        d.Point("AxisLeft", "L.X - 60", "L.Y");
        d.Point("AxisRight", "L.X + 60", "L.Y");
        d.Fig("LineTwoPoints", "Axis", "Axis", true, "AxisLeft", "AxisRight");
        d.Point("BulgeFrom", "L.X + 0.15", "L.Y");
        d.Point("BulgeTo", "L.X + 1.6", "L.Y");
        d.Fig("Segment", "BulgeRail", null, false, "BulgeFrom", "BulgeTo");
        d.OnFigure("W", "BulgeRail", 0.75, 0, (0.75 - 0.15) / 1.45, "Knob");
        d.Num("w", "W.X - L.X");
        // a lens of glass (n = 1.5) whose sides are circles of radius R has its focus at R
        d.Num("f", $"({F(H * H)} + w^2) / (2 * w)");
        d.Point("LeftTop", "L.X - w", "L.Y");
        d.Point("LeftCenter", "L.X - w + f", "L.Y");
        d.Point("RightCenter", "L.X + w - f", "L.Y");
        d.Fig("Circle", "LeftSide", null, false, "LeftCenter", "LeftTop");
        d.Fig("Circle", "RightSide", null, false, "RightCenter", "W");
        d.Point("F", "L.X - f", "L.Y", "Focus", visible: true);
        d.Point("F'", "L.X + f", "L.Y", "Focus", visible: true);

        // the candle starts on the axis, but may go above or below it
        d.Free("Candle", -11, 0, "Knob");
        d.Num("u", "L.X - Candle.X");
        d.Num("v", "u * f / (u - f)");
        d.Num("m", "f / (f - u)");

        // a picture past the lens where the rays meet (u > f), or one behind it (u < f)
        var real = " + 0 * sqrt(u - f)";
        var behind = " + 0 * sqrt(f - u) + 0 * sqrt(u - 0.3)";
        d.Point("Tip", "Candle.X", "Candle.Y + 2.2");
        d.Point("TipPicture", "L.X + v", "L.Y + (Tip.Y - L.Y) * f / (f - u)");
        d.Point("TipMeet", "TipPicture.X" + real, "TipPicture.Y" + real);
        d.Point("TipBehind", "TipPicture.X" + behind, "TipPicture.Y" + behind);

        // all the light from the tip of the flame that the lens catches, and where it meets again
        d.Point("LensTop", "L.X", $"L.Y + {F(H)}");
        d.Point("LensBottom", "L.X", $"L.Y - {F(H)}");
        d.Fig("Polygon", "LightIn", "Light", true, "Tip", "LensTop", "LensBottom");
        d.Fig("Polygon", "LightOut", "Light", true, "LensTop", "TipMeet", "LensBottom");
        d.Region("Lens", "Intersection", "Glass", true, "LeftSide", "RightSide");
        d.Point("ScreenA", "L.X + v + 0.12" + real, "L.Y - 2.8");
        d.Point("ScreenB", "L.X + v + 0.3" + real, "L.Y - 2.8");
        d.Point("ScreenC", "L.X + v + 0.3" + real, "L.Y + 2.8");
        d.Point("ScreenD", "L.X + v + 0.12" + real, "L.Y + 2.8");
        d.Fig("Polygon", "Screen", "Screen", true, "ScreenA", "ScreenB", "ScreenC", "ScreenD");

        string[] Shape(string prefix, (double X, double Y)[] points, string gate, bool image)
        {
            var result = new List<string>();
            for (int i = 0; i < points.Length; i++)
            {
                var (dx, dy) = points[i];
                var name = $"{prefix}{i + 1}";
                if (image)
                {
                    // the candle taken as flat, all at the distance of its foot: its picture
                    // is the candle scaled by m (mapped point by point, its thickness along
                    // the axis came out stretched, a candle bent out of shape)
                    d.Point(name, $"L.X + v + {F(dx)} * abs(m){gate}", $"L.Y + (Candle.Y + {F(dy)} - L.Y) * m{gate}");
                }
                else
                {
                    d.Point(name, $"Candle.X + {F(dx)}", $"Candle.Y + {F(dy)}");
                }

                result.Add(name);
            }

            return [.. result];
        }

        d.Fig("Polygon", "Wax", "Wax", true, Shape("Wax", Wax, "", image: false));
        d.Fig("Polygon", "Flame", "Flame", true, Shape("Flame", Flame, "", image: false));
        d.Fig("Polygon", "FlameCore", "FlameCore", true, Shape("Core", FlameCore, "", image: false));
        d.Point("WickBottom", "Candle.X", "Candle.Y + 1.25");
        d.Point("WickTop", "Candle.X", "Candle.Y + 1.45");
        d.Fig("Segment", "Wick", "Wick", true, "WickBottom", "WickTop");

        d.Fig("Polygon", "WaxPicture", "WaxPicture", true, Shape("WaxPicture", Wax, real, image: true));
        d.Fig("Polygon", "FlamePicture", "FlamePicture", true, Shape("FlamePicture", Flame, real, image: true));
        d.Fig("Polygon", "CorePicture", "FlameCore", true, Shape("CorePicture", FlameCore, real, image: true));
        d.Fig("Polygon", "WaxBehind", "WaxGhost", true, Shape("WaxBehind", Wax, behind, image: true));
        d.Fig("Polygon", "FlameBehind", "FlameGhost", true, Shape("FlameBehind", Flame, behind, image: true));

        // three rays from the tip of the flame
        d.Point("Level", "L.X", "Tip.Y");
        d.Point("ThroughF", "L.X", "L.Y + (Tip.Y - L.Y) * f / (f - u)");
        d.Point("ThroughFOn", "ThroughF.X + 1", "ThroughF.Y");

        Beam(d, "Segment", "RedIn", "Red", "Tip", "Level");
        Beam(d, "Ray", "RedOut", "Red", "Level", "F'");
        Beam(d, "Ray", "GoldRay", "Gold", "Tip", "L");
        Beam(d, "Segment", "SkyIn", "Sky", "Tip", "ThroughF");
        Beam(d, "Ray", "SkyOut", "Sky", "ThroughF", "ThroughFOn");
        d.Fig("Segment", "RedBack", "RedDash", true, "Level", "TipBehind");
        d.Fig("Segment", "GoldBack", "GoldDash", true, "L", "TipBehind");
        d.Fig("Segment", "SkyBack", "SkyDash", true, "ThroughF", "TipBehind");
        d.Point("Eye", "L.X + 4.5" + behind, "L.Y + 0.9", "Eye", visible: true);

        d.PointLabel("FLabel", "F", -6, 10);
        d.PointLabel("F'Label", "F'", -6, 10);

        d.Caption(
            "The Upside-Down Candle",
            "A lens bends the light from the flame so that the rays meet again on the other side, and there they draw a picture of the candle. Upside down! Hold a sheet of paper there and you would see it glowing.\n\n" +
            "Three rays are easy to follow: the red one comes in level and leaves through the focus F′, the yellow one goes straight through the middle, and the blue one comes in through the focus F and leaves level.\n\n" +
            "Slide the candle toward the lens: the picture grows and runs away. Inside F the rays never meet. Look back through the lens instead and you see a big candle the right way up: a magnifying glass.\n\n" +
            "Pull the handle on the lens to make it fatter. A fatter lens bends light more and its focus comes closer.\n\n" +
            "Candle to lens: [u]. Lens to picture: [v]. Focus: [f].",
            decimals: 1,
            "u", "v", "f");
        d.Write(path);
    }

    // ---------------------------------------------------------------------------------
    // Light It with the Sun: parallel sunbeams through a thick lens, refracted exactly at
    // both of its faces by Snell's law in vector form, d' = η d + (η cos i - √k) n.

    const double BurnAperture = 2.4;
    const double BurnN = 1.5;
    const int BurnRays = 15;
    const double BurnStep = 0.3;
    const double BurnReach = 0.25; // how near the wick a ray must pass to heat it
    const int BurnNeeded = 11;

    public static void BurningGlass(string path)
    {
        var d = new Doc();
        NightPaper(d, "Left=\"-13\" Top=\"5\" Right=\"13\" Bottom=\"-5\"", "#FF161632");
        d.Header("<Scene Left=\"-10.8\" Top=\"4.1\" Right=\"9.6\" Bottom=\"-3.6\" />");
        NightText(d);
        Knobs(d);
        GlassStyle(d);
        d.Style("""
            <PointStyle Name="Sun" Character="☀️" Size="60" Fill="#FFFFFFFF" />
            <PointStyle Name="Fire" Character="🔥" Size="52" Fill="#FFFFFFFF" />
            <ShapeStyle Name="Wax" Color="#00000000">
              <Fill>
                <LinearGradientBrush EndPoint="1,0">
                  <GradientStop Color="#FFE84A5F" />
                  <GradientStop Color="#FFFF8A9A" Offset="0.45" />
                  <GradientStop Color="#FFB02A40" Offset="1" />
                </LinearGradientBrush>
              </Fill>
            </ShapeStyle>
            <LineStyle Name="Wick" Color="#FF2A2020" StrokeWidth="3" />
            <ShapeStyle Name="Saucer" Fill="#FF8A94A6" Color="#FFC8D0DC" StrokeWidth="1.5" />
            """);
        BeamStyles(d, "Sunbeam", "FFE27A", core: 1.6, glow: 8, glowAlpha: "38", coreAlpha: "E8");
        BeamStyles(d, "Warm", "FFB347", core: 1.6, glow: 8, glowAlpha: "40", coreAlpha: "F0");

        d.Free("Sun", -9.8, 0.8, "Sun");
        d.Free("L", 0, 0, "LensKnob");
        d.Point("BulgeFrom", "L.X + 0.2", "L.Y");
        d.Point("BulgeTo", "L.X + 1.7", "L.Y");
        d.Fig("Segment", "BulgeRail", null, false, "BulgeFrom", "BulgeTo");
        const double w0 = 0.45;
        d.OnFigure("W", "BulgeRail", w0, 0, (w0 - 0.2) / 1.5, "Knob");
        d.Num("w", "W.X - L.X");
        d.Num("R", $"({F(BurnAperture * BurnAperture)} + w^2) / (2 * w)");
        d.Point("LeftTop", "L.X - w", "L.Y");
        d.Point("LeftCenter", "L.X - w + R", "L.Y");
        d.Point("RightCenter", "L.X + w - R", "L.Y");
        d.Fig("Circle", "LeftSide", null, false, "LeftCenter", "LeftTop");
        d.Fig("Circle", "RightSide", null, false, "RightCenter", "W");
        d.Region("Lens", "Intersection", "Glass", true, "LeftSide", "RightSide");

        // the direction of the sunlight, toward the middle of the lens
        d.Num("dx", "(L.X - Sun.X) / distance(Sun, L)");
        d.Num("dy", "(L.Y - Sun.Y) / distance(Sun, L)");

        // the candle, on a saucer, wherever it is put
        d.Free("Candle", 7.6, -1.6, "Knob");
        d.Point("SaucerA", "Candle.X - 0.75", "Candle.Y");
        d.Point("SaucerB", "Candle.X + 0.75", "Candle.Y");
        d.Point("SaucerC", "Candle.X + 0.55", "Candle.Y - 0.18");
        d.Point("SaucerD", "Candle.X - 0.55", "Candle.Y - 0.18");
        d.Point("WaxA", "Candle.X - 0.28", "Candle.Y");
        d.Point("WaxB", "Candle.X + 0.28", "Candle.Y");
        d.Point("WaxC", "Candle.X + 0.28", "Candle.Y + 1.1");
        d.Point("WaxD", "Candle.X - 0.28", "Candle.Y + 1.1");
        d.Point("WickBottom", "Candle.X", "Candle.Y + 1.1");
        d.Point("Wick", "Candle.X", "Candle.Y + 1.3");

        var heat = new List<string>();
        for (int k = 0; k < BurnRays; k++)
        {
            var offset = (k - (BurnRays - 1) / 2.0) * BurnStep;
            var p = $"r{k + 1}";
            // where the ray starts, beside the sun
            d.Num($"{p}ox", $"Sun.X - dy * {F(offset)}");
            d.Num($"{p}oy", $"Sun.Y + dx * {F(offset)}");
            // into the left side, a circle of radius R around LeftCenter
            d.Num($"{p}b", $"({p}ox - LeftCenter.X) * dx + ({p}oy - L.Y) * dy");
            d.Num($"{p}disc", $"{p}b^2 - ({p}ox - LeftCenter.X)^2 - ({p}oy - L.Y)^2 + R^2");
            d.Num($"{p}t", $"-{p}b - sqrt(max({p}disc, 0))");
            d.Num($"{p}hx", $"{p}ox + {p}t * dx");
            d.Num($"{p}hy", $"{p}oy + {p}t * dy");
            d.Num($"{p}hit", $"({p}disc > 0) * (abs({p}hy - L.Y) < {F(BurnAperture)})");
            // the normal there, out of the glass, and the ray bent toward it
            d.Num($"{p}nx", $"({p}hx - LeftCenter.X) / R");
            d.Num($"{p}ny", $"({p}hy - L.Y) / R");
            d.Num($"{p}ci", $"-(dx * {p}nx + dy * {p}ny)");
            var eta = F(1 / BurnN);
            d.Num($"{p}sq", $"sqrt(1 - {eta}^2 * (1 - {p}ci^2))");
            d.Num($"{p}ex", $"{eta} * dx + ({eta} * {p}ci - {p}sq) * {p}nx");
            d.Num($"{p}ey", $"{eta} * dy + ({eta} * {p}ci - {p}sq) * {p}ny");
            // out through the right side, a circle of radius R around RightCenter
            d.Num($"{p}b2", $"({p}hx - RightCenter.X) * {p}ex + ({p}hy - L.Y) * {p}ey");
            d.Num($"{p}t2", $"-{p}b2 + sqrt(max({p}b2^2 - ({p}hx - RightCenter.X)^2 - ({p}hy - L.Y)^2 + R^2, 0))");
            d.Num($"{p}gx", $"{p}hx + {p}t2 * {p}ex");
            d.Num($"{p}gy", $"{p}hy + {p}t2 * {p}ey");
            d.Num($"{p}mx", $"({p}gx - RightCenter.X) / R");
            d.Num($"{p}my", $"({p}gy - L.Y) / R");
            d.Num($"{p}co", $"{p}ex * {p}mx + {p}ey * {p}my");
            d.Num($"{p}k", $"1 - {F(BurnN * BurnN)} * (1 - {p}co^2)");
            d.Num($"{p}fx", $"{F(BurnN)} * {p}ex - ({F(BurnN)} * {p}co - sqrt({p}k)) * {p}mx");
            d.Num($"{p}fy", $"{F(BurnN)} * {p}ey - ({F(BurnN)} * {p}co - sqrt({p}k)) * {p}my");
            // a ray that misses the lens goes straight on
            d.Point($"{p}O", $"{p}ox", $"{p}oy");
            d.Point($"{p}In", $"if({p}hit, {p}hx, {p}ox + dx)", $"if({p}hit, {p}hy, {p}oy + dy)");
            d.Point($"{p}Out", $"if({p}hit, {p}gx, {p}ox + dx)", $"if({p}hit, {p}gy, {p}oy + dy)");
            d.Point($"{p}On", $"if({p}hit, {p}gx + {p}fx, {p}ox + 2 * dx)", $"if({p}hit, {p}gy + {p}fy, {p}oy + 2 * dy)");
            Beam(d, "Segment", $"{p}Sun", "Sunbeam", $"{p}O", $"{p}In");
            Beam(d, "Segment", $"{p}Glass", "Sunbeam", $"{p}In", $"{p}Out");
            Beam(d, "Ray", $"{p}Hot", "Warm", $"{p}Out", $"{p}On");
            // does it pass the wick, going away from the lens?
            d.Num($"{p}heat", $"if({p}hit * ({p}k > 0), (abs((Wick.X - {p}gx) * {p}fy - (Wick.Y - {p}gy) * {p}fx) < {F(BurnReach)}) * ((Wick.X - {p}gx) * {p}fx + (Wick.Y - {p}gy) * {p}fy > 0), 0)");
            heat.Add($"{p}heat");
        }

        d.Num("beams", string.Join(" + ", heat));
        // the candle comes after the beams so that it stands in front of them
        d.Fig("Polygon", "Saucer", "Saucer", true, "SaucerA", "SaucerB", "SaucerC", "SaucerD");
        d.Fig("Polygon", "Wax", "Wax", true, "WaxA", "WaxB", "WaxC", "WaxD");
        d.Fig("Segment", "WickLine", "Wick", true, "WickBottom", "Wick");
        d.Point("Fire", "Wick.X + 0 * sqrt(beams - " + F(BurnNeeded - 0.5) + ")", "Wick.Y + 0.4", "Fire", visible: true);

        d.Caption(
            "Light It with the Sun",
            "Sunlight comes in parallel beams. A lens bends each beam a little, the outer ones more, so that they all squeeze together in one bright, hot spot: the focus. Get enough sunlight onto the wick and the candle lights!\n\n" +
            "Move the candle, or the lens, until the beams meet right at the wick. Drag the sun up or down and watch the spot move. Pull the handle on the lens: a fatter lens bends harder and its focus comes closer.\n\n" +
            "Look closely and the beams don't meet in one point: the outer ones cross a bit nearer the lens. Every lens with round sides does that.\n\n" +
            $"Beams on the wick: [beams]. It takes {BurnNeeded}.\n\n" +
            "The Olympic flame is lit this way at Olympia in Greece, with a curved mirror and the sun.",
            decimals: 0,
            "beams");
        d.Write(path);

        // how many beams the best place for the wick gets, as a check that it can be lit
        int best = 0;
        double bestX = 0, bestY = 0;
        for (double x = 2; x < 9; x += 0.01)
        {
            for (double y = -1.5; y < 1.5; y += 0.01)
            {
                var count = BurnCount(-9.8, 0.8, 0, 0, w0, x, y);
                if (count > best)
                {
                    best = count;
                    bestX = x;
                    bestY = y;
                }
            }
        }

        Console.WriteLine($"  the wick gets at most {best} beams, at ({bestX:0.00}, {bestY:0.00}); at the start it gets {BurnCount(-9.8, 0.8, 0, 0, w0, 7.6, -0.3)}");
        var places = 0;
        for (double x = 2; x < 9; x += 0.01)
        {
            for (double y = -1.5; y < 1.5; y += 0.01)
            {
                if (BurnCount(-9.8, 0.8, 0, 0, w0, x, y) >= BurnNeeded)
                {
                    places++;
                }
            }
        }

        Console.WriteLine($"  it lights within {places / 10000.0:0.000} square units");
    }

    static int BurnCount(double sunX, double sunY, double lx, double ly, double w, double wickX, double wickY)
    {
        double r = (BurnAperture * BurnAperture + w * w) / (2 * w);
        double c1 = lx - w + r, c2 = lx + w - r;
        double length = Math.Sqrt((lx - sunX) * (lx - sunX) + (ly - sunY) * (ly - sunY));
        double dx = (lx - sunX) / length, dy = (ly - sunY) / length;
        int count = 0;
        for (int k = 0; k < BurnRays; k++)
        {
            var offset = (k - (BurnRays - 1) / 2.0) * BurnStep;
            double ox = sunX - dy * offset, oy = sunY + dx * offset;
            double b = (ox - c1) * dx + (oy - ly) * dy;
            double disc = b * b - (ox - c1) * (ox - c1) - (oy - ly) * (oy - ly) + r * r;
            double t = -b - Math.Sqrt(Math.Max(disc, 0));
            double hx = ox + t * dx, hy = oy + t * dy;
            if (!(disc > 0 && Math.Abs(hy - ly) < BurnAperture))
            {
                continue;
            }

            var (ex, ey) = Refract(dx, dy, (hx - c1) / r, (hy - ly) / r, 1 / BurnN);
            double b2 = (hx - c2) * ex + (hy - ly) * ey;
            double t2 = -b2 + Math.Sqrt(Math.Max(b2 * b2 - (hx - c2) * (hx - c2) - (hy - ly) * (hy - ly) + r * r, 0));
            double gx = hx + t2 * ex, gy = hy + t2 * ey;
            var (fx, fy) = Refract(ex, ey, -(gx - c2) / r, -(gy - ly) / r, BurnN);
            if (double.IsNaN(fx))
            {
                continue;
            }

            if (Math.Abs((wickX - gx) * fy - (wickY - gy) * fx) < BurnReach && (wickX - gx) * fx + (wickY - gy) * fy > 0)
            {
                count++;
            }
        }

        return count;
    }

    /// <summary>A unit direction through a surface whose normal faces against it, from index 1 into index 1/eta</summary>
    static (double X, double Y) Refract(double dx, double dy, double nx, double ny, double eta)
    {
        double ci = -(dx * nx + dy * ny);
        double k = 1 - eta * eta * (1 - ci * ci);
        if (k < 0)
        {
            return (double.NaN, double.NaN);
        }

        double s = eta * ci - Math.Sqrt(k);
        return (eta * dx + s * nx, eta * dy + s * ny);
    }

    // ---------------------------------------------------------------------------------
    // Make a Rainbow: a beam of white light into a triangular prism. Each color has its own
    // index of refraction (heavy flint glass); a color that can't get out where it first
    // meets a side bounces off it and gets out through the next one.

    static readonly double[] FlintIndex = [1.7734, 1.7818, 1.7863, 1.7957, 1.8111, 1.8195, 1.8338];
    const double PrismRadius = 2.6;

    public static void Prism(string path)
    {
        var d = new Doc();
        NightPaper(d, "Left=\"-11\" Top=\"5\" Right=\"12\" Bottom=\"-7\"", "#FF0A0B16");
        d.Header("<Scene Left=\"-10.5\" Top=\"3.2\" Right=\"11\" Bottom=\"-6.2\" />");
        NightText(d);
        Knobs(d);
        d.Style("""
            <ShapeStyle Name="Prism" Color="#E0DDF4FF" StrokeWidth="2">
              <Fill>
                <LinearGradientBrush StartPoint="0.5,0" EndPoint="0.5,1">
                  <GradientStop Color="#40E8F8FF" />
                  <GradientStop Color="#2080B0E0" Offset="0.7" />
                  <GradientStop Color="#40B8E0FF" Offset="1" />
                </LinearGradientBrush>
              </Fill>
            </ShapeStyle>
            <ShapeStyle Name="Lamp" Color="#FF9AA4B8" StrokeWidth="1.5">
              <Fill>
                <LinearGradientBrush>
                  <GradientStop Color="#FF5A6478" />
                  <GradientStop Color="#FF2A3040" Offset="1" />
                </LinearGradientBrush>
              </Fill>
            </ShapeStyle>
            <LineStyle Name="LampGlass" Color="#FFFFFFF0" StrokeWidth="5" />
            """);
        BeamStyles(d, "White", "FFFFFF", core: 3, glow: 14, glowAlpha: "40");
        foreach (var (name, rgb) in Colors)
        {
            BeamStyles(d, name, rgb, core: 2.2, glow: 10, glowAlpha: "50");
            d.Style($"<LineStyle Name=\"{name}Inside\" Color=\"#90{rgb}\" StrokeWidth=\"1.5\" />");
        }

        // the prism: an equilateral triangle around M, turned by its top corner K
        d.Free("M", 0, 0, null, visible: false);
        d.Point("PrismReach", $"M.X + {F(PrismRadius)}", "M.Y");
        d.Fig("Circle", "PrismTurn", null, false, "M", "PrismReach");
        d.OnFigure("K", "PrismTurn", 0, PrismRadius, Math.PI / 2, "Knob");
        d.Point("K2", "M.X - 0.5 * (K.X - M.X) - 0.866025403784 * (K.Y - M.Y)", "M.Y + 0.866025403784 * (K.X - M.X) - 0.5 * (K.Y - M.Y)");
        d.Point("K3", "M.X - 0.5 * (K.X - M.X) + 0.866025403784 * (K.Y - M.Y)", "M.Y - 0.866025403784 * (K.X - M.X) - 0.5 * (K.Y - M.Y)");
        d.Fig("Polygon", "Glass", "Prism", true, "K", "K2", "K3");
        var side = PrismRadius * Math.Sqrt(3);
        string[] corners = ["K", "K2", "K3"];
        for (int i = 0; i < 3; i++)
        {
            var a = corners[i];
            var b = corners[(i + 1) % 3];
            // a side as a vector, and its normal out of the glass (the corners go counterclockwise)
            d.Num($"s{i + 1}x", $"{b}.X - {a}.X");
            d.Num($"s{i + 1}y", $"{b}.Y - {a}.Y");
            d.Num($"n{i + 1}x", $"s{i + 1}y / {F(side)}");
            d.Num($"n{i + 1}y", $"-s{i + 1}x / {F(side)}");
        }

        // the lamp: the beam leaves it toward the aiming knob, which turns around it
        const double aim = 1.4;
        var beamAngle = 33.25 * Math.PI / 180;
        var startX = -1.3 * Math.Cos(Math.PI / 6) - 9.2 * Math.Cos(beamAngle);
        var startY = 0.65 - 9.2 * Math.Sin(beamAngle);
        d.Free("Lamp", startX, startY, "Knob");
        d.Point("AimReach", $"Lamp.X + {F(aim)}", "Lamp.Y");
        d.Fig("Circle", "AimTurn", null, false, "Lamp", "AimReach");
        d.OnFigure("Aim", "AimTurn", startX + aim * Math.Cos(beamAngle), startY + aim * Math.Sin(beamAngle), beamAngle, "Knob");
        d.Num("bx", $"(Aim.X - Lamp.X) / {F(aim)}");
        d.Num("by", $"(Aim.Y - Lamp.Y) / {F(aim)}");
        (double Along, double Across)[] lampShape = [(0, 0.48), (-0.6, 0.48), (-0.85, 0.3), (-2.3, 0.3), (-2.3, -0.3), (-0.85, -0.3), (-0.6, -0.48), (0, -0.48)];
        var lampCorners = new List<string>();
        for (int i = 0; i < lampShape.Length; i++)
        {
            var (along, across) = lampShape[i];
            lampCorners.Add(d.Point($"Lamp{i + 1}", $"Lamp.X + {F(along)} * bx - {F(across)} * by", $"Lamp.Y + {F(along)} * by + {F(across)} * bx"));
        }

        d.Fig("Polygon", "LampBody", "Lamp", true, [.. lampCorners]);
        d.Fig("Segment", "LampFront", "LampGlass", true, "Lamp1", "Lamp8");

        // where a ray from (x, y) along (dx, dy) first meets a side of the prism: the
        // distance, or a million when it meets none; and the normal of that side
        void Meet(string p, string x, string y, string dx, string dy)
        {
            for (int i = 1; i <= 3; i++)
            {
                var corner = corners[i - 1];
                // (no exponents: the expression language reads 1e-12 as 1·e - 12)
                d.Num($"{p}den{i}", $"if(abs({dx} * s{i}y - {dy} * s{i}x) < 0.000000000001, 0.000000000001, {dx} * s{i}y - {dy} * s{i}x)");
                var t = $"(({corner}.X - {x}) * s{i}y - ({corner}.Y - {y}) * s{i}x) / {p}den{i}";
                var s = $"(({corner}.X - {x}) * {dy} - ({corner}.Y - {y}) * {dx}) / {p}den{i}";
                d.Num($"{p}t{i}", $"if(({t} > 0.0000001) * ({s} >= 0) * ({s} <= 1), {t}, 1000000)");
            }

            d.Num($"{p}t", $"min({p}t1, {p}t2, {p}t3)");
            d.Num($"{p}x", $"{x} + {p}t * {dx}");
            d.Num($"{p}y", $"{y} + {p}t * {dy}");
            d.Num($"{p}nx", $"({p}t == {p}t1) * n1x + ({p}t == {p}t2) * n2x + ({p}t == {p}t3) * n3x");
            d.Num($"{p}ny", $"({p}t == {p}t1) * n1y + ({p}t == {p}t2) * n2y + ({p}t == {p}t3) * n3y");
        }

        Meet("w", "Lamp.X", "Lamp.Y", "bx", "by");
        d.Num("lit", "wt < 100000");
        d.Point("Entry", "wx", "wy");
        Beam(d, "Segment", "WhiteBeam", "White", "Lamp", "Entry");
        var gate = " + 0 * sqrt(lit - 0.5)";

        for (int c = 0; c < Colors.Length; c++)
        {
            var (name, _) = Colors[c];
            var n = FlintIndex[c];
            var p = name.ToLowerInvariant();
            // into the glass
            var eta = F(1 / n);
            d.Num($"{p}ci", "-(bx * wnx + by * wny)");
            d.Num($"{p}sq", $"sqrt(1 - {eta}^2 * (1 - {p}ci^2))");
            d.Num($"{p}ix", $"{eta} * bx + ({eta} * {p}ci - {p}sq) * wnx");
            d.Num($"{p}iy", $"{eta} * by + ({eta} * {p}ci - {p}sq) * wny");
            // the side it meets inside: out through it, or back off it
            Meet($"{p}A", "wx", "wy", $"{p}ix", $"{p}iy");
            d.Num($"{p}Aco", $"{p}ix * {p}Anx + {p}iy * {p}Any");
            d.Num($"{p}Ak", $"1 - {F(n * n)} * (1 - {p}Aco^2)");
            d.Num($"{p}bounce", $"{p}Ak < 0");
            d.Num($"{p}rx", $"{p}ix - 2 * {p}Aco * {p}Anx");
            d.Num($"{p}ry", $"{p}iy - 2 * {p}Aco * {p}Any");
            d.Num($"{p}ox", $"{F(n)} * {p}ix - ({F(n)} * {p}Aco - sqrt(max({p}Ak, 0))) * {p}Anx");
            d.Num($"{p}oy", $"{F(n)} * {p}iy - ({F(n)} * {p}Aco - sqrt(max({p}Ak, 0))) * {p}Any");
            // after a bounce, the next side
            Meet($"{p}B", $"{p}Ax", $"{p}Ay", $"{p}rx", $"{p}ry");
            d.Num($"{p}Bco", $"{p}rx * {p}Bnx + {p}ry * {p}Bny");
            d.Num($"{p}Bk", $"1 - {F(n * n)} * (1 - {p}Bco^2)");
            d.Num($"{p}qx", $"{F(n)} * {p}rx - ({F(n)} * {p}Bco - sqrt({p}Bk)) * {p}Bnx");
            d.Num($"{p}qy", $"{F(n)} * {p}ry - ({F(n)} * {p}Bco - sqrt({p}Bk)) * {p}Bny");
            d.Point($"{name}A", $"{p}Ax" + gate, $"{p}Ay" + gate);
            d.Point($"{name}B", $"if({p}bounce, {p}Bx, {p}Ax)" + gate, $"if({p}bounce, {p}By, {p}Ay)" + gate);
            d.Num($"{p}fx", $"if({p}bounce, {p}qx, {p}ox)");
            d.Num($"{p}fy", $"if({p}bounce, {p}qy, {p}oy)");
            d.Point($"{name}On", $"{name}B.X + {p}fx", $"{name}B.Y + {p}fy");
            d.Fig("Segment", $"{name}Inside", $"{name}Inside", true, "Entry", $"{name}A");
            d.Fig("Segment", $"{name}Bounce", $"{name}Inside", true, $"{name}A", $"{name}B");
            Beam(d, "Ray", $"{name}Out", name, $"{name}B", $"{name}On");
        }

        // how far the colors fan out, red to violet
        d.Num("spread", "deg(arccos(min(1, redfx * violetfx + redfy * violetfy)))");

        d.Caption(
            "Make a Rainbow",
            "White light is really all the colors mixed together. Glass bends each color by a different amount, violet the most and red the least, so a prism pulls them apart into a rainbow.\n\n" +
            "Turn the prism by its top corner, or drag it. Aim the lamp with the knob in front of it, or move the lamp. At one angle the colors bend the least, and there they are brightest. Turn too far and the light can't get out where it first meets a side: it bounces off the inside like off a mirror and leaves through the next side, and there the colors bend back together, almost white again.\n\n" +
            "This prism is heavy flint glass, which spreads the colors more than window glass. Red to violet: [spread]°.\n\n" +
            "Isaac Newton did this in 1666 with sunlight through a hole in his shutter, and was the first to say that the colors were in the white light all along.",
            decimals: 1,
            "spread");
        d.Write(path);

        // the colors at the start: none should bounce
        var (sx, sy) = (startX, startY);
        Console.WriteLine($"  the lamp starts at ({sx:0.00}, {sy:0.00})");
    }

    // ---------------------------------------------------------------------------------
    // Inside a Raindrop: a ray at height b (as a share of the radius) meets the drop at the
    // angle i, sin i = b, and bends to t, sin t = b / n. Each chord inside turns by π - 2t
    // around the center, so the points it reaches are at polar angles found by adding.

    static readonly double[] WaterIndex = [1.3311, 1.3324, 1.3332, 1.3349, 1.3373, 1.3390, 1.3435];

    public static void Raindrop(string path)
    {
        const double r = 3.6;
        var d = new Doc();
        NightPaper(d, "Left=\"-13\" Top=\"5.5\" Right=\"9\" Bottom=\"-8\"", "#FF121C3E");
        d.Header("<Scene Left=\"-12.5\" Top=\"4.4\" Right=\"5\" Bottom=\"-7.6\" />");
        NightText(d);
        Knobs(d);
        d.Style($"""
            <ShapeStyle Name="Drop" Color="#C0A8E4FF" StrokeWidth="2">
              <Fill>
                <LinearGradientBrush StartPoint="0.2,0.1" EndPoint="0.8,1">
                  <GradientStop Color="#55C8EEFF" />
                  <GradientStop Color="#2238A0E8" Offset="0.6" />
                  <GradientStop Color="#40306AC8" Offset="1" />
                </LinearGradientBrush>
              </Fill>
            </ShapeStyle>
            <ShapeStyle Name="Shine" Fill="#60FFFFFF" Color="#00000000" />
            <PointStyle Name="Sun" Character="☀️" Size="56" Fill="#FFFFFFFF" />
            <PointStyle Name="Eye" Character="👁️" Size="36" Fill="#FFFFFFFF" />
            <LineStyle Name="Back" Color="#70FFFFFF" StrokeWidth="1.5" Dash="Dash" />
            <LineStyle Name="AngleMark" Color="#FFFF8A8A" StrokeWidth="1.5" />
            <TextStyle Name="AngleText" FontSize="20" Color="#FFFF9A9A" FontFamily="Segoe UI" Bold="true" />
            <LineStyle Name="ManySun" Color="#50FFFFFF" StrokeWidth="1" />
            """);
        BeamStyles(d, "White", "FFFFFF", core: 3, glow: 14, glowAlpha: "40");
        foreach (var (name, rgb) in Colors)
        {
            BeamStyles(d, name, rgb, core: 1.8, glow: 8, glowAlpha: "40");
            d.Style($"<LineStyle Name=\"{name}Inside\" Color=\"#A0{rgb}\" StrokeWidth=\"1.2\" />");
            d.Style($"<LineStyle Name=\"{name}Faint\" Color=\"#70{rgb}\" StrokeWidth=\"1.2\" />");
            d.Style($"<LineStyle Name=\"{name}Many\" Color=\"#80{rgb}\" StrokeWidth=\"1.2\" />");
            d.Style($"<LineStyle Name=\"{name}ManyInside\" Color=\"#28{rgb}\" StrokeWidth=\"1\" />");
        }

        d.Point("C", "0", "0");
        d.Point("Rim", F(r), "0");
        d.Fig("Circle", "Drop", "Drop", true, "C", "Rim");
        d.Point("ShineA", "C.X - 2.1", "C.Y + 2.2");
        d.Point("ShineB", "C.X - 1.75", "C.Y + 2.55");
        d.Point("ShineC", "C.X - 1.2", "C.Y + 2.85");
        d.Fig("Ellipse", "Shine", "Shine", true, "ShineB", "ShineC", "ShineA");

        // the sun slides up and down; its beam comes in level
        d.Point("RailBottom", "-10.5", "0.02");
        d.Point("RailTop", "-10.5", F(r - 0.02));
        d.Fig("Segment", "Rail", null, false, "RailBottom", "RailTop");
        var b0 = 0.86;
        d.OnFigure("Sun", "Rail", -10.5, b0 * r, (b0 * r - 0.02) / (r - 0.04), "Sun");
        d.Num("b", $"(Sun.Y - C.Y) / {F(r)}");
        d.Num("i", "arcsin(b)");
        d.Point("Entry", $"C.X - {F(r)} * cos(i)", $"C.Y + {F(r)} * sin(i)");
        Beam(d, "Segment", "Sunlight", "White", "Sun", "Entry");

        string P(string angle) => $"C.X + {F(r)} * cos({angle})";
        string Q(string angle) => $"C.Y + {F(r)} * sin({angle})";

        var secondary = new List<string>();
        for (int c = 0; c < Colors.Length; c++)
        {
            var (name, _) = Colors[c];
            var p = name.ToLowerInvariant();
            d.Num($"{p}t", $"arcsin(b / {F(WaterIndex[c])})");
            d.Point($"{name}Back", P($"2 * {p}t - i"), Q($"2 * {p}t - i"));
            d.Point($"{name}Exit", P($"4 * {p}t - i - pi"), Q($"4 * {p}t - i - pi"));
            d.Point($"{name}On", $"{name}Exit.X + cos(4 * {p}t - 2 * i - pi)", $"{name}Exit.Y + sin(4 * {p}t - 2 * i - pi)");
            d.Fig("Segment", $"{name}In", $"{name}Inside", true, "Entry", $"{name}Back");
            d.Fig("Segment", $"{name}Bounce", $"{name}Inside", true, $"{name}Back", $"{name}Exit");
            Beam(d, "Ray", $"{name}Out", name, $"{name}Exit", $"{name}On");
            // a second bounce, which makes the second, fainter rainbow
            d.Point($"{name}Exit2", P($"6 * {p}t - i - 2 * pi"), Q($"6 * {p}t - i - 2 * pi"));
            d.Point($"{name}On2", $"{name}Exit2.X + cos(6 * {p}t - 2 * i)", $"{name}Exit2.Y + sin(6 * {p}t - 2 * i)");
            secondary.Add(d.Fig("Segment", $"{name}Bounce2", $"{name}Inside", false, $"{name}Exit", $"{name}Exit2"));
            secondary.Add(d.Fig("Ray", $"{name}Out2", $"{name}Faint", false, $"{name}Exit2", $"{name}On2"));
        }

        // the angle between the red light and the way back to the sun
        d.Point("TowardSun", "RedExit.X - 3", "RedExit.Y");
        d.Point("AlongRed", "RedExit.X + 3 * cos(4 * redt - 2 * i - pi)", "RedExit.Y + 3 * sin(4 * redt - 2 * i - pi)");
        d.Fig("Segment", "BackToSun", "Back", true, "RedExit", "TowardSun");
        d.Fig("AngleArc", "RedAngleMark", "AngleMark", true, "RedExit", "TowardSun", "AlongRed");
        d.Fig("AngleMeasurement", "RedAngle", "AngleText", " DecimalsToShow=\"1\"", "RedExit", "TowardSun", "AlongRed");

        // many rays at once, every color: they pile up at the largest angle
        var many = new List<string>();
        for (int k = 1; k <= 12; k++)
        {
            var bk = k / 13.0;
            var ik = F(Math.Asin(bk));
            var name = $"Many{k}";
            d.Point($"{name}In", $"C.X - {F(r)} * cos({ik})", $"C.Y + {F(r)} * sin({ik})");
            d.Point($"{name}Start", "-10.5", $"C.Y + {F(r)} * sin({ik})");
            many.Add(d.Fig("Segment", $"{name}Sun", "ManySun", false, $"{name}Start", $"{name}In"));
            for (int c = 0; c < Colors.Length; c++)
            {
                var color = Colors[c].Name;
                var tk = F(Math.Asin(bk / WaterIndex[c]));
                var ray = $"{name}{color}";
                d.Point($"{ray}Back", P($"2 * {tk} - {ik}"), Q($"2 * {tk} - {ik}"));
                d.Point($"{ray}Exit", P($"4 * {tk} - {ik} - pi"), Q($"4 * {tk} - {ik} - pi"));
                d.Point($"{ray}On", $"{ray}Exit.X + cos(4 * {tk} - 2 * {ik} - pi)", $"{ray}Exit.Y + sin(4 * {tk} - 2 * {ik} - pi)");
                many.Add(d.Fig("Segment", $"{ray}Path", $"{color}ManyInside", false, $"{name}In", $"{ray}Back"));
                many.Add(d.Fig("Segment", $"{ray}Path2", $"{color}ManyInside", false, $"{ray}Back", $"{ray}Exit"));
                many.Add(d.Fig("Ray", $"{ray}Out", $"{color}Many", false, $"{ray}Exit", $"{ray}On"));
            }
        }

        // at the left edge under the sun, out of the way of the light coming out of the drop,
        // and well apart: the boxes keep their size in pixels, and on a phone a unit is a few
        // (beside the drop their text ran off a phone's screen)
        d.Raw(ShowHide("ShowMany", "Show many rays", -12.2, -1.3, many));
        d.Raw(ShowHide("ShowSecond", "Show the second bounce", -12.2, -2.5, secondary));

        d.Caption(
            "Inside a Raindrop",
            "Sunlight goes into a drop of rain, bounces off the back of it and comes out again, heading back toward the sun. On the way in and on the way out it bends, and every color bends by its own amount, so the colors come out a little apart.\n\n" +
            "Slide the sun up and down to send its light into another part of the drop. However you aim, the red light never comes out at more than 42° from the way back to the sun, and violet never more than 40.5°. Show many rays at once: right at that largest angle they pile up and shine brightest. That is where you see the rainbow.\n\n" +
            "Some light bounces twice before it gets out. It makes a second, fainter rainbow outside the first, with its colors the other way round.");
        d.Write(path);
    }

    static string ShowHide(string name, string text, double x, double y, List<string> figures)
    {
        var sb = new StringBuilder();
        sb.Append($"<ShowHideControl Name=\"{name}\" Style=\"Explanation\" Show=\"false\" Text=\"{Doc.Escape(text)}\" X=\"{F(x)}\" Y=\"{F(y)}\">");
        foreach (var figure in figures)
        {
            sb.Append($"\n  <Dependency Name=\"{figure}\" />");
        }

        sb.Append("\n</ShowHideControl>");
        return sb.ToString();
    }

    // ---------------------------------------------------------------------------------
    // Chase the Rainbow: the view ahead, in degrees (a unit is 10°). The rainbow is a circle
    // of 42° around the point straight opposite the sun, the shadow of your head, which is
    // as far below the horizon as the sun is above it behind you.

    public static void ChaseTheRainbow(string path)
    {
        var d = new Doc();
        d.Header("<Viewport Left=\"-9.5\" Top=\"6.5\" Right=\"9.5\" Bottom=\"-4.6\">");
        d.Header("  <Background>");
        d.Header("    <LinearGradientBrush EndPoint=\"0,1\">");
        d.Header("      <GradientStop Color=\"#FF3A4560\" />");
        d.Header("      <GradientStop Color=\"#FF5C6C8C\" Offset=\"0.45\" />");
        d.Header("      <GradientStop Color=\"#FF9CB2CC\" Offset=\"0.62\" />");
        d.Header("      <GradientStop Color=\"#FF4E8A3A\" Offset=\"0.63\" />");
        d.Header("    </LinearGradientBrush>");
        d.Header("  </Background>");
        d.Header("</Viewport>");
        d.Header("<Scene Left=\"-9.2\" Top=\"6.6\" Right=\"9.2\" Bottom=\"-4.4\" />");
        // (a phone in portrait gets the middle: the rainbow, the dial and you)
        d.Header("<Scene Left=\"-6.6\" Top=\"6.6\" Right=\"6.4\" Bottom=\"-4.2\" />");
        d.Style("""
            <TextStyle Name="Text" FontSize="16" Color="#FFF4F6FF" FontFamily="Segoe UI">
              <Dark Color="#FFF4F6FF" />
            </TextStyle>
            <TextStyle Name="Heading" FontSize="30" Color="#FFFFE27A" FontFamily="Segoe UI" Bold="true">
              <Dark Color="#FFFFE27A" />
            </TextStyle>
            <TextStyle Name="Explanation" FontSize="16" Color="#FFF4F6FF" FontFamily="Segoe UI">
              <Dark Color="#FFF4F6FF" />
            </TextStyle>
            <ShapeStyle Name="Ground" Color="#00000000">
              <Fill>
                <LinearGradientBrush EndPoint="0,1">
                  <GradientStop Color="#FF7CC45A" />
                  <GradientStop Color="#FF4E9A3C" Offset="0.4" />
                  <GradientStop Color="#FF2E6A2A" Offset="1" />
                </LinearGradientBrush>
              </Fill>
            </ShapeStyle>
            <ShapeStyle Name="FarHills" Color="#00000000">
              <Fill>
                <LinearGradientBrush EndPoint="0,1">
                  <GradientStop Color="#FF6A9A8A" />
                  <GradientStop Color="#FF5A8A6A" Offset="1" />
                </LinearGradientBrush>
              </Fill>
            </ShapeStyle>
            <ShapeStyle Name="Glow" Fill="#30FFFFFF" Color="#00000000" />
            <ShapeStyle Name="DarkBand" Fill="#28101828" Color="#00000000" />
            <ShapeStyle Name="Shadow" Fill="#50102010" Color="#00000000" />
            <LineStyle Name="ShadowBody" Color="#50102010" StrokeWidth="12" />
            <PointStyle Name="You" Character="🧍" Size="64" Fill="#FFFFFFFF" />
            <PointStyle Name="Gold" Character="💰" Size="40" Fill="#FFFFFFFF" />
            <PointStyle Name="Cloud" Character="🌧️" Size="44" Fill="#E0FFFFFF" />
            <PointStyle Name="Tree" Character="🌳" Size="34" Fill="#FFFFFFFF" />
            <PointStyle Name="House" Character="🏡" Size="38" Fill="#FFFFFFFF" />
            <PointStyle Name="Sun" Character="☀️" Size="34" Fill="#FFFFFFFF" />
            <PointStyle Name="Small you" Character="🧍" Size="22" Fill="#FFFFFFFF" />
            <ShapeStyle Name="Dial" Fill="#30FFFFFF" Color="#80FFFFFF" StrokeWidth="1" />
            <LineStyle Name="DialTrack" Color="#90FFE27A" StrokeWidth="2" Dash="Dot" />
            <LineStyle Name="SunLine" Color="#C0FFE27A" StrokeWidth="1.5" Dash="Dash" />
            <LineStyle Name="BowLine" Color="#FFFF6A6A" StrokeWidth="2" />
            <LineStyle Name="DialGround" Color="#FF4E9A3C" StrokeWidth="3" />
            <TextStyle Name="DialText" FontSize="15" Color="#FFFFFFFF" FontFamily="Segoe UI" Bold="true" />
            <LineStyle Name="AngleMark" Color="#FFFFFFFF" StrokeWidth="1.5" />
            """);
        // the bands, red outside; the second rainbow is fainter and the other way round
        foreach (var (name, rgb) in Colors)
        {
            d.Style($"<ShapeStyle Name=\"{name}Band\" Fill=\"#C8{rgb}\" Color=\"#00000000\" />");
            d.Style($"<ShapeStyle Name=\"{name}Faint\" Fill=\"#48{rgb}\" Color=\"#00000000\" />");
        }

        // the sun behind you: a dial in the corner, seen from the side
        d.Point("Dial", "-4.6", "4.8");
        d.Point("DialLeft", "Dial.X - 1.9", "Dial.Y");
        d.Point("DialRight", "Dial.X + 1.9", "Dial.Y");
        d.Point("DialHigh", "Dial.X - 1.6 * cos(rad(80))", "Dial.Y + 1.6 * sin(rad(80))");
        d.Point("DialLow", "Dial.X - 1.6", "Dial.Y");
        d.Fig("CircleArc", "SunTrack", "DialTrack", true, "Dial", "DialHigh", "DialLow");
        var sun0 = 22.0;
        var sunAngle = Math.PI - sun0 * Math.PI / 180;
        d.OnFigure("Sun", "SunTrack", -4.6 + 1.6 * Math.Cos(sunAngle), 4.8 + 1.6 * Math.Sin(sunAngle), sunAngle, "Sun");
        d.Num("s", "deg(atan2(Sun.Y - Dial.Y, Dial.X - Sun.X))");
        d.Fig("Segment", "DialHorizon", "DialGround", true, "DialLeft", "DialRight");
        d.Point("DialPerson", "Dial.X", "Dial.Y + 0.25", "Small you", visible: true);
        d.Point("DialShadow", "Dial.X + 1.9 * cos(rad(s))", "Dial.Y - 1.9 * sin(rad(s))");
        d.Point("DialBow", "Dial.X + 1.9 * cos(rad(42 - s))", "Dial.Y + 1.9 * sin(rad(42 - s))");
        d.Fig("Segment", "SunToShadow", "SunLine", true, "Sun", "DialShadow");
        d.Fig("Segment", "ToBow", "BowLine", true, "Dial", "DialBow");
        d.Fig("AngleArc", "DialAngleMark", "AngleMark", true, "Dial", "DialShadow", "DialBow");
        d.Fig("AngleMeasurement", "DialAngle", "DialText", " DecimalsToShow=\"0\"", "Dial", "DialShadow", "DialBow");

        // you, walking along the meadow
        d.Point("WalkLeft", "-7.5", "-3.1");
        d.Point("WalkRight", "7.5", "-3.1");
        d.Fig("Segment", "Walk", null, false, "WalkLeft", "WalkRight");
        d.OnFigure("You", "Walk", -1, -3.1, 6.5 / 15, "You");
        d.Point("Center", "You.X", "-s / 10");
        d.Point("Below", "-40", "0");
        d.Point("Below2", "40", "0");
        d.Point("Below3", "40", "-40");
        d.Point("Below4", "-40", "-40");
        d.Fig("Polygon", "BelowHorizon", null, false, "Below", "Below2", "Below3", "Below4");

        var circles = new Dictionary<double, string>();
        string Ring(double radius)
        {
            if (!circles.TryGetValue(radius, out var name))
            {
                name = $"Ring{circles.Count + 1}";
                d.Point(name + "Edge", $"Center.X + {F(radius)}", "Center.Y");
                d.Fig("Circle", name, null, false, "Center", name + "Edge");
                circles[radius] = name;
            }

            return name;
        }

        const double band = 0.07;
        const double primary = 4.25;  // red, outside
        const double second = 5.1;    // red, inside
        var innermost = primary - band * Colors.Length;
        d.Region("Brighter", "Difference", "Glow", true, Ring(innermost), "BelowHorizon");
        d.Region("Alexander", "Difference", "DarkBand", true, Ring(second), Ring(primary), "BelowHorizon");
        for (int c = 0; c < Colors.Length; c++)
        {
            var name = Colors[c].Name;
            d.Region($"{name}Bow", "Difference", $"{name}Band", true, Ring(primary - band * c), Ring(primary - band * (c + 1)), "BelowHorizon");
            d.Region($"{name}Bow2", "Difference", $"{name}Faint", true, Ring(second + band * (c + 1)), Ring(second + band * c), "BelowHorizon");
        }

        // the land in front of the rainbow
        (double X, double Y)[] hills = [(-40, -0.1), (-9.5, -0.1), (-8, 0.35), (-6.5, 0.15), (-5.2, 0.5), (-3.5, 0.05), (-1, 0), (1.5, 0.1), (3, 0.42), (4.6, 0.12), (6.4, 0.55), (8.2, 0.2), (9.5, 0.3), (40, -0.1), (40, -40), (-40, -40)];
        var hillCorners = new List<string>();
        for (int k = 0; k < hills.Length; k++)
        {
            hillCorners.Add(d.Point($"Hill{k + 1}", F(hills[k].X), F(hills[k].Y)));
        }

        d.Fig("Polygon", "FarHills", "FarHills", true, [.. hillCorners]);
        (double X, double Y)[] meadow = [(-40, -0.6), (-6, -0.45), (-2, -0.75), (2.5, -0.5), (7, -0.8), (40, -0.6), (40, -40), (-40, -40)];
        var meadowCorners = new List<string>();
        for (int k = 0; k < meadow.Length; k++)
        {
            meadowCorners.Add(d.Point($"Meadow{k + 1}", F(meadow[k].X), F(meadow[k].Y)));
        }

        d.Fig("Polygon", "Meadow", "Ground", true, [.. meadowCorners]);

        // your shadow points at the middle of the rainbow
        d.Point("Feet", "You.X", "You.Y - 0.55");
        // (with a high sun it falls behind you, out of the picture)
        d.Point("Head", "Center.X", "Center.Y + 0 * sqrt(Center.Y + 2.6)");
        d.Fig("Segment", "ShadowLine", "ShadowBody", true, "Feet", "Head");
        d.Point("HeadSide", "Head.X + 0.32", "Head.Y");
        d.Point("HeadTop", "Head.X", "Head.Y + 0.14");
        d.Fig("Ellipse", "HeadShadow", "Shadow", true, "Head", "HeadSide", "HeadTop");

        d.Free("Gold", 3.6, -0.55, "Gold");
        d.Free("House", -6.1, 0.35, "House");
        d.Free("Tree1", -4.6, 0.55, "Tree");
        d.Free("Tree2", 6.9, 0.75, "Tree");
        d.Free("Cloud1", 0.4, 5.9, "Cloud");
        d.Free("Cloud2", 3.3, 6.2, "Cloud");
        d.Free("Cloud3", 5.6, 5.6, "Cloud");

        d.Caption(
            "Chase the Rainbow",
            "Every rainbow is a circle, and you are standing on the line through its middle. The middle is the shadow of your head: the point straight opposite the sun, which is behind you. The red light comes to you from raindrops 42° away from that point, all the way round. The ground cuts off the bottom of the circle.\n\n" +
            "Drag the sun in the corner. A low sun gives a tall rainbow; when it climbs past 42° the rainbow sinks out of sight. That is why rainbows come early or late in the day.\n\n" +
            "Now walk toward the pot of gold. The rainbow walks with you, since it is made by whichever drops are at the right angle from your own eyes. Nobody has ever reached its end.\n\n" +
            "Above it, a second bow with its colors reversed, and a darker sky between the two.",
            "s");
        d.Write(path);
    }
}
