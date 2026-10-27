#:property Nullable=disable
#:property PublishAot=false

// dissections - writes the two triangle-to-square drawings: an equilateral triangle cut into
// pieces that a slider moves into a square of the same area.
//
//   The Haberdasher's Puzzle      - Dudeney's four pieces (1902), hinged at the midpoint D of
//                                   AB, the midpoint E of BC and the point K on AC, swinging
//                                   all at once; the whole settles by 11° so that the square
//                                   comes out level
//   Triangle to Square in Five    - the top of the triangle folds down into the corners (a
//                                   rectangle, three pieces), then two pieces slide along a
//                                   slanted cut (the P-slide), five pieces in all
//
//   dotnet tools/dissections.cs -- <the gallery's Drawings folder> [--probe <folder>]
//
// then LiveGeometry.Desktop.exe --rewrite <that folder>, which gives the files byte for byte.
// --probe writes copies with the slider at 0, 0.5, 1, 1.5 and 2 into the folder, for --check.

using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

if (args.Length < 1)
{
    Console.WriteLine("usage: dissections <folder> [--probe <folder>]");
    return 1;
}

string probe = null;
for (int i = 1; i < args.Length; i++)
{
    if (args[i] == "--probe" && i + 1 < args.Length)
    {
        probe = args[++i];
    }
}

Directory.CreateDirectory(args[0]);
Dissections.Haberdasher(Path.Combine(args[0], "HaberdashersPuzzle.lgf"), value: 0);
Dissections.FivePieces(Path.Combine(args[0], "TriangleToSquare.lgf"), value: 0.8);
if (probe != null)
{
    Directory.CreateDirectory(probe);
    foreach (var value in new[] { 0.0, 0.5, 1.0, 1.5, 2.0 })
    {
        string suffix = value.ToString("0.0", CultureInfo.InvariantCulture);
        Dissections.Haberdasher(Path.Combine(probe, $"HaberdashersPuzzle-{suffix}.lgf"), value);
        Dissections.FivePieces(Path.Combine(probe, $"TriangleToSquare-{suffix}.lgf"), value);
    }
}

return 0;

/// <summary>A point of the plane, for working the motion out ahead of the drawing</summary>
record struct P(double X, double Y)
{
    public static P operator +(P a, P b) => new(a.X + b.X, a.Y + b.Y);

    public static P operator -(P a, P b) => new(a.X - b.X, a.Y - b.Y);

    public static P operator *(double k, P a) => new(k * a.X, k * a.Y);

    public double Dot(P other) => X * other.X + Y * other.Y;

    public double Length => Math.Sqrt(X * X + Y * Y);

    /// <summary>Turned about a center by an angle in degrees, counterclockwise</summary>
    public P Rotate(P center, double degrees)
    {
        double radians = degrees * Math.PI / 180;
        double cos = Math.Cos(radians);
        double sin = Math.Sin(radians);
        var d = this - center;
        return new P(center.X + d.X * cos - d.Y * sin, center.Y + d.X * sin + d.Y * cos);
    }
}

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

    public string Point(string name, P at, string style = null, bool visible = false)
    {
        return Point(name, F(at.X), F(at.Y), style, visible);
    }

    /// <summary>A point by coordinates moved from another by a vector times a number</summary>
    public string Shifted(string name, string source, P by, string times, string style = null, bool visible = false)
    {
        return Point(name, $"{source}.X + {F(by.X)} * {times}", $"{source}.Y + {F(by.Y)} * {times}", style, visible);
    }

    public string Rot(string name, string source, string center, string angle, string style = null, bool visible = false)
    {
        Deps(source, center, angle);
        Element("RotatedPoint", name, visible, style, "", [source, center, angle]);
        return name;
    }

    /// <summary>A hidden Number worked out by a formula</summary>
    public string Num(string name, string formula)
    {
        var deps = Deps(formula);
        Element("Number", name, true, null, $" Formula=\"{Escape(formula)}\"", deps);
        return name;
    }

    public string Slider(string name, P at, double value, double maximum)
    {
        Element("Slider", name, true, null, $" X=\"{F(at.X)}\" Y=\"{F(at.Y)}\" Value=\"{F(value)}\" Maximum=\"{F(maximum)}\"", []);
        return name;
    }

    public string Polygon(string name, string style, params string[] vertices)
    {
        Deps(vertices);
        Element("Polygon", name, true, style, "", vertices);
        return name;
    }

    public void Caption(string title, string text)
    {
        Raw($"<Label Name=\"Title\" Style=\"Heading\" Text=\"{Escape(title)}\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
        Raw($"<Label Name=\"Description\" Style=\"Explanation\" Text=\"{Escape(text)}\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\" />");
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

static class Dissections
{
    static string F(double value) => Doc.F(value);

    static readonly double Root3 = Math.Sqrt(3);

    /// <summary>The side of the square with the area of the triangle of side 2</summary>
    static readonly double Side = Math.Pow(3, 0.25);

    /// <summary>The hinge dots: a dark dot with a light rim, the other way round on the dark paper</summary>
    const string HingeStyle = """
        <PointStyle Name="Hinge" Size="8" Fill="#FF4A4F58" Color="#FFFFFFFF" StrokeWidth="1.5">
          <Dark Fill="#FFE0E4EC" Color="#FF202428" />
        </PointStyle>
        """;

    /// <summary>
    /// The viewport and the scene: the box every piece sweeps through, padded, and the slider
    /// at its left, as low as the pieces leave its track and caption room at every step, else
    /// under them. Returns where the slider goes.
    /// </summary>
    static P Frame(Doc d, Func<double, List<P[]>> pieces, double maximum)
    {
        // every edge of every piece at every step of the slider, as points along it
        var sweep = new List<P>();
        for (int step = 0; step <= 200; step++)
        {
            foreach (var piece in pieces(maximum * step / 200))
            {
                for (int i = 0; i < piece.Length; i++)
                {
                    var from = piece[i];
                    var to = piece[(i + 1) % piece.Length];
                    for (int k = 0; k <= 20; k++)
                    {
                        sweep.Add(from + (k / 20.0) * (to - from));
                    }
                }
            }
        }

        double minX = sweep.Min(p => p.X);
        double maxX = sweep.Max(p => p.X);
        double minY = sweep.Min(p => p.Y);
        double maxY = sweep.Max(p => p.Y);
        const double CaptionHeight = 0.35;
        const double Below = 0.1;
        var slider = new P(minX, minY - 0.45);
        for (double y = minY; y <= maxY - CaptionHeight; y += 0.05)
        {
            bool free = !sweep.Any(p => p.X >= minX - 0.05 && p.X <= minX + maximum + 0.1 && p.Y >= y - Below && p.Y <= y + CaptionHeight);
            if (free)
            {
                slider = new P(minX, y);
                break;
            }
        }

        double left = minX - 0.2;
        double right = Math.Max(maxX, slider.X + maximum) + 0.2;
        double bottom = Math.Min(minY, slider.Y - Below) - 0.2;
        double top = maxY + 0.2;
        d.Header($"<Viewport Left=\"{F(left)}\" Top=\"{F(top)}\" Right=\"{F(right)}\" Bottom=\"{F(bottom)}\" />");
        d.Header($"<Scene Left=\"{F(left)}\" Top=\"{F(top)}\" Right=\"{F(right)}\" Bottom=\"{F(bottom)}\" />");
        return slider;
    }

    /// <summary>Checks that the pieces at the end lie in the square, and that its four corners are reached</summary>
    static void CheckSquare(string name, List<P[]> pieces, double left, double bottom)
    {
        var corners = new[] { new P(left, bottom), new P(left + Side, bottom), new P(left + Side, bottom + Side), new P(left, bottom + Side) };
        var reached = new bool[4];
        foreach (var vertex in pieces.SelectMany(piece => piece))
        {
            if (vertex.X < left - 1e-9 || vertex.X > left + Side + 1e-9 || vertex.Y < bottom - 1e-9 || vertex.Y > bottom + Side + 1e-9)
            {
                throw new InvalidOperationException($"{name}: a piece ends outside the square at ({F(vertex.X)}, {F(vertex.Y)})");
            }

            for (int i = 0; i < 4; i++)
            {
                if ((vertex - corners[i]).Length < 1e-9)
                {
                    reached[i] = true;
                }
            }
        }

        if (reached.Any(r => !r))
        {
            throw new InvalidOperationException($"{name}: a corner of the square is not reached");
        }

        Console.WriteLine($"{name}: the pieces end in the square [{F(left)}, {F(left + Side)}] x [{F(bottom)}, {F(bottom + Side)}]");
    }

    /// <summary>
    /// Dudeney's four pieces. Triangle ABC of side 2, apex A up, D the midpoint of AB, E the
    /// midpoint of BC; J on AC with EJ the side of the square, K on JC with JK = 1; L and M the
    /// feet of the perpendiculars from D and K to EJ. The pieces hinge at D, E and K and each
    /// swings a half turn through the outside of the triangle; all three hinges turn together
    /// with the slider, and the whole turns about E by what makes the square level.
    /// </summary>
    public static void Haberdasher(string path, double value)
    {
        P A = new(0, Root3), B = new(-1, 0), C = new(1, 0);
        P D = 0.5 * (A + B);
        P E = 0.5 * (B + C);
        double along = (1 + Math.Sqrt(4 * Side * Side - 3)) / 4;
        P J = C + along * (A - C);
        P K = J + 0.5 * (C - A);
        P u = (1 / (J - E).Length) * (J - E);
        P L = E + (D - E).Dot(u) * u;
        P M = E + (K - E).Dot(u) * u;
        double settle = 90 - Math.Atan2(u.Y, u.X) * 180 / Math.PI;
        const double Maximum = 2;

        List<P[]> Pieces(double swing)
        {
            double g = settle * swing / Maximum;
            double aboutD = 180 * swing / Maximum;
            double aboutE = g - 180 * swing / Maximum;
            double aboutK = -180 * swing / Maximum;
            P Settle(P p) => p.Rotate(E, g);
            return
            [
                [Settle(D), Settle(L), E, Settle(B)],
                [Settle(A.Rotate(D, aboutD)), Settle(J.Rotate(D, aboutD)), Settle(L.Rotate(D, aboutD)), Settle(D)],
                [M.Rotate(E, aboutE), K.Rotate(E, aboutE), C.Rotate(E, aboutE), E],
                [J.Rotate(K, aboutK).Rotate(E, aboutE), K.Rotate(E, aboutE), M.Rotate(K, aboutK).Rotate(E, aboutE)],
            ];
        }

        CheckSquare("Haberdasher", Pieces(Maximum), left: -Side, bottom: -(K - E).Dot(u));
        var d = new Doc();
        var slider = Frame(d, Pieces, Maximum);
        d.Style(HingeStyle);
        d.Slider("swing", slider, value, Maximum);
        d.Num("settle", $"{F(settle / Maximum)} * swing");
        d.Num("aboutD", $"{F(180 / Maximum)} * swing");
        d.Num("aboutE", $"{F(settle / Maximum - 180 / Maximum)} * swing");
        d.Num("aboutK", $"{F(-180 / Maximum)} * swing");
        d.Point("A", A);
        d.Point("B", B);
        d.Point("C", C);
        d.Point("D", D);
        d.Point("E", E, style: "Hinge", visible: true);
        d.Point("J", J);
        d.Point("K", K);
        d.Point("L", L);
        d.Point("M", M);

        // the piece with B stays, but for the settling
        d.Rot("D1", "D", "E", "settle", style: "Hinge", visible: true);
        d.Rot("L1", "L", "E", "settle");
        d.Rot("B1", "B", "E", "settle");

        // the piece with A swings about D
        d.Rot("A1", "A", "D", "aboutD");
        d.Rot("J1", "J", "D", "aboutD");
        d.Rot("L2", "L", "D", "aboutD");
        d.Rot("A2", "A1", "E", "settle");
        d.Rot("J2", "J1", "E", "settle");
        d.Rot("L3", "L2", "E", "settle");

        // the piece with C swings about E, the small triangle about K with it
        d.Rot("M1", "M", "E", "aboutE");
        d.Rot("K1", "K", "E", "aboutE", style: "Hinge", visible: true);
        d.Rot("C1", "C", "E", "aboutE");
        d.Rot("J3", "J", "K", "aboutK");
        d.Rot("M2", "M", "K", "aboutK");
        d.Rot("J4", "J3", "E", "aboutE");
        d.Rot("M3", "M2", "E", "aboutE");

        d.Polygon("PieceB", "GradientBlueOutline", "D1", "L1", "E", "B1");
        d.Polygon("PieceA", "GradientRedOutline", "A2", "J2", "L3", "D1");
        d.Polygon("PieceC", "GradientGreenOutline", "M1", "K1", "C1", "E");
        d.Polygon("PieceK", "GradientOrangeOutline", "J4", "K1", "M3");
        d.Caption(
            "The Haberdasher's Puzzle",
            "Cut an equilateral triangle into four pieces that fit together into a square. Drag the slider: the pieces swing on three hinges, like links of a bracelet, and the triangle becomes the square. Swing them back and it is a triangle again.\n\n" +
            "Henry Dudeney, an English puzzle maker, set this problem in 1902 and asked for the fewest pieces. His own answer took five (see Triangle to Square in Five Pieces). A reader, C. W. McElroy, sent in this one with four, and Dudeney liked it so much that he had a model made of mahogany with brass hinges. For a hundred years nobody knew whether three pieces could do it. In 2024 it was proved that they can't.\n\n" +
            "The cuts hide a surprise. In the square, the three corners of the triangle meet at one point on a cut, and it looks like the middle of that cut. It isn't quite: the two parts are 51 and 49 hundredths of it. A model made by guessing the middle never closes.");
        d.Write(path);
    }

    /// <summary>
    /// Five pieces in two moves. The top of the triangle, above the midline DE, is cut by its
    /// altitude into two halves that fold down about D and E into the corners: a rectangle 2
    /// wide and √3/2 tall. Then the P-slide: a slant from the bottom left corner B to the
    /// point Q of the top side where the square's width ends, and a cut straight up from the
    /// bottom at the square's width from the right; the two pieces left of the slant slide
    /// along it (the left half of the top with them), the small triangle at the bottom the
    /// whole length of the slant, and the rectangle is the square.
    /// </summary>
    public static void FivePieces(string path, double value)
    {
        double height = Root3 / 2;
        P A = new(0, Root3), B = new(-1, 0), C = new(1, 0);
        P D = new(-0.5, height), E = new(0.5, height), G = new(0, height);
        P Q = new(Side - 1, height);
        P V = new(1 - Side, 0);
        P R = new(1 - Side, (2 - Side) * height / Side);
        P slideOfTop = new(2 - Side, Side - height);
        P slideOfSmall = Q - B;
        const double Maximum = 2;

        // the slider's first half folds, the third quarter slides the big piece, the last the small one
        List<P[]> Pieces(double move)
        {
            double fold = 180 * Math.Clamp(move, 0, 1);
            double slide = Math.Clamp(2 * (move - 1), 0, 1);
            double slideSmall = Math.Clamp(2 * (move - 1.5), 0, 1);
            P Top(P p) => p + slide * slideOfTop;
            P Small(P p) => p + slideSmall * slideOfSmall;
            return
            [
                [Top(A.Rotate(D, fold)), Top(D), Top(G.Rotate(D, fold))],
                [A.Rotate(E, -fold), G.Rotate(E, -fold), E],
                [Top(B), Top(Q), Top(D)],
                [V, C, E, Q, R],
                [Small(B), Small(V), Small(R)],
            ];
        }

        CheckSquare("Five pieces", Pieces(Maximum), left: 1 - Side, bottom: 0);
        var d = new Doc();
        var slider = Frame(d, Pieces, Maximum);
        d.Style(HingeStyle);
        d.Slider("move", slider, value, Maximum);
        d.Num("fold", "180 * clamp(move, 0, 1)");
        d.Num("foldBack", "-180 * clamp(move, 0, 1)");
        d.Num("slide", "clamp(2 * (move - 1), 0, 1)");
        d.Num("slideSmall", "clamp(2 * (move - 1.5), 0, 1)");
        d.Point("A", A);
        d.Point("B", B);
        d.Point("C", C);
        d.Point("D", D);
        d.Point("E", E, style: "Hinge", visible: true);
        d.Point("G", G);
        d.Point("Q", Q);
        d.Point("V", V);
        d.Point("R", R);

        // the left half of the top folds about D, then slides with the slanted piece
        d.Rot("A1", "A", "D", "fold");
        d.Rot("G1", "G", "D", "fold");
        d.Shifted("A2", "A1", slideOfTop, "slide");
        d.Shifted("G2", "G1", slideOfTop, "slide");
        d.Shifted("D2", "D", slideOfTop, "slide", style: "Hinge", visible: true);

        // the right half folds about E and stays
        d.Rot("A3", "A", "E", "foldBack");
        d.Rot("G3", "G", "E", "foldBack");

        // the slanted piece and the small triangle slide
        d.Shifted("B2", "B", slideOfTop, "slide");
        d.Shifted("Q2", "Q", slideOfTop, "slide");
        d.Shifted("B3", "B", slideOfSmall, "slideSmall");
        d.Shifted("V3", "V", slideOfSmall, "slideSmall");
        d.Shifted("R3", "R", slideOfSmall, "slideSmall");

        // the small triangle last: it glides over the base into its notch
        d.Polygon("LeftHalf", "GradientRedOutline", "A2", "D2", "G2");
        d.Polygon("RightHalf", "GradientPurpleOutline", "A3", "G3", "E");
        d.Polygon("Slant", "GradientOrangeOutline", "B2", "Q2", "D2");
        d.Polygon("Base", "GradientBlueOutline", "V", "C", "E", "Q", "R");
        d.Polygon("Small", "GradientGreenOutline", "B3", "V3", "R3");
        d.Caption(
            "Triangle to Square in Five Pieces",
            "Another way to turn an equilateral triangle into a square, with five pieces and three moves. Drag the slider. First the top of the triangle splits in two, and the halves swing down on hinges into the corners: the triangle is now a rectangle, twice as wide as it is tall. Then the piece above the slanted cut slides up along it, the red corner riding along, which opens a notch at the top right, and the small triangle from the bottom left glides along the cut into the notch. The rectangle has squared up.\n\n" +
            "The slide is an old trick for turning any rectangle into another of the same area, as long as the new one is at least half as wide. Cut from a corner to the point on the far side where the new width ends, cut a small triangle off the bottom, and slide.\n\n" +
            "Four pieces can do it too, with cleverer cuts: see The Haberdasher's Puzzle. Three can't, as was proved in 2024.");
        d.Write(path);
    }
}
