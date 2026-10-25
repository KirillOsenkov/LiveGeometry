// Writes the drawings of the Catriona Agg library (CatrionaAgg/drawings/<slug>.lgf), one
// function per puzzle: her diagrams redrawn as constructions, in her colors, so that what the
// puzzle says holds while the figure is dragged. No text but what her diagrams have (the given
// numbers, the question marks); the page around the drawing says the rest.
//
//   dotnet run tools/catrionadrawings.cs -- [slug ...]
//
// A drawing is built in the order its figures are added: hidden points by coordinates whose X
// and Y are expressions over the handles (a handle is a free point, or a point on a hidden
// segment or arc that keeps it in range), then the shapes on them. Each drawing has a scene,
// the view it is fitted to. Then catriona.cs marks the puzzles that have a drawing in
// puzzles.json.

#nullable disable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

string repository = FindRepository();
string output = Path.Combine(repository, "CatrionaAgg", "drawings");
Directory.CreateDirectory(output);

var puzzles = new Dictionary<string, Action<Lgf>>
{
    ["three-hexagons-on-a-circle"] = Puzzles.ThreeHexagonsOnACircle,
    ["a-triangle-inside-a-hexagon"] = Puzzles.ATriangleInsideAHexagon,
    ["a-semi-circle-and-a-quarter-circle"] = Puzzles.ASemiCircleAndAQuarterCircle,
    ["seven-squares"] = Puzzles.SevenSquares,
    ["six-squares-and-a-quarter-circle"] = Puzzles.SixSquaresAndAQuarterCircle,
    ["two-hexagons-in-a-rectangle-ii"] = Puzzles.TwoHexagonsInARectangleII,
    ["three-congruent-rectangles-vi"] = Puzzles.ThreeCongruentRectanglesVI,
    ["four-quarter-circles-iii"] = Puzzles.FourQuarterCirclesIII,
    ["two-triangles-in-a-semi-circle"] = Puzzles.TwoTrianglesInASemiCircle,
    ["four-squares-x"] = Puzzles.FourSquaresX,
    ["four-squares-ix"] = Puzzles.FourSquaresIX,
    ["three-regular-hexagons-and-a-semi-circle"] = Puzzles.ThreeRegularHexagonsAndASemiCircle,
    ["four-quarter-circles-ii"] = Puzzles.FourQuarterCirclesII,
    ["subdivided-triangle-ii"] = Puzzles.SubdividedTriangleII,
    ["four-squares-xii"] = Puzzles.FourSquaresXII,
    ["four-squares-xi"] = Puzzles.FourSquaresXI,
    ["four-squares-and-an-isosceles-triangle"] = Puzzles.FourSquaresAndAnIsoscelesTriangle,
    ["two-triangles-and-a-quarter-circle"] = Puzzles.TwoTrianglesAndAQuarterCircle,
    ["two-squares-and-a-quarter-circle"] = Puzzles.TwoSquaresAndAQuarterCircle,
    ["five-squares-ii"] = Puzzles.FiveSquaresII,
};

var wanted = args.Length > 0 ? args : puzzles.Keys.ToArray();
foreach (var slug in wanted)
{
    if (!puzzles.TryGetValue(slug, out var build))
    {
        Console.WriteLine($"error: no drawing called {slug}");
        return 1;
    }

    var drawing = new Lgf();
    build(drawing);
    string file = Path.Combine(output, slug + ".lgf");
    File.WriteAllText(file, drawing.ToText(), new UTF8Encoding(false));
    Console.WriteLine($"{file}: {drawing.FigureCount} figures");
}

return 0;

static string FindRepository()
{
    for (var folder = new DirectoryInfo(Directory.GetCurrentDirectory()); folder != null; folder = folder.Parent)
    {
        if (File.Exists(Path.Combine(folder.FullName, "AGENTS.md")) && Directory.Exists(Path.Combine(folder.FullName, "Main")))
        {
            return folder.FullName;
        }
    }

    throw new Exception("run this from inside the repository");
}

/// <summary>A point as two expressions</summary>
record V(string X, string Y)
{
    public static V Of(string point) => new(point + ".X", point + ".Y");

    public static V At(double x, double y) => new(Lgf.N(x), Lgf.N(y));

    public static V operator +(V a, V b) => new(Sum(a.X, b.X), Sum(a.Y, b.Y));

    public static V operator -(V a, V b) => new(Sum(a.X, Neg(b.X)), Sum(a.Y, Neg(b.Y)));

    public static V operator *(string k, V a) => new(Times(k, a.X), Times(k, a.Y));

    public static V operator *(double k, V a) => Lgf.N(k) * a;

    /// <summary>Turned counterclockwise by the angle, in degrees</summary>
    public V Turn(double degrees)
    {
        double c = Math.Round(Math.Cos(degrees * Math.PI / 180), 12), s = Math.Round(Math.Sin(degrees * Math.PI / 180), 12);
        return new(Sum(Times(Lgf.N(c), X), Times(Lgf.N(-s), Y)), Sum(Times(Lgf.N(s), X), Times(Lgf.N(c), Y)));
    }

    public static string Sum(string a, string b)
    {
        if (a == "0")
        {
            return b;
        }

        if (b == "0")
        {
            return a;
        }

        return b.StartsWith("-") ? $"({a} - {b.Substring(1)})" : $"({a} + {b})";
    }

    public static string Neg(string a) => a == "0" ? "0" : a.StartsWith("-") ? a.Substring(1) : Regex.IsMatch(a, @"^[\w.]+$") ? "-" + a : $"-1 * {a}";

    public static string Times(string k, string a)
    {
        if (k == "0" || a == "0")
        {
            return "0";
        }

        if (k == "1")
        {
            return a;
        }

        if (k == "-1")
        {
            return Neg(a);
        }

        return k.StartsWith("-") ? Neg($"{k.Substring(1)} * {Wrap(a)}") : $"{Wrap(k)} * {Wrap(a)}";
    }

    static string Wrap(string a) => Regex.IsMatch(a, @"^-?[\w.]+$") || (a.StartsWith("(") && a.EndsWith(")") && Balanced(a.Substring(1, a.Length - 2))) ? a : $"({a})";

    static bool Balanced(string a)
    {
        int depth = 0;
        foreach (char c in a)
        {
            depth += c == '(' ? 1 : c == ')' ? -1 : 0;
            if (depth < 0)
            {
                return false;
            }
        }

        return depth == 0;
    }
}

/// <summary>A drawing being written</summary>
class Lgf
{
    public const string Ink = "#FF1E1E1E";
    public const string DarkInk = "#FFE6E6E6";
    public const string Paper = "#FFFFFFFF";
    public const string DarkPaper = "#FF2B2B2B";

    private readonly XElement styles = new("Styles");
    private readonly XElement figures = new("Figures");
    private readonly HashSet<string> names = new(StringComparer.Ordinal);
    private readonly HashSet<string> styleNames = new(StringComparer.Ordinal);
    private (double Left, double Top, double Right, double Bottom) scene;
    private int counter;

    public int FigureCount => figures.Elements().Count();

    public static string N(double value) => Math.Round(value, 10).ToString("0.##########", CultureInfo.InvariantCulture);

    public void Scene(double left, double top, double right, double bottom) => scene = (left, top, right, bottom);

    // ---------- styles ----------

    /// <summary>A filled shape: a diagonal gradient from a lighter tint of the color to the color, outlined in ink (or the given color)</summary>
    public string Fill(string name, string color, string outline = Ink, double width = 2.5, double tint = 0.3)
    {
        if (styleNames.Add(name))
        {
            var style = new XElement("ShapeStyle", new XAttribute("Name", name), new XAttribute("Color", outline == null ? "#00000000" : outline), new XAttribute("StrokeWidth", N(width)),
                new XElement("Fill", Gradient(Mix(color, "#FFFFFFFF", tint), color)));
            var dark = new XElement("Dark");
            if (outline != null)
            {
                dark.Add(new XAttribute("Color", outline == Ink ? DarkInk : Mix(outline, "#FFFFFFFF", 0.45)));
            }

            if (dark.HasAttributes)
            {
                style.Add(dark);
            }

            styles.Add(style);
        }

        return name;
    }

    /// <summary>The paper's own color: white on the light paper, the dark paper's color on the dark one</summary>
    public string PaperFill(string name = "Paper", string outline = Ink, double width = 2.5)
    {
        if (styleNames.Add(name))
        {
            styles.Add(new XElement("ShapeStyle", new XAttribute("Name", name), new XAttribute("Color", outline ?? "#00000000"), new XAttribute("StrokeWidth", N(width)), new XAttribute("Fill", Paper),
                new XElement("Dark", outline == null ? null : new XAttribute("Color", outline == Ink ? DarkInk : Mix(outline, "#FFFFFFFF", 0.45)), new XAttribute("Fill", DarkPaper))));
        }

        return name;
    }

    /// <summary>An outline without a fill</summary>
    public string Outline(string name, string color = Ink, double width = 2.5)
    {
        if (styleNames.Add(name))
        {
            styles.Add(new XElement("ShapeStyle", new XAttribute("Name", name), new XAttribute("Color", color), new XAttribute("StrokeWidth", N(width)), new XAttribute("IsFilled", "false"),
                new XElement("Dark", new XAttribute("Color", color == Ink ? DarkInk : Mix(color, "#FFFFFFFF", 0.45)))));
        }

        return name;
    }

    public string Line(string name, string color = Ink, double width = 2.5)
    {
        if (styleNames.Add(name))
        {
            styles.Add(new XElement("LineStyle", new XAttribute("Name", name), new XAttribute("Color", color), new XAttribute("StrokeWidth", N(width)),
                new XElement("Dark", new XAttribute("Color", color == Ink ? DarkInk : Mix(color, "#FFFFFFFF", 0.4)))));
        }

        return name;
    }

    /// <summary>A filled angle mark</summary>
    public string Wedge(string name, string color)
    {
        if (styleNames.Add(name))
        {
            styles.Add(new XElement("ShapeStyle", new XAttribute("Name", name), new XAttribute("Color", color), new XAttribute("StrokeWidth", "1"), new XAttribute("Fill", color),
                new XElement("Dark", new XAttribute("Color", Mix(color, "#FFFFFFFF", 0.35)), new XAttribute("Fill", Mix(color, "#FFFFFFFF", 0.35)))));
        }

        return name;
    }

    /// <summary>A text on the paper: lightened on the dark paper</summary>
    public string Text(string name = "Mark", string color = Ink, double size = 26)
    {
        if (styleNames.Add(name))
        {
            styles.Add(new XElement("TextStyle", new XAttribute("Name", name), new XAttribute("FontSize", N(size)), new XAttribute("Color", color), new XAttribute("FontFamily", "Segoe UI"), new XAttribute("Bold", "true"),
                new XElement("Dark", new XAttribute("Color", color == Ink ? DarkInk : Mix(color, "#FFFFFFFF", 0.4)))));
        }

        return name;
    }

    /// <summary>A text on a filled shape, whose colors are the same under both themes: the same under both too</summary>
    public string ShapeText(string name = "ShapeMark", string color = Ink, double size = 26)
    {
        if (styleNames.Add(name))
        {
            styles.Add(new XElement("TextStyle", new XAttribute("Name", name), new XAttribute("FontSize", N(size)), new XAttribute("Color", color), new XAttribute("FontFamily", "Segoe UI"), new XAttribute("Bold", "true"),
                new XElement("Dark", new XAttribute("Color", color))));
        }

        return name;
    }

    /// <summary>What a reader drags: a white dot with an ink rim (dark with a light rim on the dark paper)</summary>
    public string Handle()
    {
        const string name = "Handle";
        if (styleNames.Add(name))
        {
            styles.Add(new XElement("PointStyle", new XAttribute("Name", name), new XAttribute("Size", "15"), new XAttribute("Fill", "#FFFFFFFF"), new XAttribute("Color", "#FF1E1E1E"), new XAttribute("StrokeWidth", "2.5"),
                new XElement("Dark", new XAttribute("Fill", "#FF1E2228"), new XAttribute("Color", "#FFF2F2F2"))));
        }

        return name;
    }

    public string Dot()
    {
        const string name = "Dot";
        if (styleNames.Add(name))
        {
            styles.Add(new XElement("PointStyle", new XAttribute("Name", name), new XAttribute("Size", "8"), new XAttribute("Fill", Ink), new XAttribute("Color", Ink), new XAttribute("StrokeWidth", "1"),
                new XElement("Dark", new XAttribute("Fill", DarkInk), new XAttribute("Color", DarkInk))));
        }

        return name;
    }

    static XElement Gradient(string from, string to) => new("LinearGradientBrush",
        new XElement("GradientStop", new XAttribute("Color", from)),
        new XElement("GradientStop", new XAttribute("Color", to), new XAttribute("Offset", "1")));

    /// <summary>The first color moved toward the second by the fraction, alpha kept</summary>
    public static string Mix(string a, string b, double t)
    {
        uint x = Convert.ToUInt32(a.TrimStart('#'), 16), y = Convert.ToUInt32(b.TrimStart('#'), 16);
        int Channel(uint c, int shift) => (int)((c >> shift) & 0xFF);
        string Part(int shift) => ((int)Math.Round(Channel(x, shift) + (Channel(y, shift) - Channel(x, shift)) * t)).ToString("X2");
        return "#" + Channel(x, 24).ToString("X2") + Part(16) + Part(8) + Part(0);
    }

    // ---------- figures ----------

    string Name(string prefix)
    {
        string name;
        do
        {
            name = prefix + (++counter);
        }
        while (names.Contains(name));
        return name;
    }

    XElement Add(string kind, string name, params object[] content)
    {
        if (!names.Add(name))
        {
            throw new Exception("two figures called " + name);
        }

        var element = new XElement(kind, new XAttribute("Name", name), content);
        figures.Add(element);
        return element;
    }

    static IEnumerable<XElement> Dependencies(params string[] dependencies) => dependencies.Select(d => new XElement("Dependency", new XAttribute("Name", d)));

    /// <summary>The figures an expression names: identifiers that are figures of the drawing, not property names after a dot</summary>
    IEnumerable<string> Named(params string[] expressions) => expressions
        .SelectMany(e => Regex.Matches(e, @"(?<![\w.])[A-Za-z_][A-Za-z0-9_]*").Select(m => m.Value))
        .Where(names.Contains)
        .Distinct();

    /// <summary>A hidden point at the expressions' place</summary>
    public V Pt(V at, string name = null)
    {
        name ??= Name("p");
        Add("PointByCoordinates", name, new XAttribute("Visible", "false"), new XAttribute("X", at.X), new XAttribute("Y", at.Y), Dependencies(Named(at.X, at.Y).ToArray()));
        return V.Of(name);
    }

    public V Pt(double x, double y, string name = null) => Pt(V.At(x, y), name);

    /// <summary>A visible point at the expressions' place: a dot that can't be dragged</summary>
    public V DotAt(V at, string name = null)
    {
        name ??= Name("d");
        Add("PointByCoordinates", name, new XAttribute("Style", Dot()), new XAttribute("X", at.X), new XAttribute("Y", at.Y), Dependencies(Named(at.X, at.Y).ToArray()));
        return V.Of(name);
    }

    /// <summary>A hidden number worked out by a formula</summary>
    public string Num(string formula, string name = null)
    {
        name ??= Name("n");
        Add("Number", name, new XAttribute("Formula", formula), Dependencies(Named(formula).ToArray()));
        return name;
    }

    public V Free(double x, double y, string name = null)
    {
        name ??= Name("h");
        Add("FreePoint", name, new XAttribute("Style", Handle()), new XAttribute("X", N(x)), new XAttribute("Y", N(y)));
        return V.Of(name);
    }

    /// <summary>A handle sliding along a hidden segment from a to b, at the fraction t of the way</summary>
    public V Slider(V a, V b, double t, (double X, double Y) at, string name = null)
    {
        var from = Pt(a);
        var to = Pt(b);
        string track = Name("track");
        Add("Segment", track, new XAttribute("Visible", "false"), Dependencies(PointName(from), PointName(to)));
        name ??= Name("h");
        Add("PointOnFigure", name, new XAttribute("Style", Handle()), new XAttribute("X", N(at.X)), new XAttribute("Y", N(at.Y)), new XAttribute("Parameter", N(t)), Dependencies(track));
        return V.Of(name);
    }

    /// <summary>A handle sliding along a hidden arc of the circle around a fixed center, counterclockwise from one angle to another (degrees), at the given angle (a point on an arc keeps its angle from the center as its parameter)</summary>
    public V ArcSlider(V center, double radius, double fromDegrees, double toDegrees, double atDegrees, (double X, double Y) centerAt, string name = null)
    {
        var start = Pt(center + V.At(radius * Math.Cos(fromDegrees * Math.PI / 180), radius * Math.Sin(fromDegrees * Math.PI / 180)));
        var end = Pt(center + V.At(radius * Math.Cos(toDegrees * Math.PI / 180), radius * Math.Sin(toDegrees * Math.PI / 180)));
        string track = Name("track");
        Add("CircleArc", track, new XAttribute("Visible", "false"), Dependencies(PointName(center), PointName(start), PointName(end)));
        name ??= Name("h");
        double a = atDegrees * Math.PI / 180;
        Add("PointOnFigure", name, new XAttribute("Style", Handle()),
            new XAttribute("X", N(centerAt.X + radius * Math.Cos(a))), new XAttribute("Y", N(centerAt.Y + radius * Math.Sin(a))),
            new XAttribute("Parameter", N(a)), Dependencies(track));
        return V.Of(name);
    }

    static string PointName(V point) => point.X.EndsWith(".X") && point.Y == point.X.Substring(0, point.X.Length - 2) + ".Y"
        ? point.X.Substring(0, point.X.Length - 2)
        : throw new Exception("not a point: " + point);

    string Points(V point) => names.Contains(PointName(point)) ? PointName(point) : throw new Exception("unknown point " + point);

    public string Polygon(string style, params V[] points)
    {
        string name = Name("poly");
        Add("Polygon", name, new XAttribute("Style", style), Dependencies(points.Select(Points).ToArray()));
        return name;
    }

    /// <summary>A rectangle with sides along the axes, between two corners given as expressions</summary>
    public string Box(string style, V lowerLeft, V upperRight)
    {
        var a = Pt(lowerLeft);
        var b = Pt(new V(upperRight.X, lowerLeft.Y));
        var c = Pt(upperRight);
        var d = Pt(new V(lowerLeft.X, upperRight.Y));
        return Polygon(style, a, b, c, d);
    }

    public string Segment(string style, V a, V b, string decoration = null)
    {
        string name = Name("s");
        Add("Segment", name, new XAttribute("Style", style), decoration == null ? null : new XAttribute("Decoration", decoration), Dependencies(Points(a), Points(b)));
        return name;
    }

    public string Vector(string style, V a, V b)
    {
        string name = Name("v");
        Add("Vector", name, new XAttribute("Style", style), Dependencies(Points(a), Points(b)));
        return name;
    }

    public string Circle(string style, V center, V through)
    {
        string name = Name("c");
        Add("Circle", name, new XAttribute("Style", style), Dependencies(Points(center), Points(through)));
        return name;
    }

    /// <summary>An arc, sector or circular segment ("CircleArc", "CircleSector", "CircleSegment"), counterclockwise from start to end</summary>
    public string Arc(string kind, string style, V center, V start, V end, bool visible = true)
    {
        string name = Name("a");
        Add(kind, name, style == null ? null : new XAttribute("Style", style), visible ? null : new XAttribute("Visible", "false"), Dependencies(Points(center), Points(start), Points(end)));
        return name;
    }

    public string Region(string style, string operation, params string[] operands)
    {
        string name = Name("r");
        Add("Region", name, new XAttribute("Style", style), new XAttribute("Operation", operation), Dependencies(operands));
        return name;
    }

    /// <summary>An angle's mark at the vertex, between the two sides (the angle under 180°)</summary>
    public string Angle(string style, V vertex, V a, V b, double radius = 26)
    {
        string name = Name("angle");
        Add("AngleArc", name, new XAttribute("Style", style), new XAttribute("Radius", N(radius)), Dependencies(Points(vertex), Points(a), Points(b)));
        return name;
    }

    /// <summary>
    /// An angle's mark as a plain arc of the given radius (units), counterclockwise from the side
    /// toward a to the side toward b: an angle mark (AngleArc) draws a square at 90°, which would say
    /// the answer when the answer is 90°
    /// </summary>
    public string ArcMark(string style, V vertex, V a, V b, double radius)
    {
        string v = PointName(vertex);
        string toA = Num($"distance({v}, {PointName(a)})");
        string toB = Num($"distance({v}, {PointName(b)})");
        var start = Pt(vertex + Lgf.N(radius) * new V($"({PointName(a)}.X - {v}.X) / {toA}", $"({PointName(a)}.Y - {v}.Y) / {toA}"));
        var end = Pt(vertex + Lgf.N(radius) * new V($"({PointName(b)}.X - {v}.X) / {toB}", $"({PointName(b)}.Y - {v}.Y) / {toB}"));
        return Arc("CircleArc", style, vertex, start, end);
    }

    /// <summary>
    /// A text centered on a place (the offsets put the middle of a short text there). A text attached
    /// to a point hides with it, so the point is shown, in a style that draws nothing. Neither is
    /// locked: a point counts as locked when anything built on it is, and the handles would be.
    /// </summary>
    public string Label(string text, V at, string style = null, double size = 26)
    {
        style ??= Text();
        string anchor = Name("at");
        Add("PointByCoordinates", anchor, new XAttribute("Style", Anchor()), new XAttribute("X", at.X), new XAttribute("Y", at.Y), Dependencies(Named(at.X, at.Y).ToArray()));
        string name = Name("t");
        // the box of a bold digit is about 0.62 of the size wide per character and 1.35 high, with the label's padding
        double width = text.Length * size * 0.62 + 6;
        double height = size * 1.35 + 4;
        Add("AttachedLabel", name, new XAttribute("Style", style), new XAttribute("Z", "1"), new XAttribute("OffsetX", N(-width / 2)), new XAttribute("OffsetY", N(-height / 2)), new XAttribute("Text", text), Dependencies(anchor));
        return name;
    }

    /// <summary>A point that draws nothing: what a text hangs from</summary>
    string Anchor()
    {
        const string name = "Anchor";
        if (styleNames.Add(name))
        {
            styles.Add(new XElement("PointStyle", new XAttribute("Name", name), new XAttribute("Size", "1"), new XAttribute("Fill", "#00000000"), new XAttribute("Color", "#00000000"), new XAttribute("StrokeWidth", "0")));
        }

        return name;
    }

    /// <summary>A dimension: a double arrow beside the segment from a to b, the given distance to its left (in units), with a text in a gap in the middle</summary>
    public void Dimension(string lineStyle, string textStyle, V a, V b, double offset, string text, double gap = 0.45)
    {
        string length = Num($"distance({PointName(a)}, {PointName(b)})");
        var along = Pt(new V($"({PointName(b)}.X - {PointName(a)}.X) / {length}", $"({PointName(b)}.Y - {PointName(a)}.Y) / {length}"));
        var normal = new V(V.Neg(along.Y), along.X);
        var from = Pt(a + offset * normal);
        var to = Pt(b + offset * normal);
        var middle = Pt(new V($"({PointName(from)}.X + {PointName(to)}.X) / 2", $"({PointName(from)}.Y + {PointName(to)}.Y) / 2"));
        var left = Pt(middle - gap * along);
        var right = Pt(middle + gap * along);
        Vector(lineStyle, left, from);
        Vector(lineStyle, right, to);
        Label(text, middle, textStyle);
    }

    public string ToText()
    {
        // the scenes are written tight around the figure: the paper around it is added here
        double margin = 0.07 * Math.Max(scene.Right - scene.Left, scene.Top - scene.Bottom);
        var (left, top, right, bottom) = (scene.Left - margin, scene.Top + margin, scene.Right + margin, scene.Bottom - margin);
        scene = (left, top, right, bottom);
        var root = new XElement("Drawing", new XAttribute("Version", "1"),
            new XElement("Viewport", new XAttribute("Left", N(scene.Left)), new XAttribute("Top", N(scene.Top)), new XAttribute("Right", N(scene.Right)), new XAttribute("Bottom", N(scene.Bottom))),
            new XElement("Scene", new XAttribute("Left", N(scene.Left)), new XAttribute("Top", N(scene.Top)), new XAttribute("Right", N(scene.Right)), new XAttribute("Bottom", N(scene.Bottom))),
            styles,
            figures);
        var settings = new XmlWriterSettings { Indent = true, IndentChars = "  ", NewLineChars = "\r\n", Encoding = new UTF8Encoding(false) };
        using var stream = new MemoryStream();
        using (var writer = XmlWriter.Create(stream, settings))
        {
            new XDocument(new XDeclaration("1.0", "utf-8", null), root).Save(writer);
        }

        return Encoding.UTF8.GetString(stream.ToArray()) + "\r\n";
    }
}

static class Puzzles
{
    const double S3 = 0.8660254037844386; // sqrt(3) / 2
    static readonly string s3 = "0.8660254037844386";

    /// <summary>A regular hexagon with a horizontal side from the first vertex to the right (flat top and bottom), counterclockwise: its first vertex and side as expressions</summary>
    static V[] HexagonUp(Lgf d, V start, string side)
    {
        return
        [
            d.Pt(start),
            d.Pt(start + new V(side, "0")),
            d.Pt(start + new V($"1.5 * {side}", $"{s3} * {side}")),
            d.Pt(start + new V(side, $"2 * {s3} * {side}")),
            d.Pt(start + new V("0", $"2 * {s3} * {side}")),
            d.Pt(start + new V($"-0.5 * {side}", $"{s3} * {side}")),
        ];
    }

    /// <summary>A regular hexagon with flat top and bottom, given its left vertex and side: left, top left, top right, right, bottom right, bottom left</summary>
    static V[] HexagonFromLeft(Lgf d, V left, string side)
    {
        return
        [
            d.Pt(left),
            d.Pt(left + new V($"0.5 * {side}", $"{s3} * {side}")),
            d.Pt(left + new V($"1.5 * {side}", $"{s3} * {side}")),
            d.Pt(left + new V($"2 * {side}", "0")),
            d.Pt(left + new V($"1.5 * {side}", $"-{s3} * {side}")),
            d.Pt(left + new V($"0.5 * {side}", $"-{s3} * {side}")),
        ];
    }

    static V Mid(Lgf d, params V[] points) => d.Pt(new V(
        $"({string.Join(" + ", points.Select(p => p.X))}) / {points.Length}",
        $"({string.Join(" + ", points.Select(p => p.Y))}) / {points.Length}"));

    // Three regular hexagons; the smallest has area 6, the pink one? The yellow one has two vertices on
    // the circle; the orange one hangs from its top right vertex and the pink one from its bottom right,
    // each with a vertex on the circle. Dragging the yellow one's left vertex along the circle resizes all
    // three, and the pink one stays four times the orange one.
    public static void ThreeHexagonsOnACircle(Lgf d)
    {
        const double R = 5;
        var o = d.Pt(0, 0, "O");
        var rim = d.Pt(R, 0);
        d.Circle(d.Line("Circle", "#FFE5352B", 2), o, rim);
        var l = d.ArcSlider(o, R, 158, 176.5, 172, (0, 0), "L");
        string a = d.Num($"-2 * (0.5 * L.X + {s3} * L.Y)", "a");
        var yellow = HexagonFromLeft(d, l, a);
        string orange = d.Num($"(-{yellow[2].X} + sqrt({R * R} - {yellow[2].Y} ^ 2)) / 2", "o");
        string b = d.Num($"0.5 * {yellow[4].X} - {s3} * {yellow[4].Y}", "pb");
        string pink = d.Num($"-{b} + sqrt({b} ^ 2 - ({yellow[4].X} ^ 2 + {yellow[4].Y} ^ 2 - {R * R}))", "pk");
        var small = HexagonFromLeft(d, yellow[2], orange);
        var big = HexagonFromLeft(d, yellow[4], pink);
        d.Polygon(d.Fill("Yellow", "#FFFFE000"), yellow);
        d.Polygon(d.Fill("Orange", "#FFF26313"), small);
        d.Polygon(d.Fill("Pink", "#FFF5589B"), big);
        d.Label("6", Mid(d, small[0], small[3]), d.ShapeText());
        d.Label("?", Mid(d, big[0], big[3]), d.ShapeText());
        d.Scene(-5.6, 5.7, 6.4, -5.8);
    }

    // An equilateral triangle inside a regular hexagon, one vertex at a corner of the hexagon and one on
    // the long diagonal. Dragging that one along the diagonal swings the third vertex, and the blue line
    // from it to the next corner keeps its angle.
    public static void ATriangleInsideAHexagon(Lgf d)
    {
        var right = d.Pt(4, 0);
        var topRight = d.Pt(2, 4 * S3);
        var topLeft = d.Pt(-2, 4 * S3, "D");
        var left = d.Pt(-4, 0);
        var bottomLeft = d.Pt(-2, -4 * S3);
        var bottomRight = d.Pt(2, -4 * S3);
        var h = d.Slider(V.At(-4, 0), V.At(-0.8, 0), 0.5, (-2.4, 0), "H");
        var i = d.Pt(topLeft + (h - topLeft).Turn(60), "I");
        d.Polygon(d.Fill("LightGreen", "#FFD3F7A6"), left, right, topRight, topLeft);
        d.Polygon(d.Fill("Green", "#FF7FD8B6"), left, bottomLeft, bottomRight, right);
        d.Polygon(d.Fill("Purple", "#FFC29BF0"), topLeft, h, i);
        string blue = d.Line("Blue", "#FF1D3C99", 2.5);
        d.Segment(blue, i, topRight);
        d.Angle(blue, topRight, i, right, 30);
        d.Label("?", topRight + V.At(0.15, -1.25), d.Text("BlueMark", "#FF1D3C99", 24));
        d.Scene(-4.4, 3.9, 4.4, -3.9);
    }

    // A quarter circle and a semicircle on its vertical radius. The two blue angles are equal; the orange
    // angle? The point on the quarter circle is dragged, and the point on the semicircle follows so that
    // the blue angles stay equal: it is where the line from the dragged point through the middle of the
    // semicircle's other half meets it again.
    public static void ASemiCircleAndAQuarterCircle(Lgf d)
    {
        var q = d.Pt(0, 0, "Q");
        var a = d.Pt(-4, 0, "A");
        var t = d.Pt(0, 4, "T");
        var m = d.Pt(0, 2, "M");
        var k = d.Pt(-2, 2, "K");
        var p =d.ArcSlider(q, 4, 138, 177, 157, (0, 0), "P");
        string length = d.Num("distance(K, P)", "kp");
        string dx = d.Num($"(K.X - P.X) / {length}", "dx");
        string dy = d.Num($"(K.Y - P.Y) / {length}", "dy");
        var s = d.Pt(new V($"K.X + 4 * {dx} ^ 2", $"K.Y + 4 * {dx} * {dy}"), "S");
        d.Arc("CircleSector", d.Fill("Orange", "#FFF9B46A"), q, t, a);
        d.Arc("CircleSector", d.Fill("Yellow", "#FFFDF35A"), m, q, t);
        string red = d.Line("Red", "#FFE8402A", 2.5);
        d.Segment(red, a, p);
        d.Segment(red, p, t);
        d.Segment(red, p, s);
        d.Segment(red, t, s);
        d.Segment(red, s, q);
        string wedge = d.Wedge("BlueWedge", "#FF1414D2");
        d.Angle(wedge, p, a, s, 22);
        d.Angle(wedge, t, p, s, 22);
        d.Angle(d.Line("OrangeLine", "#FFF07A12", 3), s, p, q, 26);
        d.Label("?", s + V.At(-1.05, -0.6), d.Text("OrangeMark", "#FFE8710E", 24));
        d.Scene(-4.3, 4.3, 2.3, -0.3);
    }

    // Seven squares: six small ones in a net and a big one with a corner at the bottom of the net's middle
    // column and its far side through the top of it. Dragging the big square's corner turns it (its far
    // side always through that point), and the angle stays.
    public static void SevenSquares(Lgf d)
    {
        var v0 = d.Pt(1, 0, "V");
        var g = d.Pt(3, 2, "G");
        var c = d.Pt(1, 2, "C");
        var w = d.ArcSlider(c, 2, 96, 172, 133, (1, 2), "W");
        var side = w - v0;
        var x = d.Pt(v0 + new V(side.Y, V.Neg(side.X)));
        var y = d.Pt(w + new V(side.Y, V.Neg(side.X)));
        d.Polygon(d.Fill("Lavender", "#FFC2BDEA"), v0, x, y, w);
        string pink = d.Fill("Pink", "#FFF7C1EB");
        foreach (var (i, j) in new[] { (0, 0), (0, 1), (1, 1), (1, 2), (1, 3), (2, 2) })
        {
            d.Box(pink, V.At(i, j), V.At(i + 1, j + 1));
        }

        string blue = d.Line("Blue", "#FF2A2AB0", 2.5);
        d.Segment(blue, w, g);
        d.Angle(blue, w, v0, g, 26);
        d.Label("?", w + V.At(0.2, -0.65), d.Text("BlueMark", "#FF2A2AB0", 22));
        d.Scene(-1.5, 5.2, 4.8, -0.4);
    }

    // Six squares and a quarter circle, the purple square with two corners on the arc; the blue square
    // has area 2, the yellow? Dragging the purple square's corner along the arc changes the sizes, and
    // the yellow square stays twice as wide as the blue.
    public static void SixSquaresAndAQuarterCircle(Lgf d)
    {
        const double R = 10;
        var c = d.Pt(0, 0, "C");
        var top = d.Pt(0, R);
        var left = d.Pt(-R, 0);
        var h = d.ArcSlider(c, R, 136.5, 145.5, 142.2, (0, 0), "H");
        string b = d.Num("-H.X - H.Y", "b");
        string r = d.Num("2 * H.Y + H.X", "r");
        string g = d.Num($"2 * {b}", "g");
        d.Box(d.Fill("Yellow", "#FFFCEE18"), new V($"-{g}", "0"), new V("0", g));
        d.Box(d.Fill("Green", "#FF14B67F"), new V($"-{g}", g), new V("0", $"2 * {g}"));
        d.Box(d.Fill("Red", "#FFFD4246"), new V($"-{g} - {r}", "0"), new V($"-{g}", r));
        d.Box(d.Fill("Blue", "#FF22D4F5"), new V($"-{g} - {r}", r), new V($"-{g} - {r} + {b}", $"{r} + {b}"));
        d.Box(d.Fill("Purple", "#FFA98BE0"), new V($"-{g} - {r}", $"{r} + {b}"), new V($"-{g} - {r} + {b}", $"{r} + 2 * {b}"));
        d.Box(d.Fill("Orange", "#FFFDA94E"), new V($"-{g} - {r} + {b}", r), new V($"-{g}", $"2 * {r} - {b}"));
        d.Arc("CircleSector", d.Outline("Frame", Lgf.Ink, 3), c, top, left);
        d.Label("2", new V($"-{g} - {r} + {b} / 2", $"{r} + {b} / 2"), d.ShapeText());
        d.Label("?", new V($"-{g} / 2", $"{g} / 2"), d.ShapeText());
        d.Scene(-10.4, 10.4, 0.4, -0.4);
    }

    // Two regular hexagons in a rectangle whose diagonal runs along a side of each. Nothing to drag: the
    // rectangle is the one the hexagons fit.
    public static void TwoHexagonsInARectangleII(Lgf d)
    {
        const double s = 2, h = 2.5 * s, w = 2.5 * 2 * S3 * s;
        var bl = d.Pt(0, 0);
        var br = d.Pt(w, 0);
        var tr = d.Pt(w, h);
        var tl = d.Pt(0, h);
        d.Polygon(d.Fill("LightBlue", "#FF2FD3F8"), tl, tr, br);
        d.Polygon(d.Fill("Blue", "#FF178EE8"), tl, br, bl);
        string paper = d.PaperFill();
        foreach (var (cx, cy, sign) in new[] { (S3 * s, s, 1.0), (w - S3 * s, h - s, -1.0) })
        {
            var corners = Enumerable.Range(0, 6).Select(i => d.Pt(cx + s * Math.Cos((90 + 60 * i) * Math.PI / 180), cy + s * Math.Sin((90 + 60 * i) * Math.PI / 180))).ToArray();
            d.Polygon(paper, corners);
        }

        d.Scene(-0.4, h + 0.4, w + 0.4, -0.4);
    }

    // Three congruent rectangles, two lines from the top corner making two white triangles. Dragging
    // the blue rectangle along the top of the green one, or the green one's corner to change the shape
    // of all three: the white triangles always add up to a quarter of a rectangle... (the fraction shaded stays).
    public static void ThreeCongruentRectanglesVI(Lgf d)
    {
        var o = d.Pt(0, 0, "O");
        var corner = d.Free(5, 3, "G");
        var a = d.Slider(new V("0.05 * G.Y", "G.Y"), new V("0.95 * G.Y", "G.Y"), (1.2 / 3 - 0.05) / 0.9, (1.2, 3), "A");
        string w = "G.X", h = "G.Y";
        var t = d.Pt(new V($"A.X + {w}", $"2 * {h}"), "T");
        var br = d.Pt(new V($"A.X + {w}", h), "Br");
        var dl = d.Pt(new V(w, $"{h} - {w}"), "Dl");
        d.Box(d.Fill("Blue", "#FF3CD2F5"), a, t);
        d.Box(d.Fill("LightGreen", "#FF9BEB74"), o, corner);
        d.Box(d.Fill("Green", "#FF3CCBA3"), dl, new V($"{w} + {h}", h));
        string paper = d.PaperFill();
        d.Polygon(paper, t, a, o);
        d.Polygon(paper, t, br, dl);
        d.Scene(-0.4, 6.3, 8.4, -2.4);
    }

    // Four quarter circles; the red one has area 12, the orange? Every quarter circle meets at the point on
    // the diagonal, which is dragged: the red one and the orange one change, the orange one stays twice the red.
    public static void FourQuarterCirclesIII(Lgf d)
    {
        const double R = 10;
        var o = d.Pt(0, 0, "O");
        var t = d.Pt(0, R, "T");
        var k = d.Pt(R, 0, "K");
        var p = d.Slider(V.At(0.15 * R, 0.15 * R), V.At(0.42 * R, 0.42 * R), (0.3 - 0.15) / 0.27, (0.3 * R, 0.3 * R), "P");
        var f = d.Pt(new V("P.X", "0"), "F");
        var c1 = d.Pt(new V(Lgf.N(R / 2), $"{Lgf.N(R / 2)} + P.X"), "C1");
        var e = d.Pt(new V($"{Lgf.N(R)} + P.X", $"{Lgf.N(R)} - P.X"), "E");
        string ink = d.Line("Ink");
        d.Segment(ink, o, t);
        d.Segment(ink, o, k);
        d.Arc("CircleSector", d.Fill("Pink", "#FFFBA6CC"), f, p, o);
        d.Arc("CircleSector", d.Fill("Red", "#FFD97689"), c1, t, p);
        d.Arc("CircleSector", d.Fill("Orange", "#FFF9C268"), k, e, p);
        d.Arc("CircleArc", ink, o, k, t);
        d.Label("12", c1 + 0.42 * ((t - c1) + (p - c1)), d.ShapeText());
        d.Label("?", k + 0.42 * ((e - k) + (p - k)), d.ShapeText());
        d.Scene(-0.5, 10.6, 13.6, -0.5);
    }

    // Two congruent isosceles triangles in a semicircle, sharing the end of its diameter; the purple one's
    // legs are 4, the semicircle's area? Dragging the blue one's corner along the arc: the triangles change,
    // the legs stay the radius.
    public static void TwoTrianglesInASemiCircle(Lgf d)
    {
        const double R = 4;
        var m = d.Pt(0, 0, "M");
        var c = d.Pt(R, 0, "C");
        var west = d.Pt(-R, 0);
        d.Arc("CircleSegment", d.Outline("Frame"), m, c, west);
        var a1 = d.ArcSlider(m, R, 100, 160, 135.6, (0, 0), "A");
        string length = d.Num("distance(C, A)", "b");
        string ux = d.Num($"(A.X - C.X) / {length}", "ux");
        string uy = d.Num($"(A.Y - C.Y) / {length}", "uy");
        var b1 = d.Pt(new V($"C.X - {R} * ({ux} ^ 2 - {uy} ^ 2)", $"-{2 * R} * {ux} * {uy}"), "B");
        var b2 = d.Pt(new V($"C.X + {R} * {ux}", $"{R} * {uy}"), "B2");
        var a2 = d.Pt(new V($"C.X - {length}", "0"), "A2");
        d.Polygon(d.Fill("Blue", "#FF1689F2"), a1, b1, c);
        d.Polygon(d.Fill("Purple", "#FF8A12B5"), a2, b2, c);
        d.Dimension(d.Line("PurpleArrow", "#FF8A12B5", 2.5), d.Text("PurpleMark", "#FF8A12B5", 26), a2, b2, 0.4, "4", 0.42);
        d.Scene(-4.4, 4.4, 4.4, -0.5);
    }

    // Four squares: two yellow of area 9, the blue? and a big one on the blue one's corner and the
    // top yellow one's, its side through the left yellow one's corner. Nothing to drag: the condition fixes it.
    public static void FourSquaresX(Lgf d)
    {
        const double a = 3;
        double b = a * Math.Sqrt(2);
        var o = d.Pt(0, 0);
        var v = d.Pt(a, a + b);
        var u = d.Pt(a - (a + b), a + b + a);
        var w = d.Pt(-(a + b), a);
        d.Polygon(d.PaperFill(), o, v, u, w);
        d.Box(d.Fill("Blue", "#FF06BEF8"), V.At(0, 0), V.At(b, b));
        string yellow = d.Fill("Yellow", "#FFFCF51E");
        d.Box(yellow, V.At(0, b), V.At(a, b + a));
        d.Box(yellow, V.At(-a, b - a), V.At(0, b));
        d.Segment(d.Line("Ink"), o, v);
        d.Label("9", V.At(a / 2, b + a / 2), d.ShapeText());
        d.Label("9", V.At(-a / 2, b - a / 2), d.ShapeText());
        d.Label("?", V.At(b * 0.68, b * 0.42), d.ShapeText());
        d.Scene(-7.7, 10.7, 4.6, -0.4);
    }

    // A green square with a corner at the end of a row of three squares and its sides through two
    // corners of the first one. The angle at its far corner, between its side and the line to the row's
    // first corner? Nothing to drag: the row decides the square.
    public static void FourSquaresIX(Lgf d)
    {
        var w = d.Pt(1.2, 0.4);
        var v = d.Pt(6, 2);
        var top = d.Pt(2.8, 3.6);
        var bottom = d.Pt(4.4, -1.2);
        var origin = d.Pt(0, 0);
        d.Polygon(d.Fill("Green", "#FF62D3AF"), w, bottom, v, top);
        string frame = d.Outline("Frame");
        d.Box(frame, V.At(0, 0), V.At(2, 2));
        d.Box(frame, V.At(2, 0), V.At(4, 2));
        d.Box(frame, V.At(4, 0), V.At(6, 2));
        string magenta = d.Line("Magenta", "#FFD0249C", 2.5);
        d.Segment(magenta, origin, w);
        d.Angle(magenta, w, origin, top, 22);
        d.Label("?", V.At(0.95, 0.95), d.Text("MagentaMark", "#FFD0249C", 22));
        d.Scene(-0.4, 4.0, 6.4, -1.6);
    }

    // Three regular hexagons: the red one and the yellow one on the diameter sharing a corner, the blue
    // one, of side 2, on their outer top corners, which are on the semicircle. Its diameter? Dragging the
    // blue one's corner along the arc changes the red one and the yellow one; the semicircle stays.
    public static void ThreeRegularHexagonsAndASemiCircle(Lgf d)
    {
        const double R = 4;
        const string sqrt3 = "1.7320508075688772";
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0);
        var west = d.Pt(-R, 0);
        var y0 = d.ArcSlider(o, R, 48, 104, 73, (0, 0), "Y");
        var r0 = d.Pt(V.Of("Y").Turn(60), "Rr");
        var p = d.Pt(new V($"Y.X - Y.Y / {sqrt3}", "0"), "P");
        string ys = d.Num($"Y.Y / {sqrt3}", "ys");
        string rs = d.Num($"Rr.Y / {sqrt3}", "rs");
        var yellow = HexagonUp(d, p, ys);
        var red = new[]
        {
            p,
            d.Pt(p + new V($"-{rs}", "0")),
            d.Pt(p + new V($"-1.5 * {rs}", $"{s3} * {rs}")),
            r0,
            d.Pt(p + new V("0", $"2 * {s3} * {rs}")),
            d.Pt(p + new V($"0.5 * {rs}", $"{s3} * {rs}")),
        };
        var center = Mid(d, y0, r0);
        var blue = Enumerable.Range(0, 6).Select(i => i == 0 ? y0 : i == 3 ? r0 : d.Pt(center + (y0 - center).Turn(60 * i))).ToArray();
        string none = null;
        string yellowShape = d.Polygon(d.Fill("Yellow", "#FFFCF56E", none), yellow);
        string redShape = d.Polygon(d.Fill("Red", "#FFE56F9B", none), red);
        string blueShape = d.Polygon(d.Fill("Blue", "#FF68CCF7", none), blue);
        d.Region(d.Fill("Green", "#FFC2E58A", none), "Intersection", blueShape, yellowShape);
        d.Region(d.Fill("Purple", "#FF8E6FDF", none), "Intersection", blueShape, redShape);
        d.Region(d.Fill("Orange", "#FFF8B377", none), "Intersection", redShape, yellowShape);
        d.Region(d.PaperFill("Hole", none), "Intersection", blueShape, redShape, yellowShape);
        string frame = d.Outline("Frame");
        d.Polygon(frame, yellow);
        d.Polygon(frame, red);
        d.Polygon(frame, blue);
        d.Arc("CircleSegment", frame, o, east, west);
        string ink = d.Line("Arrow", Lgf.Ink, 2.5);
        d.Dimension(ink, d.Text(), y0, blue[1], -0.4, "2", 0.38);
        d.Dimension(ink, d.Text(), east, west, 0.55, "?", 0.38);
        d.Scene(-4.4, 6.2, 4.4, -1.1);
    }

    // Four quarter circles; what fraction is shaded? The white quarter circle has its corner on the
    // vertical radius and its other end on the arc, the inner left one reaches the same height. Dragging
    // that end along the arc: the shaded parts change, together they stay half.
    public static void FourQuarterCirclesII(Lgf d)
    {
        const double R = 6;
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0);
        var north = d.Pt(0, R);
        var west = d.Pt(-R, 0);
        var q = d.ArcSlider(o, R, 46, 80, 56, (0, 0), "Q");
        var corner = d.Pt(new V("0", "Q.Y"), "Pc");
        var foot = d.Pt(new V("0", "Q.Y - Q.X"), "F");
        var inner = d.Pt(new V("-Q.Y", "0"), "W");
        string rightQuarter = d.Arc("CircleSector", null, o, east, north, visible: false);
        string white = d.Arc("CircleSector", null, corner, foot, q, visible: false);
        string leftQuarter = d.Arc("CircleSector", null, o, north, west, visible: false);
        string innerQuarter = d.Arc("CircleSector", null, o, corner, inner, visible: false);
        string pink = d.Fill("Pink", "#FFF9A4F1", null);
        d.Region(pink, "Difference", rightQuarter, white);
        d.Region(pink, "Difference", leftQuarter, innerQuarter);
        string ink = d.Line("Ink");
        d.Arc("CircleSegment", d.Outline("Frame"), o, east, west);
        d.Segment(ink, o, north);
        d.Segment(ink, corner, q);
        d.Arc("CircleArc", ink, o, corner, inner);
        d.Arc("CircleArc", ink, corner, foot, q);
        d.Scene(-6.4, 6.4, 6.4, -0.4);
    }

    // A triangle, one side halved, another cut in three; what fraction is shaded? Every corner can be
    // dragged: the fraction stays.
    public static void SubdividedTriangleII(Lgf d)
    {
        var a = d.Free(0, 0, "A");
        var t = d.Free(5.25, 4.2, "T");
        var c = d.Free(9.4, 0, "C");
        d.Polygon(d.PaperFill(), a, t, c);
        var m = d.DotAt(Mid(d, a, t), "M");
        var p1 = d.DotAt(t + (1.0 / 3) * (c - t), "P1");
        var p2 = d.DotAt(t + (2.0 / 3) * (c - t), "P2");
        d.Polygon(d.Fill("Orange", "#FFFE9001", null), a, m, p1, p2);
        string ink = d.Line("Ink");
        d.Segment(ink, a, m, "TwoTicks");
        d.Segment(ink, m, t, "TwoTicks");
        d.Segment(ink, t, p1, "OneTick");
        d.Segment(ink, p1, p2, "OneTick");
        d.Segment(ink, p2, c, "OneTick");
        d.Scene(-0.5, 4.7, 9.9, -0.5);
    }

    // Four squares: three the same size in a big one; the angle at its corner? Dragging the corner of the
    // left small square resizes the three: the angle stays.
    public static void FourSquaresXII(Lgf d)
    {
        const double L = 10;
        var h = d.Slider(V.At(1.5, 3), V.At(4.8, 9.6), (3.96 - 1.5) / 3.3, (3.96, 7.92), "H");
        string s = "H.X";
        var v = d.Pt(L, L, "V");
        d.Box(d.PaperFill(), V.At(0, 0), V.At(L, L));
        string purple = d.Fill("Purple", "#FFEBB7F6", "#FF4A2793");
        d.Box(purple, new V("0", s), new V(s, $"2 * {s}"));
        d.Box(purple, new V($"{L} - 2 * {s}", "0"), new V($"{L} - {s}", s));
        d.Box(purple, new V($"{L} - {s}", "0"), V.At(L, 0) + new V("0", s));
        var a = d.Pt(new V("0", $"2 * {s}"), "Ap");
        var b = d.Pt(new V($"{L} - {s}", s), "Bp");
        string magenta = d.Line("Magenta", "#FFC4269B", 2.5);
        d.Segment(magenta, v, a);
        d.Segment(magenta, v, b);
        d.Angle(magenta, v, a, b, 34);
        d.Label("?", V.At(L - 1.85, L - 1.55), d.Text("MagentaMark", "#FFC4269B", 24));
        d.Scene(-0.4, 10.4, 10.4, -0.4);
    }

    // A big square, a blue rectangle of two squares (area 12 each) hanging from its corner down to its
    // bottom side, and a green square between the rectangle's middle and the big square's other corner.
    // Dragging the rectangle's foot along the bottom tilts it: the green square stays twice the blue one.
    public static void FourSquaresXI(Lgf d)
    {
        const double L = 8;
        string navy = "#FF22307A", teal = "#FF2A8C8C";
        var tl = d.Pt(0, L, "TL");
        var br = d.Pt(L, 0, "BR");
        var q = d.Slider(V.At(0.3, 0), V.At(5.5, 0), (1.96 - 0.3) / 5.2, (1.96, 0), "Q");
        string a = d.Num("distance(TL, Q) / 2", "a");
        string dx = d.Num($"(Q.X - TL.X) / (2 * {a})", "dx");
        string dy = d.Num($"(Q.Y - TL.Y) / (2 * {a})", "dy");
        var down = new V(dx, dy);
        var right = new V(V.Neg(dy), dx);
        var midLeft = d.Pt(tl + a * down);
        var m = d.Pt(midLeft + a * right, "M");
        var topRight = d.Pt(tl + a * right);
        var bottomRight = d.Pt(q + a * right);
        d.Box(d.PaperFill("Frame", navy), V.At(0, 0), V.At(L, L));
        var side = br - m;
        var g2 = d.Pt(m + new V(V.Neg(side.Y), side.X));
        var g3 = d.Pt(br + new V(V.Neg(side.Y), side.X));
        d.Polygon(d.Fill("Green", "#FFC8E89A", teal), m, br, g3, g2);
        d.Polygon(d.Fill("DarkBlue", "#FF5E99DA", navy), tl, midLeft, m, topRight);
        d.Polygon(d.Fill("LightBlue", "#FF9ECCF0", navy), midLeft, q, bottomRight, m);
        d.Label("12", Mid(d, midLeft, bottomRight), d.ShapeText("NavyMark", navy));
        d.Label("?", Mid(d, m, g3), d.ShapeText("TealMark", teal));
        d.Scene(-0.6, 9.8, 13.6, -0.6);
    }

    // Four squares (a 2 by 2 grid of side 3) and an isosceles triangle whose sides run through grid
    // points. Its area? Nothing to drag: the grid decides it.
    public static void FourSquaresAndAnIsoscelesTriangle(Lgf d)
    {
        var v = d.Pt(3, 3);
        var p = d.Pt(-1, 1);
        var q = d.Pt(1, -1);
        d.Polygon(d.Fill("Blue", "#FF0A93CF"), v, p, q);
        string frame = d.Outline("Frame");
        foreach (var (i, j) in new[] { (0, 0), (1, 0), (0, 1), (1, 1) })
        {
            d.Box(frame, V.At(i * 1.5, j * 1.5), V.At(i * 1.5 + 1.5, j * 1.5 + 1.5));
        }

        var low = d.Pt(3.5, 0);
        var high = d.Pt(3.5, 3);
        d.Dimension(d.Line("Arrow"), d.Text(), high, low, 0, "3", 0.42);
        d.Scene(-1.4, 3.4, 4.1, -1.4);
    }

    // Two congruent right triangles side by side under a quarter circle as high as they are; the angle
    // where the arc crosses the left one's long side? Dragging the corner they share changes their shape:
    // the angle stays.
    public static void TwoTrianglesAndAQuarterCircle(Lgf d)
    {
        const double H = 4;
        var o = d.Pt(0, 0, "O");
        var t = d.Pt(0, H, "T");
        var west = d.Pt(-H, 0);
        var b = d.Slider(V.At(-0.9, 0), V.At(-2.2, 0), (1.52 - 0.9) / 1.3, (-1.52, 0), "B");
        string w = "(-B.X)";
        var far = d.Pt(new V($"2 * B.X", "0"), "F");
        var apex = d.Pt(new V($"2 * B.X", Lgf.N(H)), "G");
        string k = d.Num($"2 * {w} ^ 2 / ({w} ^ 2 + {H * H})", "k");
        var x = d.Pt(new V($"G.X + {k} * {w}", $"{Lgf.N(H)} - {k} * {Lgf.N(H)}"), "X");
        string teal = d.Fill("Teal", "#FF4C9EA3", "#FF1D5458");
        d.Polygon(teal, far, b, apex);
        d.Polygon(teal, b, o, t);
        d.Arc("CircleArc", d.Line("Green", "#FF79C96A", 2.5), o, t, west);
        string ink = d.Line("Ink");
        d.Segment(ink, x, t);
        d.ArcMark(ink, x, b, t, 0.42);
        d.Label("?", x + V.At(0.55, -0.6));
        d.Scene(-4.7, 4.3, 0.4, -0.3);
    }

    // Two squares and a quarter circle: the red square (area 12) in the big one's corner, its sides drawn
    // on across, and lines through where they cross the arc. The yellow area? Dragging the red square's
    // corner along the diagonal: the yellow area stays the red one's.
    public static void TwoSquaresAndAQuarterCircle(Lgf d)
    {
        const double L = 8;
        string blue = "#FF2B3FA0";
        var o = d.Pt(0, 0, "O");
        var top = d.Pt(0, L);
        var west = d.Pt(-L, 0);
        var h = d.Slider(V.At(-1.5, 1.5), V.At(-5.6, 5.6), (4 - 1.5) / 4.1, (-4, 4), "H");
        string k = d.Num($"sqrt({L * L} - H.Y ^ 2)", "k");
        d.Box(d.PaperFill("Frame", blue), V.At(-L, 0), V.At(0, L));
        d.Polygon(d.Fill("Yellow", "#FFFCF886", null),
            d.Pt(-L, 0), d.Pt(new V($"-{k}", "0")), d.Pt(new V($"-{k}", k)), d.Pt(new V("0", k)), d.Pt(0, L), d.Pt(-L, L));
        d.Box(d.Fill("Red", "#FFDC6F93", blue), new V("H.X", "0"), new V("0", "H.Y"));
        string line = d.Line("Line", blue, 2.5);
        d.Segment(line, d.Pt(new V("H.X", "0")), d.Pt(new V("H.X", Lgf.N(L))));
        d.Segment(line, d.Pt(new V(Lgf.N(-L), "H.Y")), d.Pt(new V("0", "H.Y")));
        d.Segment(line, d.Pt(new V($"-{k}", "0")), d.Pt(new V($"-{k}", Lgf.N(L))));
        d.Segment(line, d.Pt(new V(Lgf.N(-L), k)), d.Pt(new V("0", k)));
        d.Arc("CircleArc", d.Line("Arc", blue, 3), o, top, west);
        d.Label("12", new V("H.X / 2", "H.Y / 2"), d.ShapeText());
        d.Scene(-8.4, 8.4, 0.4, -0.4);
    }

    // Five squares in a row, a fan of triangles from one corner to their tops. What fraction is shaded?
    // Every square's corner can be dragged to resize it: the fraction stays.
    public static void FiveSquaresII(Lgf d)
    {
        var p = d.Pt(0, 0, "P");
        var h3 = d.Slider(V.At(1.5, 1.5), V.At(6.5, 6.5), (4.9 - 1.5) / 5, (4.9, 4.9), "S3");
        var h2 = d.Slider(V.At(-1.5, 1.5), V.At(-6, 6), (3.8 - 1.5) / 4.5, (-3.8, 3.8), "S2");
        var h1 = d.Slider(new V("S2.X - 0.8", "0.8"), new V("S2.X - 5", "5"), (2.4 - 0.8) / 4.2, (-6.2, 2.4), "S1");
        var h4 = d.Slider(new V("1.3 * S3.X", "0.3 * S3.X"), new V("1.95 * S3.X", "0.95 * S3.X"), (3.2 / 4.9 - 0.3) / 0.65, (8.1, 3.2), "S4");
        var h5 = d.Slider(new V("S4.X + 0.3 * S4.Y", "0.3 * S4.Y"), new V("S4.X + 0.95 * S4.Y", "0.95 * S4.Y"), (2.0 / 3.2 - 0.3) / 0.65, (10.1, 2.0), "S5");
        string s1 = "S1.Y", s2 = "(-S2.X)", s3 = "S3.X", s4 = "S4.Y", s5 = "S5.Y";
        string paper = d.PaperFill();
        d.Box(paper, new V($"-{s2} - {s1}", "0"), new V($"-{s2}", s1));
        d.Box(paper, new V($"-{s2}", "0"), new V("0", s2));
        d.Box(paper, V.At(0, 0), new V(s3, s3));
        d.Box(paper, new V(s3, "0"), new V($"{s3} + {s4}", s4));
        d.Box(paper, new V($"{s3} + {s4}", "0"), new V($"{s3} + {s4} + {s5}", s5));
        d.Polygon(d.Fill("Blue", "#FF1A2C9C", tint: 0.2), d.Pt(new V($"-{s2} - {s1}", s1)), d.Pt(new V($"-{s2}", s1)), p);
        d.Polygon(d.Fill("Purple", "#FFA3127F", tint: 0.2), d.Pt(new V($"-{s2}", s2)), d.Pt(new V("0", s2)), p);
        d.Polygon(d.Fill("Red", "#FFE5332A", tint: 0.2), d.Pt(new V(s3, s3)), d.Pt(new V(s3, s4)), p);
        d.Polygon(d.Fill("Orange", "#FFEE7D2C", tint: 0.2), d.Pt(new V($"{s3} + {s4}", s4)), d.Pt(new V($"{s3} + {s4}", s5)), p);
        d.Polygon(d.Fill("Yellow", "#FFE8E60C", tint: 0.2), d.Pt(new V($"{s3} + {s4} + {s5}", s5)), d.Pt(new V($"{s3} + {s4} + {s5}", "0")), p);
        d.Scene(-6.7, 5.4, 10.6, -0.4);
    }
}
