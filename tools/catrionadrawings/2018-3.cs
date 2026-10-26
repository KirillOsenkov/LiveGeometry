#nullable disable

static partial class Puzzles
{
    // A regular hexagon standing on a corner, four lines from its top and bottom corners to the far
    // corners of the sides: the rhombus in the middle and the two triangles at the sides shaded. What
    // fraction? Nothing to drag: the hexagon fixes it.
    [Puzzle("crossed-hexagon")]
    public static void CrossedHexagon(Lgf d)
    {
        const double R = 2;
        var h = Enumerable.Range(0, 6).Select(i => d.Pt(R * Math.Cos((90 + 60 * i) * Math.PI / 180), R * Math.Sin((90 + 60 * i) * Math.PI / 180))).ToArray();
        // top, upper left, lower left, bottom, lower right, upper right
        var left = d.Crossing(h[0], h[2], h[3], h[1]);
        var right = d.Crossing(h[0], h[4], h[3], h[5]);
        string blue = d.Fill("Blue", "#FFB9C8F5", null);
        d.Polygon(blue, h[0], left, h[3], right);
        d.Polygon(blue, h[1], h[2], left);
        d.Polygon(blue, h[5], h[4], right);
        string line = d.Line("BlueLine", "#FF2233A0", 2.5);
        d.Segment(line, h[0], h[2]);
        d.Segment(line, h[0], h[4]);
        d.Segment(line, h[3], h[1]);
        d.Segment(line, h[3], h[5]);
        d.Inked(d.Outline("Frame", Lgf.Ink, 3.5), h);
        d.Scene(-R * S3, R, R * S3, -R);
    }

    // A semicircle on the bottom side of a square, as wide as it, and the pink square standing on its
    // diameter with its top corners on the arc. What fraction of the big square is pink? Nothing to
    // drag: it is all one shape (the pink square's side is the big one's over √5).
    [Puzzle("square-in-a-semi-circle-in-a-square")]
    public static void SquareInASemiCircleInASquare(Lgf d)
    {
        const double L = 5;
        double s = L / Math.Sqrt(5);
        var center = d.Pt(L / 2, 0);
        var east = d.Pt(L, 0);
        var west = d.Pt(0, 0);
        d.Box(d.Outline("Frame", "#FF6A14D8", 3.5), V.At(0, 0), V.At(L, L));
        d.Box(d.Fill("Pink", "#FFFC9AB2", null), V.At(L / 2 - s / 2, 0), V.At(L / 2 + s / 2, s));
        d.Arc("CircleArc", d.Line("Arc", "#FF1010F0", 3.5), center, east, west);
        d.Scene(0, L, L, 0);
    }

    // Two equilateral triangles in a semicircle: the big one on the diameter with its apex on the arc
    // (the equal ticks center it), the small one upside down from its right foot, along its side, with
    // the other top corner on the arc. The small one is 3, the big one? Nothing to drag: the conditions
    // fix the shape (the small one is half as wide).
    [Puzzle("two-equilateral-triangles-in-a-semi-circle")]
    public static void TwoEquilateralTrianglesInASemiCircle(Lgf d)
    {
        const double R = 4;
        double a = R / Math.Sqrt(3), t = R / Math.Sqrt(3);
        var o = d.Pt(0, 0);
        var east = d.Pt(R, 0);
        var west = d.Pt(-R, 0);
        var apex = d.Pt(0, R);
        var left = d.Pt(-a, 0);
        var right = d.Pt(a, 0);
        var p = d.Pt(a - t / 2, t * 2 * S3 / 2);
        var q = d.Pt(a + t / 2, t * 2 * S3 / 2);
        d.Polygon(d.Fill("Green", "#FFDBFDDA", "#FF10A018", 3.5), left, right, apex);
        d.Polygon(d.Fill("Lime", "#FFF5FDA0", "#FFB4DC10", 3.5), right, q, p);
        string ink = d.Line("Ink", Lgf.Ink, 3.5);
        d.Arc("CircleArc", ink, o, east, west);
        d.Segment(ink, west, left, "TwoTicks");
        d.Segment(ink, left, right);
        d.Segment(ink, right, east, "TwoTicks");
        d.Label("?", V.At(0, R / 3), d.ShapeText("GreenMark", "#FF108A30", 40), 40);
        d.Label("3", Mid(d, right, p, q), d.ShapeText("LimeMark", "#FF8CB810", 26));
        d.Scene(-R, R, R, 0);
    }

    // Three squares in a circle: the big one with its left corners on it, the small one on its top right
    // corner and the middle one beside it, below, each with its outer top corner on the circle (which
    // makes the middle one twice the small one). The blue triangle is 5, the red one? Dragging the
    // middle square's bottom corner (the red triangle's) along its diagonal resizes the small squares
    // and the circle: the red triangle stays four times the blue.
    [Puzzle("three-squares-in-a-circle")]
    public static void ThreeSquaresInACircle(Lgf d)
    {
        const double a = 4;
        var m = d.Slider(V.At(a + 0.8, a - 0.8), V.At(a + 2.8, a - 2.8), 0.55, (a + 1.9, a - 1.9), "M");
        string b = d.Num($"(M.X - {Lgf.N(a)}) / 2", "b");
        string h = d.Num($"(({Lgf.N(a)} + {b}) ^ 2 + {Lgf.N(a)} * {b} + {b} ^ 2) / (2 * ({Lgf.N(a)} + {b}))", "hc");
        var o = d.Pt(0, 0);
        var tl = d.Pt(0, a);
        var tr = d.Pt(a, a);
        var center = d.Pt(new V(h, Lgf.N(a / 2)));
        var top = d.Pt(new V("0", $"{Lgf.N(a)} + {b}") + V.At(a, 0));
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        d.Box(frame, V.At(0, 0), V.At(a, a));
        d.Box(frame, V.At(a, a), new V($"{Lgf.N(a)} + {b}", $"{Lgf.N(a)} + {b}"));
        d.Box(frame, new V(Lgf.N(a), $"{Lgf.N(a)} - 2 * {b}"), new V($"{Lgf.N(a)} + 2 * {b}", Lgf.N(a)));
        d.Circle(d.Line("Circle", "#FF1010A8", 3.5), center, o);
        d.Inked(d.Outline("Cyan", "#FF10C4D8", 3), tl, tr, top);
        d.Inked(d.Outline("Red", "#FFD01818", 3), o, tr, m);
        d.Scene(-0.65, 5.6, 6.55, -1.6);
    }

    // A square, a quarter circle from its bottom right corner, the small square (4) in the top left
    // corner with its corner on the arc, and the circle in the bottom right corner touching the arc
    // there too; a square in the circle. Its area? Nothing to drag: it is all one shape.
    [Puzzle("square-in-a-circle-in-a-quarter-circle-in-a-square")]
    public static void SquareInACircleInAQuarterCircleInASquare(Lgf d)
    {
        var (center, r, _) = QuarterCircleSquareAndCircle(d, out var circle);
        var corners = new[] { V.At(r, 0), V.At(0, r), V.At(-r, 0), V.At(0, -r) }.Select(v => d.Pt(center + v)).ToArray();
        d.Polygon(d.Fill("Blue", "#FFCBEBF7", "#FF1A5EAD", 3.5), corners);
        d.Circle(circle, center, corners[0]);
        d.Label("?", center, d.ShapeText("BlueMark", "#FF1A5EAD", 44), 44);
    }

    // The same square, quarter circle, small square (4) and circle; the triangle in the circle between
    // the point where it touches the arc and the points where it touches the sides. Its area? Nothing
    // to drag: it is all one shape.
    [Puzzle("triangle-in-a-circle-in-a-quarter-circle-in-a-square")]
    public static void TriangleInACircleInAQuarterCircleInASquare(Lgf d)
    {
        var (center, r, touch) = QuarterCircleSquareAndCircle(d, out var circle);
        var bottom = d.Pt(center + V.At(0, -r));
        var right = d.Pt(center + V.At(r, 0));
        d.Polygon(d.Fill("Blue", "#FFBAF0FA", "#FF1898D0", 3.5), touch, bottom, right);
        d.Circle(circle, center, bottom);
        d.Label("?", Mid(d, touch, bottom, right), d.ShapeText("BlueMark", "#FF1A5EAD", 44), 44);
    }

    // An equilateral triangle (green, 3) and two arcs from the ends of its base, each through the middle
    // of the opposite side (its radius the triangle's height); the blue equilateral triangle stands on
    // the base with its top where the arcs cross. Its area? Nothing to drag: it is all one shape. Her
    // blue base lies on the green one and shows green there.
    [Puzzle("arcs-inside-a-triangle")]
    public static void ArcsInsideATriangle(Lgf d)
    {
        const double s = 6;
        double h = s * S3, top = s / Math.Sqrt(2), half = top / Math.Sqrt(3);
        var b = d.Pt(0, 0);
        var c = d.Pt(s, 0);
        var a = d.Pt(s / 2, h);
        var x = d.Pt(s / 2, top);
        var p = d.Pt(s / 2 - half, 0);
        var q = d.Pt(s / 2 + half, 0);
        d.Polygon(d.Outline("Green", "#FFA8D838", 3.5), b, c, a);
        string blue = d.Line("Blue", "#FF1EBFEA", 3.5);
        d.Segment(blue, p, x);
        d.Segment(blue, x, q);
        d.Segment(d.Line("Base", "#FF10B030", 3.5), p, q);
        // the arcs run a little past their crossing, as hers do
        double cross = Math.Atan(Math.Sqrt(2)) * 180 / Math.PI + 6;
        string ink = d.Line("Ink", Lgf.Ink, 3.5);
        d.Arc("CircleArc", ink, b, d.Pt(h, 0), d.Pt(h * Math.Cos(cross * Math.PI / 180), h * Math.Sin(cross * Math.PI / 180)));
        d.Arc("CircleArc", ink, c, d.Pt(s - h * Math.Cos(cross * Math.PI / 180), h * Math.Sin(cross * Math.PI / 180)), d.Pt(s - h, 0));
        d.Scene(0, h, s, 0);
    }

    // A square of side 12 and two tilted squares in it: the big one with corners on the right and bottom
    // sides, the small one with corners on the top and left sides, standing on the big one's side. The
    // quadrilateral through the four dots? Dragging the big square's corner up and down the right side
    // resizes both (being a square puts the small one's left corner 4 above the big one's right one): the
    // quadrilateral, not drawn, stays 128.
    [Puzzle("quadrilateral-from-two-squares-in-a-square")]
    public static void QuadrilateralFromTwoSquaresInASquare(Lgf d)
    {
        const double L = 12;
        var r = d.Slider(V.At(L, 2.5), V.At(L, 5.8), (5.3 - 2.5) / 3.3, (L, 5.3), "R");
        string y = "R.Y", w = $"(R.Y + {Lgf.N(L / 3)})";
        var b = d.Pt(new V($"{Lgf.N(L)} - {y}", "0"), "B");
        var bigTop = d.Pt(new V($"{Lgf.N(L)} - {y}", $"2 * {y}"));
        var bigLeft = d.Pt(new V($"{Lgf.N(L)} - 2 * {y}", y));
        var t = d.Pt(new V($"{Lgf.N(L)} - {w}", Lgf.N(L)), "T");
        var west = d.Pt(new V("0", w), "W");
        string u = $"({Lgf.N(L)} - {w})";
        var foot = d.Pt(west + new V(u, V.Neg(u)));
        var side = d.Pt(t + new V(u, V.Neg(u)));
        d.Box(d.Outline("Frame", "#FFB848B0", 3.5), V.At(0, 0), V.At(L, L));
        d.Polygon(d.Outline("Cyan", "#FF14C4D0", 3.5), west, t, side, foot);
        d.Polygon(d.Outline("Blue", "#FF0A4CC4", 3.5), b, r, bigTop, bigLeft);
        // the fourth dot is the handle
        foreach (var p in new[] { t, west, b })
        {
            d.DotAt(p, size: 12);
        }

        d.Dimension(d.Line("Arrow", "#FF2A2A9A", 3), d.Text("NavyMark", "#FF2A2A9A", 30), d.Pt(0, L), d.Pt(L, L), 1.1, "12", 1.2, 30);
        d.Scene(0, L + 1.6, L, 0);
    }

    // A square on its corner (diagonal 6) with equilateral triangles on its two upper sides, outside it;
    // the red square has its corners at the centers of the triangles on all four sides (the lower two
    // not drawn). The green parts, the red square's corners in the triangles? Nothing to drag: the 6
    // fixes it.
    [Puzzle("triangle-centres-and-square-corners")]
    public static void TriangleCentresAndSquareCorners(Lgf d)
    {
        const double D = 3;
        double apex = D / 2 + D / 2 * Math.Sqrt(3), k = (D + apex) / 3;
        var left = d.Pt(-D, 0);
        var top = d.Pt(0, D);
        var right = d.Pt(D, 0);
        var bottom = d.Pt(0, -D);
        var leftApex = d.Pt(-apex, apex);
        var rightApex = d.Pt(apex, apex);
        var red = new[] { d.Pt(-k, -k), d.Pt(k, -k), d.Pt(k, k), d.Pt(-k, k) };
        string square = d.Polygon(null, red);
        string leftTriangle = d.Polygon(null, left, top, leftApex);
        string rightTriangle = d.Polygon(null, right, top, rightApex);
        string green = d.Fill("Green", "#FFBEC898", null);
        d.Region(green, "Intersection", square, leftTriangle);
        d.Region(green, "Intersection", square, rightTriangle);
        string orange = d.Outline("Orange", "#FFFCA000", 3.5);
        d.Polygon(orange, left, top, leftApex);
        d.Polygon(orange, right, top, rightApex);
        d.Polygon(d.Outline("Red", "#FFC03A08", 3.5), red);
        d.Polygon(d.Outline("Frame", Lgf.Ink, 4), left, top, right, bottom);
        d.Dimension(d.Line("Arrow", Lgf.Ink, 3), d.Text("Mark", Lgf.Ink, 36), left, right, 0, "6", 0.6, 36);
        d.Scene(-apex, apex, apex, -D);
    }

    // A regular hexagon of side 3 and three equally spaced circles around its top left corner, the
    // biggest through its two neighbors: shaded the middle and the outer ring's parts outside the
    // hexagon and the middle ring's part inside it. The total? Nothing to drag: the side fixes it.
    [Puzzle("hexagon-and-concentric-circles")]
    public static void HexagonAndConcentricCircles(Lgf d)
    {
        const double s = 3;
        var hexagon = HexagonUp(d, V.At(0, -2 * S3 * s), Lgf.N(s));
        // bottom left, bottom right, right, top right, top left (the center), left
        var o = hexagon[4];
        var c1 = d.Pt(1, 0);
        var c2 = d.Pt(2, 0);
        string inner = d.Circle(null, o, c1);
        string middle = d.Circle(null, o, c2);
        string outer = d.Circle(null, o, hexagon[3]);
        string shape = d.Polygon(null, hexagon);
        string green = d.Fill("Green", "#FFF0FEB8", null);
        d.Region(green, "Difference", inner, shape);
        d.Region(green, "Difference", outer, middle, shape);
        d.Region(green, "Difference", d.Region(null, "Intersection", middle, shape), inner);
        // across the pale green, which stays bright in the dark theme: a mid green there
        string line = d.Line("Circle", "#FF129418", 3.5, dark: "#FF3DAA3F");
        d.Circle(line, o, c1);
        d.Circle(line, o, c2);
        d.Circle(line, o, hexagon[3]);
        d.Inked(d.Outline("Frame", Lgf.Ink, 3.5), hexagon);
        d.Dimension(d.Line("Arrow", "#FF108A18", 2.5), d.Text("GreenMark", "#FF108A18", 30), hexagon[0], hexagon[1], 0.4, "3", 0.32, 30);
        d.Scene(-s, s, 1.5 * s, -2 * S3 * s);
    }

    // An annulus and a chord of the outer circle cut into three equal parts (2) by the inner one. The
    // annulus's area? Dragging the circles' center to and from the chord changes both circles (the
    // chord and its dots stay): the area stays 8π.
    [Puzzle("chord-and-annulus")]
    public static void ChordAndAnnulus(Lgf d)
    {
        double angle = 22 * Math.PI / 180, cos = Math.Cos(angle), sin = Math.Sin(angle);
        var a = d.Pt(-3 * cos, -3 * sin);
        var b = d.Pt(-cos, -sin);
        var c = d.Pt(cos, sin);
        var e = d.Pt(3 * cos, 3 * sin);
        var o = d.Slider(V.At(0.3 * sin, -0.3 * cos), V.At(1.35 * sin, -1.35 * cos), (0.94 - 0.3) / 1.05, (0.94 * sin, -0.94 * cos), "O");
        string outer = d.Circle(null, o, a);
        string inner = d.Circle(null, o, b);
        d.Region(d.Fill("Pink", "#FFFEC8E8", null), "Difference", outer, inner);
        string magenta = d.Line("Magenta", "#FFE41476", 3.5);
        d.Circle(magenta, o, a);
        d.Circle(magenta, o, b);
        d.Segment(d.Line("Chord", Lgf.Ink, 4), a, e);
        foreach (var p in new[] { a, b, c, e })
        {
            d.DotAt(p, size: 12);
        }

        d.Dimension(d.Line("Arrow", Lgf.Ink, 2.5), d.Text("Mark", Lgf.Ink, 30), b, c, -0.38, "2", 0.32, 30);
        d.Scene(-2.8, 2.3, 3.5, -4.2);
    }

    // A regular hexagon and two isosceles triangles in it, each on one of its flat sides with the apex in
    // the middle of the opposite side; the pink parts are what neither covers. What fraction? Nothing to
    // drag: the hexagon fixes it.
    [Puzzle("dissected-hexagon")]
    public static void DissectedHexagon(Lgf d)
    {
        const double s = 2;
        var h = HexagonFromLeft(d, V.At(0, 0), Lgf.N(s));
        // left, top left, top right, right, bottom right, bottom left
        var top = Mid(d, h[1], h[2]);
        var bottom = Mid(d, h[4], h[5]);
        var left = d.Crossing(h[1], bottom, h[5], top);
        var right = d.Crossing(h[2], bottom, h[4], top);
        string pink = d.Fill("Pink", "#FFFECCF6", null);
        d.Polygon(pink, h[1], h[0], h[5], left);
        d.Polygon(pink, h[2], h[3], h[4], right);
        d.Polygon(d.Outline("Frame", Lgf.Ink, 3.5), h);
        d.Inked(d.Outline("Navy", "#FF1A14B0", 3.5), h[1], h[2], bottom);
        d.Inked(d.Outline("Cyan", "#FF30C8EC", 3.5), h[5], h[4], top);
        d.Scene(0, S3 * s, 2 * s, -S3 * s);
    }

    /// <summary>
    /// The square (side 6) of the two quarter circle puzzles with its quarter circle from the bottom right
    /// corner and the small square (4) in the top left corner: the circle's center and radius, the point
    /// where the circle, the arc and the small square meet, and the circle's style
    /// </summary>
    static (V Center, double Radius, V Touch) QuarterCircleSquareAndCircle(Lgf d, out string circleStyle)
    {
        const double L = 6;
        double s = L * (1 - 1 / Math.Sqrt(2)), r = L * (Math.Sqrt(2) - 1);
        var corner = d.Pt(L, 0);
        var top = d.Pt(L, L);
        var west = d.Pt(0, 0);
        d.Box(d.Outline("Frame", Lgf.Ink, 3.5), V.At(0, 0), V.At(L, L));
        d.Box(d.Fill("Cyan", "#FFE2F8F9", "#FF10B0B8", 3.5), V.At(0, L - s), V.At(s, L));
        d.Arc("CircleArc", d.Line("Arc", "#FFB040A0", 3.5), corner, top, west);
        d.Label("4", V.At(s / 2, L - s / 2), d.ShapeText("CyanMark", "#FF10A0A8", 36), 36);
        circleStyle = d.Line("Circle", "#FF1A1090", 3.5);
        d.Scene(0, L, L, 0);
        return (d.Pt(L - r, r), r, d.Pt(s, L - s));
    }
}
