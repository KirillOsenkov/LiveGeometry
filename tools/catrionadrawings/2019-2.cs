#nullable disable

static partial class Puzzles
{
    // A rectangle with a circle touching its left, top and bottom sides, and a square with its right
    // side on the circle and its left corners on it; the rectangle's right side runs through where the
    // square's top and bottom cross the circle again. The square's part inside the rectangle is shaded:
    // what fraction of the total area? Nothing to drag: that fixes it (for a radius of 5 the square's
    // side is 8, its left corners at (-3, 4) and (-3, -4), the rectangle's right side at 3).
    [Puzzle("a-square-a-circle-and-a-rectangle")]
    public static void ASquareACircleAndARectangle(Lgf d)
    {
        const double R = 5;
        var o = d.Pt(0, 0, "O");
        d.Box(d.Fill("Green", "#FFDCF47E", null), V.At(-3, -4), V.At(3, 4));
        d.Box(d.Outline("SquareLine", "#FF9ACC14", 3.5), V.At(-3, -4), V.At(R, 4));
        d.Box(d.Outline("RectangleLine", "#FF10C87E", 3.5), V.At(-R, -R), V.At(3, R));
        d.Circle(d.Outline("Circle", Lgf.Ink, 3.5), o, d.Pt(R, 0));
        d.Scene(-R, R, R, -R);
    }

    // A right triangle; two points on its hypotenuse as far from either end (the red lengths), lines
    // from the right angle to both and down from both to the base: three right triangles. The orange
    // part is 7, the pink? With the points a fraction t along from either end, both are
    // w h t (1 - 2t) / (2 (1 - t)). The top corner is dragged up and down, the upper point along the
    // hypotenuse: the triangles change, pink stays orange.
    [Puzzle("three-right-angled-triangles")]
    public static void ThreeRightAngledTriangles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double W = 8.1;
        var o = d.Pt(0, 0, "O");
        var c = d.Pt(W, 0, "C");
        var a = d.Slider(V.At(0, 4), V.At(0, 9), 0.4, (0, 6), "A");
        var p = d.Slider(a + 0.15 * (c - a), a + 0.42 * (c - a), (0.293 - 0.15) / 0.27, (0.293 * W, 0.707 * 6), "P");
        var q = d.Pt(a + c - p, "Q");
        var footP = d.Pt(new V("P.X", "0"), "Fp");
        var footQ = d.Pt(new V("Q.X", "0"), "Fq");
        var x = d.Pt(new V("P.X", "Q.Y * P.X / Q.X"), "X");
        d.Polygon(d.Fill("Pink", "#FFEC6E9F", null), o, p, x);
        d.Polygon(d.Fill("Orange", "#FFF6AE03", null), x, footP, footQ, q);
        string magenta = d.Line("Magenta", "#FFD72587", 3);
        d.Segment(magenta, o, p);
        d.Segment(magenta, p, footP);
        string orange = d.Line("OrangeLine", "#FFC83A11", 3);
        d.Segment(orange, o, q);
        d.Segment(orange, q, footQ);
        string red = d.Line("Red", "#FFA80001", 3.5);
        d.Segment(red, a, p, "TwoTicks");
        d.Segment(red, q, c, "TwoTicks");
        string ink = d.Line("Ink", Lgf.Ink, 5);
        d.Segment(ink, p, q);
        d.Segment(ink, o, a);
        d.Segment(ink, o, c);
        string mark = d.ShapeText("ShapeMark", Lgf.Ink, 40);
        d.Label("?", Mid(d, o, p, x), mark, 40);
        d.Label("7", Mid(d, x, footP, footQ, q), mark, 40);
        d.Scene(0, 6, W, 0);
    }

    // A red rectangle 12 wide and a green one side by side on a line, and a yellow square standing on the
    // corner they share, its left corner on the red one's top, its right corner on the green one's right
    // side, its top left side through the green one's top left corner. The square is twice the red
    // rectangle: the green one's height? The square's left corner is dragged along the circle it keeps
    // to (the square's side s = 24 cos t at a tilt t, so the corner is at (-12 sin 2t, 12 + 12 cos 2t)):
    // the square tilts and the rectangles change, the green one stays 24 tall.
    [Puzzle("two-rectangles-overlapping-a-square")]
    public static void TwoRectanglesOverlappingASquare(Lgf d)
    {
        d.Texture = "FeltTip";
        var b = d.Pt(0, 0, "B");
        var center = d.Pt(0, 12, "K");
        double at = 138 * Math.PI / 180;
        var l = d.ArcSlider(center, 12, 112, 166, 138, (0, 12), "L");
        var r = d.Pt(new V("L.Y", "-L.X"), "R");
        var t = d.Pt(new V("L.X + L.Y", "L.Y - L.X"), "T");
        d.Polygon(d.Fill("Yellow", "#FFFFFF01", null), b, r, t, l);
        string redColor = "#FFF41E3C", greenColor = "#FF05CF88";
        // across the yellow, which stays bright in the dark theme: mid-tones there
        string red = d.Outline("RedLine", redColor, 3.5, dark: "#FFE8405A");
        string green = d.Outline("GreenLine", greenColor, 3.5, dark: "#FF20B878");
        d.Box(red, V.At(-12, 0), new V("0", "L.Y"));
        d.Box(green, V.At(0, 0), new V("L.Y", "24"));
        var foot = d.Pt(new V("L.Y", "0"));
        var head = d.Pt(new V("L.Y", "24"));
        d.Dimension(d.Line("RedArrow", redColor, 3), d.Text("RedMark", redColor, 30), b, d.Pt(-12, 0), 2.2, "12", 1.6);
        d.Dimension(d.Line("GreenArrow", greenColor, 3), d.Text("GreenMark", greenColor, 30), head, foot, 2.2, "?", 1.6);
        double lx = 12 * Math.Cos(at), ly = 12 + 12 * Math.Sin(at);
        d.Scene(-12, ly - lx, ly + 2.8, -2.8);
    }

    // A trapezium and its diagonals: blue the side triangles, purple the bottom one, the top one white.
    // Blue equals purple when the bottom side is twice the top one (the side triangles are k times the
    // top one, the bottom one k^2 times, for a ratio k of the sides): what fraction is shaded? The top
    // left corner is dragged anywhere in a box, the top side stays half the bottom one: the fraction
    // stays.
    [Puzzle("areas-in-a-crossed-trapezium")]
    public static void AreasInACrossedTrapezium(Lgf d)
    {
        d.Texture = "FeltTip";
        const double B = 9.2;
        var a = d.Pt(0, 0, "A");
        var b = d.Pt(B, 0, "Bt");
        string box = d.Polygon(null, d.Pt(-1.5, 2.8), d.Pt(5.6, 2.8), d.Pt(5.6, 7), d.Pt(-1.5, 7));
        var c = d.HandleIn(box, 1.26, 5, "C");
        var e = d.Pt(c + V.At(B / 2, 0), "E");
        var x = d.Crossing(a, e, b, c, "X");
        string blue = d.Fill("Blue", "#FF2AEFEA", null);
        d.Polygon(blue, a, x, c);
        d.Polygon(blue, b, e, x);
        d.Polygon(d.Fill("Purple", "#FF9418B8", null), a, b, x);
        string ink = d.Line("Ink", Lgf.Ink, 5);
        d.Segment(ink, a, e);
        d.Segment(ink, b, c);
        d.Polygon(d.Outline("Frame", Lgf.Ink, 5), a, b, e, c);
        d.Scene(0, 5, B, 0);
    }

    // Seven regular hexagons in a flower, a circle in each, and a circle through the flower's outer
    // corners. What fraction is shaded? Nothing to drag: the hexagons fix it (for a side of 1 the outer
    // circle's radius is sqrt 7, the small circles' sqrt 3 / 2).
    [Puzzle("seven-hexagons")]
    public static void SevenHexagons(Lgf d)
    {
        string purple = "#FF9A6FD0";
        var o = d.Pt(0, 0, "O");
        d.Circle(d.Outline("Circle", purple, 3), o, d.Pt(Math.Sqrt(7), 0));
        string blue = d.Fill("Blue", "#FF4ACFE6", purple, 2.5);
        string frame = d.Outline("Frame", "#FF594994", 4);
        var centers = new List<(double X, double Y)> { (0, 0) };
        for (int k = 0; k < 6; k++)
        {
            double angle = (30 + 60 * k) * Math.PI / 180;
            centers.Add((2 * S3 * Math.Cos(angle), 2 * S3 * Math.Sin(angle)));
        }

        foreach (var (x, y) in centers)
        {
            var center = d.Pt(x, y);
            d.Circle(blue, center, d.Pt(x, y + S3));
            d.Inked(frame, Enumerable.Range(0, 6).Select(i => d.Pt(x + Math.Cos(i * Math.PI / 3), y + Math.Sin(i * Math.PI / 3))).ToArray());
        }

        double r = Math.Sqrt(7);
        d.Scene(-r, r, r, -r);
    }

    // A square with a square in each corner and a red square on its corner in the middle, an orange
    // square on its corner inside that. All six colored areas are equal: what fraction is shaded? The
    // orange square is half the red one; the red one's sides run through the corner squares' inner
    // corners. Nothing to drag: that fixes it (corner squares of side 1 in a square of side 3).
    [Puzzle("a-pattern-of-squares")]
    public static void APatternOfSquares(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 3, h = 1.5;
        d.Box(d.PaperFill("Paper", Lgf.Ink, 5), V.At(0, 0), V.At(L, L));
        d.Box(d.Fill("Blue", "#FF30EFF8", Lgf.Ink, 5), V.At(0, 2), V.At(1, 3));
        d.Box(d.Fill("Yellow", "#FFFFFB01", Lgf.Ink, 5), V.At(2, 2), V.At(3, 3));
        d.Box(d.Fill("Green", "#FFB7ED03", Lgf.Ink, 5), V.At(0, 0), V.At(1, 1));
        d.Box(d.Fill("Purple", "#FF4124F8", Lgf.Ink, 5), V.At(2, 0), V.At(3, 1));
        double k = 1 / Math.Sqrt(2);
        d.Polygon(d.Fill("Red", "#FFFD0302", Lgf.Ink, 5), d.Pt(h + 1, h), d.Pt(h, h + 1), d.Pt(h - 1, h), d.Pt(h, h - 1));
        d.Polygon(d.Fill("Orange", "#FFFE7202", Lgf.Ink, 5), d.Pt(h + k, h), d.Pt(h, h + k), d.Pt(h - k, h), d.Pt(h, h - k));
        d.Scene(0, L, L, 0);
    }

    // A big regular hexagon (outlined in yellow), a rectangle across its lower part from its left corner
    // to its right corner, and two shaded regular hexagons in the rectangle: a small one in the big one's
    // right corner, up against the rectangle's top, and a middle one standing on the bottom, reaching the
    // top, its right corner the small one's lower left one. What fraction of the rectangle is shaded?
    // Nothing to drag: that fixes them (sides 1/3 and 2/3 of the big one's).
    [Puzzle("three-hexagons-and-a-rectangle")]
    public static void ThreeHexagonsAndARectangle(Lgf d)
    {
        const double S = 3, t = S / 3, m = 2 * S / 3;
        var big = HexagonFromLeft(d, V.At(-S, 0), Lgf.N(S));
        d.Polygon(d.Outline("Yellow", "#FFBCDF10", 3), big);
        string green = d.Fill("Green", "#FFC8FCD8", "#FF10E890", 3.5);
        d.Polygon(green, HexagonFromLeft(d, V.At(S - 2 * t, 0), Lgf.N(t)));
        d.Polygon(green, HexagonFromLeft(d, V.At(S / 2 - 2 * m, -S3 * S + S3 * m), Lgf.N(m)));
        double top = S3 * t, bottom = -S3 * S;
        d.Box(d.Outline("Frame", Lgf.Ink, 4.5), V.At(-S, bottom), V.At(S, top));
        d.Scene(-S, S3 * S, S, bottom);
    }

    // A square, its inscribed circle, six touching circles of radius 1 in two rows of three inside it,
    // the corner ones touching it, and red lines across the square along the rows' tops and bottoms.
    // The shaded strips' total area? Nothing to drag: that fixes it (the big circle's radius is
    // 1 + sqrt 5).
    [Puzzle("six-circles-inside-a-circle-inside-a-square")]
    public static void SixCirclesInsideACircleInsideASquare(Lgf d)
    {
        d.Texture = "FeltTip";
        double R = 1 + Math.Sqrt(5);
        string orange = d.Fill("Orange", "#FFFEC002", null);
        d.Box(orange, V.At(-R, 2), V.At(R, R));
        d.Box(orange, V.At(-R, -R), V.At(R, -2));
        d.Box(d.Outline("Frame", Lgf.Ink, 4.5), V.At(-R, -R), V.At(R, R));
        // the circle crosses the orange, which stays bright in the dark theme: a mid-tone blue there
        d.Circle(d.Outline("Circle", "#FF26028B", 3, dark: "#FF6B7BFF"), d.Pt(0, 0, "O"), d.Pt(R, 0));
        string purple = d.Outline("Small", "#FF97017E", 3.5);
        foreach (int y in new[] { -1, 1 })
        {
            foreach (int x in new[] { -2, 0, 2 })
            {
                d.Circle(purple, d.Pt(x, y), d.Pt(x + 1, y));
            }
        }

        string red = d.Line("Red", "#FFF52A04", 3.5);
        d.Segment(red, d.Pt(-R, 2), d.Pt(R, 2));
        d.Segment(red, d.Pt(-R, -2), d.Pt(R, -2));
        d.Scene(-R, R, R, -R);
    }

    // Two squares side by side, three equal purple squares in each: in the left one stacked up its
    // middle, one on its corner between two upright ones, touching both sides; in the right one, of
    // side 12, two in the bottom corners and one on its corner above them, touching the top, its lower
    // sides through their inner top corners. The left one's side? Nothing to drag: the right square
    // fixes the purple ones (side 18 / (2 + sqrt 2)), and they fix the left one (18).
    [Puzzle("stacked-squares-inside-two-squares")]
    public static void StackedSquaresInsideTwoSquares(Lgf d)
    {
        double a = 18 / (2 + Math.Sqrt(2)), h = a / Math.Sqrt(2);
        string frame = d.Outline("Frame", Lgf.Ink, 4.5);
        d.Box(frame, V.At(0, 0), V.At(18, 18));
        d.Box(frame, V.At(18, 0), V.At(30, 12));
        string purple = d.Fill("Purple", "#FFDDD8F7", "#FF2014A8", 3);
        V[] Diamond(double x, double y) => [d.Pt(x + h, y), d.Pt(x, y + h), d.Pt(x - h, y), d.Pt(x, y - h)];
        d.Box(purple, V.At(9 - a / 2, 0), V.At(9 + a / 2, a));
        d.Polygon(purple, Diamond(9, a + h));
        d.Box(purple, V.At(9 - a / 2, 18 - a), V.At(9 + a / 2, 18));
        d.Box(purple, V.At(18, 0), V.At(18 + a, a));
        d.Box(purple, V.At(30 - a, 0), V.At(30, a));
        d.Polygon(purple, Diamond(24, 12 - h));
        string arrow = d.Line("Arrow", Lgf.Ink, 3.5);
        string mark = d.Text("Mark", Lgf.Ink, 30);
        d.Dimension(arrow, mark, d.Pt(0, 18), d.Pt(18, 18), 1.5, "?", 1);
        d.Dimension(arrow, mark, d.Pt(18, 12), d.Pt(30, 12), 1.5, "12", 1.4);
        d.Scene(0, 20.2, 30, 0);
    }

    // A right triangle and three squares in it, each with a corner on the hypotenuse: a big white one in
    // the right angle, a small green one on top of it and a middle green one beside it. Their corners on
    // the hypotenuse are 6 apart: the green squares' total area? For green sides a and b the white one's
    // is a + b, and its corner on the hypotenuse is the middle of the 6, which runs (2b, -2a): a^2 + b^2
    // is 9. The middle square's corner is dragged around a circle of radius 3 about the white one's
    // corner: the triangle changes shape, the green total stays.
    [Puzzle("three-squares-inside-a-triangle")]
    public static void ThreeSquaresInsideATriangle(Lgf d)
    {
        double home = -29 * Math.PI / 180, a0 = -3 * Math.Sin(home), b0 = 3 * Math.Cos(home), s0 = a0 + b0;
        var m = d.Pt(s0, s0, "M");
        var pb = d.ArcSlider(m, 3, 323, 337, 331, (s0, s0), "Pb");
        string b = d.Num("Pb.X - M.X", "b");
        string a = d.Num("M.Y - Pb.Y", "a");
        string s = d.Num($"{a} + {b}", "s");
        var o = d.Pt(new V($"M.X - {s}", $"M.Y - {s}"), "O");
        var pa = d.Pt(new V("2 * M.X - Pb.X", "2 * M.Y - Pb.Y"), "Pa");
        var t = d.Pt(new V("O.X", $"O.Y + {s} + {s} * {a} / {b}"), "T");
        var c = d.Pt(new V($"O.X + {s} + {s} * {b} / {a}", "O.Y"), "C");
        string greenLine = "#FFB4E030", blueColor = "#FF1A3CDB";
        d.Box(d.Outline("White", greenLine, 3), o, m);
        string green = d.Fill("Green", "#FFD6FC68", greenLine, 3);
        d.Box(green, new V("O.X", "M.Y"), pa);
        d.Box(green, new V("M.X", "O.Y"), pb);
        d.Polygon(d.Outline("Triangle", blueColor, 3.5), o, c, t);
        d.DotAt(pa, "Da", 10);
        d.DotAt(pb, "Db", 10);
        var normal = new V($"{a} / 3", $"{b} / 3");
        string arrow = d.Line("Arrow", blueColor, 3);
        var middle = d.Pt(m + 0.42 * normal);
        d.Vector(arrow, middle, d.Pt(pa + 0.42 * normal));
        d.Vector(arrow, middle, d.Pt(pb + 0.42 * normal));
        d.Label("6", m + 1.2 * normal, d.Text("BlueMark", blueColor, 32));
        d.Scene(0, s0 + s0 * a0 / b0, s0 + s0 * b0 / a0, 0);
    }

    // Three congruent right triangles: two with their right angles at one point on a horizontal line,
    // one along it to the left with its short leg down, one with its short leg along it to the right
    // and its long leg down; the third has its hypotenuse on the line, ending at the second one's far
    // corner. The line bisecting its top (right) angle runs to the first one's bottom corner. The
    // angle it makes with the horizontal? Nothing to drag: the bisector through that corner makes the
    // triangles 30-60-90 (hypotenuse twice the short leg).
    [Puzzle("three-congruent-triangles")]
    public static void ThreeCongruentTriangles(Lgf d)
    {
        double p = 2, q = 2 * Math.Sqrt(3), r = 4;
        var c = d.Pt(0, 0, "C");
        var a = d.Pt(-q, 0, "A");
        var b = d.Pt(p - r, 0, "B");
        var dd = d.Pt(p, 0, "D");
        var e = d.Pt(0, -p, "E");
        var f = d.Pt(0, -q, "F");
        var t = d.Pt(p - p * p / r, p * q / r, "T");
        string top = d.Polygon(null, t, b, dd);
        string left = d.Polygon(null, a, c, e);
        string bottom = d.Polygon(null, c, dd, f);
        d.Region(d.Fill("Mint", "#FFCCFEEE", null), "Union", top, left, bottom);
        string green = d.Outline("Green", "#FF10C882", 3);
        d.Polygon(green, t, b, dd);
        d.Polygon(green, a, c, e);
        d.Polygon(green, c, dd, f);
        string greenLine = d.Line("GreenMark", "#FF10C882", 2.5);
        d.ArcMark(greenLine, t, b, e, 0.34);
        d.ArcMark(greenLine, t, b, e, 0.44);
        d.ArcMark(greenLine, t, e, dd, 0.34);
        d.ArcMark(greenLine, t, e, dd, 0.44);
        string ink = d.Line("Ink", Lgf.Ink, 4);
        d.Segment(ink, t, e);
        var x = d.Crossing(t, e, a, dd, "X");
        d.ArcMark(d.Line("AngleInk", Lgf.Ink, 3), x, e, dd, 0.32);
        d.Label("?", x + V.At(0.28, -0.85), d.Text("Mark", Lgf.Ink, 32), 32);
        d.Scene(-q, p * q / r, p, -q);
    }

    // A semicircle; in it a pink semicircle of diameter 2 standing on the diameter at its left end, and a
    // yellow one of diameter 3 hanging upside down, its arc touching the diameter and the pink one, its
    // right corner on the big arc. The big one's diameter? Nothing to drag: that fixes it (5).
    [Puzzle("two-semi-circles-inside-a-semi-circle-ii")]
    public static void TwoSemiCirclesInsideASemiCircleII(Lgf d)
    {
        const double R = 2.5;
        string red = "#FFD80204", orange = "#FFF49003";
        // she colored the pink one in pencil and the yellow one in felt tip
        string pink = d.Fill("Pink", "#FFFDCFB9", red, 3);
        d.Texture = "FeltTip";
        string yellow = d.Fill("Yellow", "#FFFFFF18", orange, 3);
        d.Arc("CircleSegment", pink, d.Pt(-1.5, 0), d.Pt(-0.5, 0), d.Pt(-2.5, 0));
        d.Arc("CircleSegment", yellow, d.Pt(0.5, 1.5), d.Pt(-1, 1.5), d.Pt(2, 1.5));
        var east = d.Pt(R, 0, "E");
        var west = d.Pt(-R, 0, "W");
        d.Arc("CircleSegment", d.Outline("Frame", Lgf.Ink, 3.5), d.Pt(0, 0, "O"), east, west);
        d.Dimension(d.Line("RedArrow", red, 2.5, dark: "#FF8A0A0A"), d.ShapeText("RedMark", red, 28), d.Pt(-2.5, 0), d.Pt(-0.5, 0), 0.13, "2", 0.28);
        d.Dimension(d.Line("OrangeArrow", orange, 2.5), d.Text("OrangeMark", orange, 28), d.Pt(-1, 1.5), d.Pt(2, 1.5), 0.18, "3", 0.28);
        d.Dimension(d.Line("Arrow", Lgf.Ink, 3.5), d.Text("Mark", Lgf.Ink, 30), east, west, 0.3, "?", 0.28);
        d.Scene(-R, R, R, -0.5);
    }

    // Three semicircles on a line from one end, two white ones filling the gaps between them: purple the
    // smallest, blue the middle one less it and the first gap, yellow the biggest less the middle one and
    // the second gap. The three colored areas are equal: what fraction do they cover? Nothing to drag:
    // that fixes the diameters (6, 9 and 11).
    [Puzzle("multiple-semi-circles")]
    public static void MultipleSemiCircles(Lgf d)
    {
        const double P = 6, M = 9, D = 11;
        var left = d.Pt(0, 0, "A");
        var p = d.Pt(P, 0, "P");
        var m = d.Pt(M, 0, "M");
        var e = d.Pt(D, 0, "E");
        string Half(double from, double to, V start, V end) => d.Arc("CircleSegment", null, d.Pt((from + to) / 2, 0), end, start, visible: false);
        string big = Half(0, D, left, e);
        string middle = Half(0, M, left, m);
        string small = Half(0, P, left, p);
        string gap1 = Half(P, M, p, m);
        string gap2 = Half(M, D, m, e);
        d.Region(d.Fill("Yellow", "#FFECFE68", null), "Difference", big, middle, gap2);
        d.Region(d.Fill("Blue", "#FF5CE2F9", null), "Difference", middle, small, gap1);
        d.Arc("CircleSegment", d.Fill("Purple", "#FFC990EA", null), d.Pt(P / 2, 0), p, left);
        string ink = d.Line("Ink", Lgf.Ink, 5);
        foreach (var (from, to, start, end) in new[] { (0.0, D, left, e), (0.0, M, left, m), (0.0, P, left, p), (P, M, p, m), (M, D, m, e) })
        {
            d.Arc("CircleArc", ink, d.Pt((from + to) / 2, 0), end, start);
        }

        d.Segment(ink, left, e);
        d.Scene(0, D / 2, D, 0);
    }
}
