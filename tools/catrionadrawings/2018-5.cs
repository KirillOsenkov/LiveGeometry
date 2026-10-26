#nullable disable

static partial class Puzzles
{
    /// <summary>The far corners of the square standing on the side from p to q, on its left (turned counterclockwise from p to q): the one over q, then the one over p</summary>
    static (V OverQ, V OverP) SquareOnLeftOfSide(Lgf d, V p, V q)
    {
        var n = (q - p).Turn(90);
        return (d.Pt(q + n), d.Pt(p + n));
    }

    // A quadrilateral of area 10 with a square on each side, outward (outlined in green, blue, red and
    // cyan); between each two neighboring squares a gray triangle closed by a black line. The total
    // gray area? Dragging two opposite corners of the quadrilateral changes its shape: each gray
    // triangle stays as big as the triangle of the quadrilateral's two sides at that corner, so the
    // gray stays twice the quadrilateral.
    [Puzzle("four-squares-around-a-quadrilateral")]
    public static void FourSquaresAroundAQuadrilateral(Lgf d)
    {
        d.Texture = "FeltTip";
        var a = d.HandleIn(d.Circle(null, d.Pt(-1.5, 1.6), d.Pt(-1.1, 1.6)), -1.5, 1.6, "A");
        var b = d.Pt(1.1, 2.0, "B");
        var c = d.HandleIn(d.Circle(null, d.Pt(1.56, -0.64), d.Pt(1.96, -0.64)), 1.56, -0.64, "C");
        var e = d.Pt(-1.56, -1.64, "D");
        // the corners go clockwise: the squares stand on the left of each side
        var (abB, abA) = SquareOnLeftOfSide(d, a, b);
        var (bcC, bcB) = SquareOnLeftOfSide(d, b, c);
        var (ceE, ceC) = SquareOnLeftOfSide(d, c, e);
        var (eaA, eaE) = SquareOnLeftOfSide(d, e, a);
        string gray = d.Fill("Gray", "#FFBFC5B8", null);
        d.Polygon(gray, a, eaA, abA);
        d.Polygon(gray, b, abB, bcB);
        d.Polygon(gray, c, bcC, ceC);
        d.Polygon(gray, e, ceE, eaE);
        d.Polygon(d.Outline("Green", "#FFA8DC48", 3.5), a, b, abB, abA);
        d.Polygon(d.Outline("Blue", "#FF1382B0", 3.5), b, c, bcC, bcB);
        d.Polygon(d.Outline("Red", "#FFC8461E", 3.5), c, e, ceE, ceC);
        d.Polygon(d.Outline("Cyan", "#FF45A8B8", 3.5), e, a, eaA, eaE);
        string line = d.Line("GrayLine", "#FF4A4C48", 3);
        d.Segment(line, eaA, abA);
        d.Segment(line, abB, bcB);
        d.Segment(line, bcC, ceC);
        d.Segment(line, ceE, eaE);
        d.Label("10", Mid(d, a, b, c, e), d.Text("Mark", Lgf.Ink, 30), 30);
        d.Scene(-5.1, 4.9, 4.3, -4.7);
    }

    // Two equilateral triangles of the same size sharing their top corner: the light one stands on a
    // level base, the dark one has its side along the light one's median down from that corner, so the
    // dark one's median (dashed) is level, perpendicular to it. The small circle (radius 1) is
    // inscribed between the light one's left side, the shared line and the dashed median; the large
    // one in the light one's right half. Its area? Nothing to drag: the triangles fit one way.
    [Puzzle("two-perpendicular-equilateral-triangles")]
    public static void TwoPerpendicularEquilateralTriangles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double s = 4, h = 2 * S3 * s / 2;
        var t = d.Pt(0, 0, "T");
        var baseLeft = d.Pt(-s / 2, -h);
        var baseRight = d.Pt(s / 2, -h);
        var b = d.Pt(0, -s, "B");
        var l = d.Pt(-h, -s / 2, "L");
        var m = d.Pt(0, -s / 2, "M");
        // the small circle is inscribed in the right triangle T, M and the left side at M's height
        double r1 = (s / 2 + s / 2 / Math.Sqrt(3) - s / Math.Sqrt(3)) / 2;
        // the large one in the right triangle T, the base's middle and the right corner
        double r2 = (s / 2 + h - s) / 2;
        d.Polygon(d.Outline("LightOrange", "#FFEFA055", 3.5), t, baseLeft, baseRight);
        string dark = d.Line("DarkOrange", "#FFE0681E", 3.5);
        d.Segment(dark, t, l);
        d.Segment(dark, l, b);
        d.Segment(dark, b, t);
        d.Dashed(dark, l, m, 8);
        string magenta = d.Outline("Magenta", "#FFCC4C80", 3.5);
        var small = d.Pt(-r1, -s / 2 + r1, "C1");
        var large = d.Pt(r2, -h + r2, "C2");
        d.Circle(magenta, small, d.Pt(0, -s / 2 + r1));
        d.Circle(magenta, large, d.Pt(0, -h + r2));
        string text = d.Text("MagentaText", "#FFCC4C80", 24);
        d.Label("1", small, text, 24);
        d.Label("?", large, d.Text("MagentaQuestion", "#FFCC4C80", 32), 32);
        d.Scene(-h, 0, s / 2, -s);
    }

    // A circle, a square in it, a circle in the square, a regular hexagon (a corner at the top) in that
    // circle and the shaded circle in the hexagon. What fraction of the design is shaded? Nothing to
    // drag: every shape fits the one around it.
    [Puzzle("nested-circles-and-polygons-ii")]
    public static void NestedCirclesAndPolygonsII(Lgf d)
    {
        const double R = 4, r = R / 1.4142135623730951, inner = r * S3;
        var o = d.Pt(0, 0, "O");
        string teal = d.Outline("Teal", "#FF3A9EA0", 3.5);
        d.Circle(d.Fill("Blue", "#FFD3E8DF", "#FF3A9EA0", 3.5), o, d.Pt(0, inner));
        d.Circle(teal, o, d.Pt(0, R));
        d.Box(d.Outline("LightGreen", "#FFA8E070", 3.5), V.At(-r, -r), V.At(r, r));
        d.Circle(teal, o, d.Pt(0, r));
        // over the shaded circle, which touches its sides
        d.Inked(d.Outline("Green", "#FF0E9A3A", 3.5), Enumerable.Range(0, 6).Select(i => d.Pt(r * Math.Cos((90 + 60 * i) * Math.PI / 180), r * Math.Sin((90 + 60 * i) * Math.PI / 180))).ToArray());
        d.Scene(-R, R, R, -R);
    }

    // Ten dots 4 apart on a vertical line joined by semicircles into one double spiral: from the
    // bottom dot it winds in to the middle, turns on the small semicircle and winds out to the top dot.
    // The left semicircles are centered on the fifth dot, the right ones on the sixth. Is more red or
    // orange, by how much? Nothing to drag: the spacing fixes it.
    [Puzzle("spiralling-semi-circles")]
    public static void SpirallingSemiCircles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double g = 4;
        var dots = Enumerable.Range(0, 10).Select(k => d.Pt(0, -g * k, "D" + k)).ToArray();
        var left = dots[4];
        var right = dots[5];
        var small = d.Pt(0, -4.5 * g, "S");
        // a half disc left of the line (counterclockwise from top to bottom) or right of it
        string Left(V center, double r) => d.Arc("CircleSector", null, center, d.Pt(center + V.At(0, r)), d.Pt(center + V.At(0, -r)), visible: false);
        string Right(V center, double r) => d.Arc("CircleSector", null, center, d.Pt(center + V.At(0, -r)), d.Pt(center + V.At(0, r)), visible: false);
        string l16 = Left(left, 4 * g), l12 = Left(left, 3 * g), l8 = Left(left, 2 * g), l4 = Left(left, g), s2 = Left(small, g / 2);
        string r16 = Right(right, 4 * g), r12 = Right(right, 3 * g), r8 = Right(right, 2 * g), r4 = Right(right, g);
        // one region per color: pieces of one fill side by side would show seams under the dark theme
        d.Region(d.Fill("Red", "#FFFA1505", null), "Union",
            d.Region(null, "Difference", l16, l12), d.Region(null, "Difference", r12, r8), d.Region(null, "Difference", l8, l4), r4, s2);
        d.Region(d.Fill("Orange", "#FFFDA204", null), "Union",
            d.Region(null, "Difference", r16, r12), d.Region(null, "Difference", l12, l8), d.Region(null, "Difference", r8, r4), d.Region(null, "Difference", l4, s2));
        // across the orange, which stays bright in the dark theme: a mid-tone blue there
        string blue = d.Line("Blue", "#FF1E28B4", 3.5, dark: "#FF6B7BFF");
        for (int k = 0; k < 4; k++)
        {
            // left: dot k to dot 8 - k around the fifth dot; right: dot 9 - k to dot k + 1 around the sixth
            d.Arc("CircleArc", blue, left, dots[k], dots[8 - k]);
            d.Arc("CircleArc", blue, right, dots[9 - k], dots[k + 1]);
        }

        d.Arc("CircleArc", blue, small, dots[4], dots[5]);
        foreach (var dot in dots)
        {
            d.DotAt(dot, null, 10);
        }

        string arrow = d.Line("Arrow");
        var top = d.Pt(1.4, 0);
        var middle = d.Pt(1.4, -g / 2);
        var bottom = d.Pt(1.4, -g);
        d.Vector(arrow, middle, top);
        d.Vector(arrow, middle, bottom);
        d.Label("4", V.At(4, -g / 2), d.Text(), 26);
        d.Scene(-4 * g, 0, 4 * g, -9 * g);
    }

    // Five squares: the big green one standing on a corner; the light green one on its left corner;
    // the teal one with a side of the light green one as its diagonal and its next corner at the big
    // one's top; the blue one on the teal and light green ones' shared corner, its top side the
    // shaded square's diagonal; the shaded one in the big one's top corner. What fraction of the big
    // square is shaded? Nothing to drag: the corners the squares share fix them (the light green one
    // leans 1 in 2; the shaded one is 1/25 of the big one).
    [Puzzle("five-decreasing-squares")]
    public static void FiveDecreasingSquares(Lgf d)
    {
        d.Texture = "FeltTip";
        var a = d.Pt(0, 0, "A");
        var t = d.Pt(5, 5, "T");
        var q = d.Pt(4, 2, "Q");
        var p = d.Pt(2, 6, "P");
        var l = d.Pt(4, 4, "L");
        var r = d.Pt(6, 4, "R");
        d.Polygon(d.Outline("Green", "#FF2E9A6A", 3.5), a, d.Pt(5, -5), d.Pt(10, 0), t);
        d.Polygon(d.Outline("LightGreen", "#FF9EE060", 3.5), a, q, p, d.Pt(-2, 4));
        d.Polygon(d.Outline("Teal", "#FF3E92B8", 3.5), t, p, d.Pt(1, 3), q);
        d.Polygon(d.Outline("Blue", "#FF2C5CB0", 3.5), q, d.Pt(6, 2), r, l);
        d.Polygon(d.Fill("Navy", "#FF04089C", "#FF04089C", 3.5), t, l, d.Pt(5, 3), r);
        d.Scene(-2, 6, 10, -5);
    }

    // An equilateral triangle in a circle and a semicircle on its base; the parts inside one of them and
    // not the other are shaded. What fraction of the circle? Nothing to drag: the circle fixes both.
    [Puzzle("triangle-and-semi-circle-in-a-circle")]
    public static void TriangleAndSemiCircleInACircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4;
        var o = d.Pt(0, 0, "O");
        var apex = d.Pt(0, R, "Apex");
        var left = d.Pt(-S3 * R, -R / 2, "Bl");
        var right = d.Pt(S3 * R, -R / 2, "Br");
        var m = d.Pt(0, -R / 2, "M");
        string triangle = d.Polygon(null, apex, left, right);
        string half = d.Arc("CircleSector", null, m, right, left, visible: false);
        string orange = d.Fill("Orange", "#FFFBA42A", null);
        d.Region(orange, "Difference", triangle, half);
        d.Region(orange, "Difference", half, triangle);
        d.Circle(d.Outline("Blue", "#FF14C4DC", 3.5), o, apex);
        string red = d.Line("Red", "#FFF03C0A", 3.5);
        d.Segment(red, left, apex);
        d.Segment(red, apex, right);
        string magenta = d.Line("Magenta", "#FFF0189A", 3.5);
        d.Arc("CircleArc", magenta, m, right, left);
        d.Segment(magenta, left, right);
        d.Scene(-R, R, R, -R);
    }

    // A parallelogram, its top side cut 2 and 3 at P; lines from P to the bottom corners, and from the
    // top corners square onto those lines. The three white triangles are isosceles: the two at the top
    // are right-angled with 45° at P, so the lines from P lean at 45° and the middle triangle is
    // right-angled too, its height half the base. What fraction is shaded? Nothing to drag: 2 and 3 fix
    // it (the shaded part is 2 × 3 / 2).
    [Puzzle("subdivided-parallelogram")]
    public static void SubdividedParallelogram(Lgf d)
    {
        d.Texture = "FeltTip";
        const double a = 2, b = 3, h = (a + b) / 2;
        var topLeft = d.Pt(0, 0, "TL");
        var p = d.Pt(a, 0, "P");
        var topRight = d.Pt(a + b, 0, "TR");
        var bottomLeft = d.Pt(a - h, -h, "BL");
        var bottomRight = d.Pt(a + h, -h, "BR");
        var x = d.Pt(a / 2, -a / 2, "X");
        var y = d.Pt(a + b / 2, -b / 2, "Y");
        string green = d.Fill("Green", "#FFBEE04A", null);
        d.Polygon(green, topLeft, x, bottomLeft);
        d.Polygon(green, topRight, bottomRight, y);
        // along the light green, which stays bright in the dark theme: a mid green there
        const string darkGreen = "#FF34B848";
        string line = d.Line("Line", "#FF10A81C", 3.5, dark: darkGreen);
        d.Segment(line, p, bottomLeft);
        d.Segment(line, p, bottomRight);
        d.Segment(line, topLeft, x);
        d.Segment(line, topRight, y);
        d.Inked(d.Outline("Frame", "#FF10A81C", 3.5, dark: darkGreen), topLeft, topRight, bottomRight, bottomLeft);
        d.Label("2", V.At(a / 2, 0.32));
        d.Label("3", V.At(a + b / 2, 0.32));
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(a - h, 0.55, a + b, -h - 0.5);
    }

    // Four rectangles of area 12 filling a big one: a long one across the top, under it the yellow one
    // on the left and two equal ones stacked on the right. The triangle from the top left corner to
    // the yellow one's bottom right corner and the right end of the line between the stacked ones: its
    // area? The areas make the yellow one a third of the width and the top one a quarter of the height
    // (the width times the height is 48). Dragging the bottom left corner changes the width: the
    // rectangles change shape, the triangle stays 19.
    [Puzzle("triangle-in-four-rectangles")]
    public static void TriangleInFourRectangles(Lgf d)
    {
        var l = d.Slider(V.At(-9.6, 0), V.At(-7.8, 0), (9.6 - 8.5) / 1.8, (-8.5, 0), "L");
        string w = d.Num("-L.X", "w");
        string h = d.Num($"48 / {w}", "h");
        var topLeft = d.Pt(new V($"-{w}", h), "TL");
        var p = d.Pt(new V($"-2 * {w} / 3", "0"), "P");
        var q = d.Pt(new V("0", $"18 / {w}"), "Q");
        d.Box(d.Outline("Cyan", "#FF3CB8C4", 3.5), new V($"-{w}", $"36 / {w}"), new V("0", h));
        d.Box(d.Outline("Yellow", "#FFF8DA10", 3.5), l, new V($"-2 * {w} / 3", $"36 / {w}"));
        d.Box(d.Outline("Green", "#FFB4DC40", 3.5), q + new V($"-2 * {w} / 3", "0"), new V("0", $"36 / {w}"));
        d.Box(d.Outline("Blue", "#FF3E90C0", 3.5), p, q);
        // a light gray on the dark paper: the default light tint would be as yellow as the yellow one
        d.Polygon(d.Outline("Triangle", "#FF5A5030", 3.5, dark: "#FFC8C4BC"), topLeft, p, q);
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(-9.6, 48 / 7.8, 0, -0.8);
    }

    // A circle of circumference 10 inscribed in a trapezium of perimeter 15 (level top and bottom).
    // What fraction is shaded? Dragging the top left corner along the top tilts the left side; the
    // right side tilts so the perimeter stays 15 (with the circle inside, the top and bottom add up to
    // half of it, and each side is the radius over the sine of its angle): the shape changes, the
    // fraction stays 2/3.
    [Puzzle("circle-in-a-trapezium")]
    public static void CircleInATrapezium(Lgf d)
    {
        double r = 5 / Math.PI;
        string rr = Lgf.N(r);
        var o = d.Pt(0, 0, "O");
        var topLeft = d.Slider(V.At(-0.58 * r, r), V.At(-1.15 * r, r), (0.8 - 0.58) / (1.15 - 0.58), (-0.8 * r, r), "TL");
        // tan of half the bottom left angle; 1 / sin of the bottom right one; its sine, cosine and tan of its half
        string t = d.Num($"-TL.X / {rr}", "t");
        string k = d.Num($"{Lgf.N(3 * Math.PI / 4)} - (1 + {t} ^ 2) / (2 * {t})", "k");
        string sin = d.Num($"1 / {k}", "sinb");
        string cos = d.Num($"sqrt(1 - {sin} ^ 2)", "cosb");
        string u = d.Num($"(1 - {cos}) / {sin}", "u");
        var topRight = d.Pt(new V($"{rr} * {u}", rr), "TR");
        var bottomRight = d.Pt(new V($"{rr} / {u}", Lgf.N(-r)), "BR");
        var bottomLeft = d.Pt(new V($"-{rr} / {t}", Lgf.N(-r)), "BL");
        const string teal = "#FF4AAFBA", magenta = "#FFDC3C88";
        d.Circle(d.Fill("Green", "#FFCDE6BE", teal, 3.5), o, d.Pt(0, r));
        // over the circle, which touches every side
        d.Inked(d.Outline("Magenta", magenta, 3.5), topLeft, topRight, bottomRight, bottomLeft);
        // above the trapezium's level, where the texts (as wide as most of the phone's view) stay clear of it
        d.Label("Circumference = 10", V.At(3.5, r + 1.45), d.Text("TealText", teal, 24), 24);
        d.Label("Perimeter = 15", V.At(3.9, r + 0.6), d.Text("MagentaText", magenta, 24), 24);
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(-2.9, r + 1.9, 7, -r - 0.6);
    }

    // A right triangle with sides 6, 8 and 10 (her colors: the left side two yellow 3s, the bottom two
    // red 5s, the right side red 5 then yellow 3); blue lines from the left side's middle to the right
    // corner and from the top to the bottom's middle. The shaded triangle between them and the top?
    // Nothing to drag: the lengths fix it (3, 3 / 5, 5 / 5, 3 make the sides 6, 10, 8: right-angled at the top).
    [Puzzle("subdivided-triangle")]
    public static void SubdividedTriangle(Lgf d)
    {
        d.Texture = "FeltTip";
        var a = d.Pt(0, 0, "A");
        var c = d.Pt(10, 0, "C");
        var t = d.Pt(3.6, 4.8, "T");
        var m = d.Pt(1.8, 2.4, "M");
        var e = d.Pt(5, 0, "E");
        var f = d.Pt(7.6, 1.8, "F");
        var x = d.Crossing(m, c, t, e, "X");
        d.Polygon(d.Fill("Teal", "#FF32C0C8", null), t, m, x);
        const string yellow = "#FFF4DE06", red = "#FFD05A10";
        string yellowLine = d.Line("Yellow", yellow, 3.5);
        string redLine = d.Line("Red", red, 3.5);
        d.Segment(yellowLine, a, m);
        d.Segment(yellowLine, m, t);
        d.Segment(redLine, a, e);
        d.Segment(redLine, e, c);
        d.Segment(redLine, t, f);
        d.Segment(yellowLine, f, c);
        string blue = d.Line("Blue", "#FF6A98C4", 3.5);
        d.Segment(blue, m, c);
        d.Segment(blue, t, e);
        d.DotAt(m, "Md");
        d.DotAt(e, "Ed");
        d.Label("3", V.At(0.25, 1.85), d.Text("YellowText", "#FFE6C400", 26));
        d.Label("5", V.At(2.5, -0.65), d.Text("RedText", red, 26));
        d.Scene(-0.3, 4.8, 10, -1);
    }

    // A square with a circle of area 8π in its top right corner and one of area 2π in its bottom left,
    // touching each other; the rectangle in the bottom right corner between them. Its area? Nothing to
    // drag: the circles fix the square (its side is 3 + 3√2).
    [Puzzle("two-circles-inside-a-square")]
    public static void TwoCirclesInsideASquare(Lgf d)
    {
        double big = 2 * Math.Sqrt(2), small = Math.Sqrt(2), s = 3 + 3 * Math.Sqrt(2);
        d.Box(d.Outline("Gray", "#FF8A96A8", 3), V.At(0, 0), V.At(s, s));
        const string lightBlue = "#FF15A4D2", blue = "#FF0A5CB4", red = "#FFD0503E";
        var bigCenter = d.Pt(s - big, s - big, "Cb");
        var smallCenter = d.Pt(small, small, "Cs");
        d.Circle(d.Outline("LightBlue", lightBlue, 3.5), bigCenter, d.Pt(s - big, s));
        d.Circle(d.Outline("Blue", blue, 3.5), smallCenter, d.Pt(small, 0));
        d.Box(d.Outline("Red", red, 3.5), V.At(2 * small, 0), V.At(s, s - 2 * big));
        d.Label("8π", bigCenter, d.Text("LightBlueText", lightBlue, 32), 32);
        d.Label("2π", smallCenter, d.Text("BlueText", blue, 30), 30);
        d.Label("?", V.At((2 * small + s) / 2, (s - 2 * big) / 2), d.Text("RedText", red, 26));
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(0, s, s, -0.8);
    }

    // Three concentric semicircles on one line and four squares standing on it: the middle two from the
    // center out to the yellow one's and the orange one's ends, the left one's corner on the orange one
    // and the right one's on the pink one (so each radius is √2 times the one inside it); the small ones
    // fill the gaps between the ends. The left one has area 4, the right one? Nothing to drag: the area
    // fixes it.
    [Puzzle("concentric-semi-circles")]
    public static void ConcentricSemiCircles(Lgf d)
    {
        double ry = 2 / (Math.Sqrt(2) - 1), ro = Math.Sqrt(2) * ry, rp = 2 * ry;
        var o = d.Pt(0, 0, "O");
        d.Arc("CircleArc", d.Line("Pink", "#FFE5677E", 3), o, d.Pt(rp, 0), d.Pt(-rp, 0));
        d.Arc("CircleArc", d.Line("Orange", "#FFEE8A3A", 3), o, d.Pt(ro, 0), d.Pt(-ro, 0));
        d.Arc("CircleArc", d.Line("Yellow", "#FFF8D838", 3), o, d.Pt(ry, 0), d.Pt(-ry, 0));
        const string ink = "#FF4A4840";
        string frame = d.Outline("Frame", ink, 3);
        d.Box(frame, V.At(-ro, 0), V.At(-ry, ro - ry));
        d.Box(frame, V.At(-ry, 0), V.At(0, ry));
        d.Box(frame, V.At(0, 0), V.At(ro, ro));
        d.Box(frame, V.At(ro, 0), V.At(rp, rp - ro));
        d.Segment(d.Line("Ink", ink, 3), d.Pt(-rp, 0), d.Pt(rp, 0));
        // small enough for the small squares on a phone
        string text = d.Text("Mark", ink, 22);
        d.Label("4", V.At(-(ro + ry) / 2, (ro - ry) / 2), text, 22);
        d.Label("?", V.At((ro + rp) / 2, (rp - ro) / 2), text, 22);
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(-rp, rp, rp, -1.5);
    }
}
