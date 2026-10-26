#nullable disable

static partial class Puzzles
{
    // Three circles of radius 1 in a rectangle cut into three cells by two upright lines, with its
    // diagonal across them: the left circle in the top left corner of its cell, above the diagonal and
    // touching it, the right one in the bottom right corner, below it and touching it, the middle one
    // halved by it; each touches its cell's upright sides. The rectangle's area? Nothing to drag: the
    // circles fix it (6 by 2.5).
    [Puzzle("three-circles-in-a-rectangle")]
    public static void ThreeCirclesInARectangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double W = 6, H = 2.5;
        // her lines are thick felt pen
        string yellow = d.Fill("Yellow", "#FFFEE502", Lgf.Ink, 3.5);
        string teal = d.Fill("Teal", "#FF76E6D0", Lgf.Ink, 3.5);
        d.Circle(yellow, d.Pt(1, H - 1), d.Pt(0, H - 1));
        d.Circle(teal, d.Pt(5, 1), d.Pt(6, 1));
        // the diagonal's direction is (12, 5) / 13
        var m = d.Pt(3, H / 2);
        d.Arc("CircleSegment", yellow, m, d.Pt(3 + 12 / 13.0, H / 2 + 5 / 13.0), d.Pt(3 - 12 / 13.0, H / 2 - 5 / 13.0));
        d.Arc("CircleSegment", teal, m, d.Pt(3 - 12 / 13.0, H / 2 - 5 / 13.0), d.Pt(3 + 12 / 13.0, H / 2 + 5 / 13.0));
        string ink = d.Line("Ink", Lgf.Ink, 3.5);
        d.Segment(ink, d.Pt(2, 0), d.Pt(2, H));
        d.Segment(ink, d.Pt(4, 0), d.Pt(4, H));
        d.Segment(ink, d.Pt(0, 0), d.Pt(W, H));
        d.Inked(d.Outline("Frame", Lgf.Ink, 3.5),d.Pt(0, 0), d.Pt(W, 0), d.Pt(W, H), d.Pt(0, H));
        d.Scene(0, H, W, 0);
    }

    // Two equal semicircles in a semicircle: one standing on the diameter at its right end, the other
    // with its diameter a chord of the big arc, its own arc touching the big diameter and the first
    // one. What fraction do they cover? Nothing to drag: that fixes them (radius R / sqrt 3, the
    // tilted one's center at (-R, R) / sqrt 3).
    [Puzzle("two-semi-circles-inside-a-semi-circle")]
    public static void TwoSemiCirclesInsideASemiCircle(Lgf d)
    {
        const double R = 4;
        double r = R / Math.Sqrt(3), u = r / Math.Sqrt(2);
        var o = d.Pt(0, 0, "O");
        d.Arc("CircleSegment", d.Outline("Semicircle", "#FFB060C4"), o, d.Pt(R, 0), d.Pt(-R, 0));
        string fill = d.Fill("Purple", "#FFBA8BD6");
        var m = d.Pt(-r, r, "M");
        d.Arc("CircleSegment", fill, m, d.Pt(-r - u, r - u, "A"), d.Pt(-r + u, r + u, "B"));
        d.Arc("CircleSegment", fill, d.Pt(R - r, 0, "C"), d.Pt(R, 0), d.Pt(R - 2 * r, 0));
        d.Scene(-R, R, R, 0);
    }

    // A yellow square standing on a red one of twice its area, their left sides in line; lines from the
    // yellow one's top right corner to the red one's bottom left and top right corners. The angle
    // between them? Nothing to drag: the areas fix it up to scale.
    [Puzzle("two-squares")]
    public static void TwoSquares(Lgf d)
    {
        const double s = 2;
        double b = s * Math.Sqrt(2);
        string redLine = "#FFBE2410";
        d.Box(d.Fill("Red", "#FFEA8E7C", null), V.At(0, 0), V.At(b, b));
        d.Box(d.Fill("Yellow", "#FFFFFF8A", "#FFFEEE02"), V.At(0, b), V.At(s, b + s));
        // inked after coloring: the red square's top runs under the yellow one
        d.Inked(d.Outline("RedFrame", redLine), d.Pt(0, 0), d.Pt(b, 0), d.Pt(b, b), d.Pt(0, b));
        var t = d.Pt(s, b + s, "T");
        var o = d.Pt(0, 0, "O");
        var c = d.Pt(b, b, "C");
        // the line to the bottom corner runs over the fills, which the dark theme leaves bright: ink
        // there too; the other is on the paper, and the mark across both is a mid-tone
        d.Segment(d.Line("InkOnFill", Lgf.Ink, 2.5, dark: Lgf.Ink), t, o);
        d.Segment(d.Line("Ink"), t, c);
        d.ArcMark(d.Line("MarkLine", Lgf.Ink, 2.5, dark: "#FF9A9A9A"), t, o, c, 0.55);
        // on the paper, beside the yellow square's side, which halves the angle
        d.Label("?", t + V.At(0.17, -1.0), d.Text("Mark", Lgf.Ink, 26));
        d.Scene(0, b + s, b, 0);
    }

    // An equilateral triangle standing on the line of the purple square's bottom side (area 8), its
    // right corner the square's; the pink square has the triangle's left side as its diagonal, and its
    // top side runs on through the purple square's top left corner, where the red square stands on the
    // pink one's side. The pink area? Nothing to drag: that fixes the triangle (its side is the purple
    // square's).
    [Puzzle("three-squares-and-an-equilateral-triangle")]
    public static void ThreeSquaresAndAnEquilateralTriangle(Lgf d)
    {
        double s = 2 * Math.Sqrt(2);
        string purple = "#FFA235E2", pink = "#FFEC4A8C";
        var p = d.Pt(0, 0, "P");
        var a = d.Pt(-s, 0, "A");
        var t = d.Pt(-s / 2, s * S3, "T");
        var q = d.Pt(0, s, "Q");
        // the pink square around its diagonal AT
        var m = Mid(d, a, t);
        var half = 0.5 * (t - a);
        var across = new V(V.Neg(half.Y), half.X);
        d.Polygon(d.Fill("Pink", "#FFFED2F4", pink), a, d.Pt(m - across), t, d.Pt(m + across));
        // the red square on TQ, away from the triangle
        var side = q - t;
        var down = new V(side.Y, V.Neg(side.X));
        d.Polygon(d.Outline("Red", "#FFE04535"), t, q, d.Pt(q + down), d.Pt(t + down));
        d.Polygon(d.Outline("Orange", "#FFF78D00"), a, p, t);
        d.Box(d.Outline("Purple", purple), V.At(0, 0), V.At(s, s));
        d.Label("8", V.At(s / 2, s / 2), d.Text("PurpleMark", purple, 46));
        d.Label("?", m + V.At(-0.32, 0.05), d.ShapeText("PinkMark", pink, 46));
        d.Scene(-3.35, s, s, 0);
    }

    // A pentagram in a circle, one of its lines through the center; angles of the same color are
    // equal (red at two points, blue at three). The orange angle, where the diameter crosses the line
    // from the top point? Nothing to drag: the equal angles fix the arcs between the points (90, 90,
    // 60, 60 and 60 degrees).
    [Puzzle("pentagram-in-a-circle")]
    public static void PentagramInACircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4;
        var o = d.Pt(0, 0, "O");
        // the points counterclockwise from the diameter's upper end
        double[] at = [35, 95, 155, 215, 305];
        var p = at.Select((a, i) => d.Pt(R * Math.Cos(a * Math.PI / 180), R * Math.Sin(a * Math.PI / 180), "P" + i)).ToArray();
        d.Circle(d.Outline("Circle", "#FF4A4A4A", 3.5, dark: Lgf.DarkInk), o, d.Pt(R, 0));
        string yellow = d.Line("Yellow", "#FFFFF206", 4);
        for (int i = 0; i < 5; i++)
        {
            d.Segment(yellow, p[i], p[(i + 2) % 5]);
        }

        string red = d.Wedge("RedWedge", "#FFF50502");
        string blue = d.Wedge("BlueWedge", "#FF04B4EE");
        for (int i = 0; i < 5; i++)
        {
            d.SectorMark(i == 1 || i == 2 ? red : blue, p[i], p[(i + 2) % 5], p[(i + 3) % 5], 0.55);
        }

        var x = d.Crossing(p[0], p[3], p[1], p[4], "X");
        string orange = "#FFFF8C00";
        d.ArcMark(d.Line("Orange", orange, 3.5), x, p[4], p[0], 0.5);
        d.Label("?", x + V.At(1.0, -0.12), d.Text("OrangeMark", orange, 32));
        d.DotAt(o, "Center", 12);
        d.Scene(-R, R, R, -R);
    }

    // A rectangle cut in four by a vertical and a horizontal line through the point where two
    // semicircles touch: the big one on the bottom side, the small one hanging from the top side at the
    // top right corner. The bottom right part is 24, the top left? The touching point is dragged along
    // the big arc: the small semicircle and the rectangle's height follow, and the two blue parts stay
    // equal (both are W^2 (1 - cos t) sin t / 4).
    [Puzzle("rectangle-quartered-by-semi-circles")]
    public static void RectangleQuarteredBySemiCircles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double W = 8, R = W / 2;
        string navy = "#FF0A2A9E";
        var b = d.Pt(R, 0, "B");
        var x = d.ArcSlider(b, R, 52, 70, 63, (R, 0), "X");
        string ux = d.Num($"(X.X - {R}) / {R}", "ux");
        string uy = d.Num($"X.Y / {R}", "uy");
        // the small one's radius, and its center on the line from B through X
        string r = d.Num($"{W} * (1 - {ux}) / (2 * (1 + {ux}))", "r");
        var s = d.Pt(b + $"({R} + {r})" * new V(ux, uy), "S");
        string h = "S.Y";
        d.Box(d.Fill("Blue", "#FF059CFB", null), new V("0", "X.Y"), new V("X.X", h));
        d.Box(d.Fill("Cyan", "#FF03CEFD", null), new V("X.X", "0"), new V(Lgf.N(W), "X.Y"));
        string line = d.Line("Navy", navy, 4);
        d.Segment(line, d.Pt(new V("X.X", "0")), d.Pt(new V("X.X", h)));
        d.Segment(line, d.Pt(new V("0", "X.Y")), d.Pt(new V(Lgf.N(W), "X.Y")));
        d.Polygon(d.Outline("Frame", navy, 4), d.Pt(0, 0), d.Pt(W, 0), d.Pt(new V(Lgf.N(W), h)), d.Pt(new V("0", h)));
        string pink = d.Line("Pink", "#FFE85A96", 3);
        d.Arc("CircleArc", pink, b, d.Pt(W, 0), d.Pt(0, 0));
        d.Arc("CircleArc", pink, s, d.Pt(new V($"S.X - {r}", h)), d.Pt(new V(Lgf.N(W), h)));
        string mark = d.ShapeText("NavyMark", navy, 38);
        d.Label("?", new V("X.X / 2", $"(X.Y + {h}) / 2"), mark);
        d.Label("24", new V($"(X.X + {W}) / 2", "X.Y / 2"), mark);
        d.Scene(0, 4.9, W, 0);
    }

    // A right triangle and a circle through its top and right-angle corners, cutting the base and the
    // long side; the chord between those two points (length 2) splits the triangle into equal areas.
    // The circle's area? The puzzle is under-specified (Mathforge says so): the chord is square to the
    // long side (the circle's diameter runs from the top corner), the green triangle is half the whole
    // and similar to it, so the upright side is 2 sqrt 2 and nothing fixes the base. Dragging the far
    // corner along the base: the chord stays 2 and the areas equal, the circle changes.
    [Puzzle("chord-dividing-a-triangle-into-two")]
    public static void ChordDividingATriangleIntoTwo(Lgf d)
    {
        double k = 2 * Math.Sqrt(2);
        string blue = "#FF1A4FB0";
        var c = d.Pt(0, 0, "C");
        var a = d.Pt(0, k, "A");
        var b = d.Slider(V.At(4.5, 0), V.At(9.5, 0), (7.55 - 4.5) / 5, (7.55, 0), "B");
        string hypotenuse = d.Num("distance(A, B)", "ab");
        var q = d.Pt(new V($"B.X - {hypotenuse} / sqrt(2)", "0"), "Q");
        var p = d.Pt(b + $"B.X / (sqrt(2) * {hypotenuse})" * (a - b), "P");
        d.Polygon(d.Fill("Blue", "#FF76CBE2", null), a, c, q, p);
        d.Polygon(d.Fill("Green", "#FFD6EB68", null), q, b, p);
        d.Circle(d.Outline("Circle", "#FF1AA8C0", 3), Mid(d, a, q), a);
        d.Inked(d.Outline("Frame", Lgf.Ink, 3.5), a, c, b);
        // on the blue and the green, which the dark theme leaves bright: deep there too
        string arrow = d.Line("Arrow", blue, 2.5, dark: "#FF14286E");
        var m = Mid(d, p, q);
        d.Vector(arrow, m, p);
        d.Vector(arrow, m, q);
        d.Label("2", m + V.At(-0.42, 0), d.ShapeText("BlueMark", blue, 34));
        d.Scene(-0.8, k + 0.4, 7.55, -0.4);
    }

    // Two regular pentagons sharing a corner; lines from the yellow one's top corner to the orange
    // one's bottom corner (each next to the shared corner) and from the yellow one's bottom left corner
    // to the orange one's right corner (each two away). The angle where they cross? Dragging the orange
    // one's top corner turns and resizes it: the angle stays (72 degrees between the lines).
    [Puzzle("two-pentagons")]
    public static void TwoPentagons(Lgf d)
    {
        d.Texture = "FeltTip";
        const double A = 4;
        var v = d.Pt(0, 0, "V");
        // the yellow one counterclockwise from the shared corner, its bottom side level
        var yellow = new List<V> { v };
        double x = 0, y = 0;
        for (int i = 0; i < 4; i++)
        {
            double angle = (252 - 72 * i) * Math.PI / 180;
            x += A * Math.Cos(angle);
            y += A * Math.Sin(angle);
            yellow.Add(d.Pt(x, y));
        }

        var (bl, t) = (yellow[2], yellow[4]);
        // the orange one's top corner stays in an annular sector: turned 28 to 70 degrees, side 0.55 to 0.98 of the yellow one's
        var bounds = new List<V>();
        for (int i = 0; i <= 6; i++)
        {
            double angle = (28 + 7 * i) * Math.PI / 180;
            bounds.Add(d.Pt(0.98 * A * Math.Cos(angle), 0.98 * A * Math.Sin(angle)));
        }

        for (int i = 6; i >= 0; i--)
        {
            double angle = (28 + 7 * i) * Math.PI / 180;
            bounds.Add(d.Pt(0.55 * A * Math.Cos(angle), 0.55 * A * Math.Sin(angle)));
        }

        double b = 0.78 * A;
        var h = d.HandleIn(d.Polygon(null, bounds.ToArray()), b * Math.Cos(44 * Math.PI / 180), b * Math.Sin(44 * Math.PI / 180), "H");
        // the orange one clockwise from the shared corner
        var edge = h - v;
        var right = d.Pt(h + edge.Turn(-72), "Or");
        var lower = d.Pt(right + edge.Turn(-144));
        var bottom = d.Pt(lower + edge.Turn(-216), "Ob");
        d.Polygon(d.Fill("Yellow", "#FFFFF801", "#FFF5E000", 3), yellow.ToArray());
        d.Polygon(d.Fill("Orange", "#FFFDC701", "#FFF88010", 3), v, h, right, lower, bottom);
        // mostly over the fills, which the dark theme leaves bright, crossing on the paper: a mid-tone there
        string ink = d.Line("Ink", Lgf.Ink, 3.5, dark: "#FF666666");
        d.Segment(ink, t, bottom);
        d.Segment(ink, bl, right);
        var c = d.Crossing(t, bottom, bl, right, "X");
        string pink = "#FFF34D65";
        d.ArcMark(d.Line("Pink", pink, 3.5), c, bl, bottom, 0.6);
        d.Label("?", c + V.At(0.05, -1.35), d.Text("PinkMark", pink, 36));
        d.Scene(-6.5, 2.4, 5, -3.8);
    }

    // A rectangle with two touching circles inside: the small one in the bottom left corner, the big
    // one touching the top, the bottom and the right side. Lines from where they touch to where the
    // small one touches the left side and the big one the top: the angle between them? Dragging the
    // small one's point on the left side resizes it (and the rectangle's width): the angle stays.
    [Puzzle("two-circles-in-a-rectangle")]
    public static void TwoCirclesInARectangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 3;
        var l = d.Slider(V.At(0, 0.35 * R), V.At(0, 0.82 * R), (2.08 / R - 0.35) / 0.47, (0, 2.08), "L");
        string r = "L.Y";
        var s = d.Pt(new V(r, r), "S");
        var b = d.Pt(new V($"{r} + 2 * sqrt({R} * {r})", Lgf.N(R)), "Bc");
        var t = d.Pt(s + $"{r} / ({R} + {r})" * (b - s), "T");
        var u = d.Pt(new V("Bc.X", Lgf.N(2 * R)), "U");
        string width = $"Bc.X + {R}";
        d.Circle(d.Outline("Cyan", "#FF04D8E8", 3.5), s, l);
        d.Circle(d.Outline("Yellow", "#FFF7EE10", 3.5), b, u);
        d.Polygon(d.Outline("Blue", "#FF2A4FD0", 3.5), d.Pt(0, 0), d.Pt(new V(width, "0")), d.Pt(new V(width, Lgf.N(2 * R))), d.Pt(0, 2 * R));
        string ink = d.Line("Ink", Lgf.Ink, 3.5);
        d.Segment(ink, t, l);
        d.Segment(ink, t, u);
        d.Angle(ink, t, l, u, 30);
        d.Label("?", t + V.At(-0.2, 1.05), d.Text("Mark", Lgf.Ink, 34));
        d.Scene(0, 2 * R, 2.08 + 2 * Math.Sqrt(R * 2.08) + R, 0);
    }

    // Three congruent rectangles: an upright one and a lying one sharing the bottom left corner, and
    // a tilted one whose diagonal runs from the upright one's top right corner to the lying one's.
    // The angle between the upright one's top and the tilted one's short side? Nothing to drag: that
    // fixes the rectangles (2 + sqrt 3 to 1).
    [Puzzle("angle-formed-by-three-rectangles")]
    public static void AngleFormedByThreeRectangles(Lgf d)
    {
        const double w = 1;
        double l = 2 + Math.Sqrt(3), diagonal = Math.Sqrt(l * l + w * w);
        var p = d.Pt(w, l, "P");
        var q = d.Pt(l, w, "Q");
        // from P a short side to the tilted one's top corner, turned from the diagonal toward the top
        double ux = 1 / Math.Sqrt(2), uy = -ux;
        double cos = w / diagonal, sin = l / diagonal;
        var top = d.Pt(w + w * (cos * ux - sin * uy), l + w * (cos * uy + sin * ux), "A");
        var bottom = d.Pt(new V("P.X + Q.X - A.X", "P.Y + Q.Y - A.Y"), "B");
        d.Box(d.Outline("Blue", "#FF1A35FA", 3), V.At(0, 0), V.At(w, l));
        d.Box(d.Outline("Red", "#FFFC0202", 3), V.At(0, 0), V.At(l, w));
        d.Polygon(d.Outline("Orange", "#FFFF8C10", 3), p, top, q, bottom);
        string purple = "#FFBF5FE3";
        var left = d.Pt(0, l);
        d.ArcMark(d.Line("Purple", purple, 3), p, top, left, 0.3);
        d.Label("?", V.At(w - 0.12, l + 0.62), d.Text("PurpleMark", purple, 30));
        // room at the bottom right for the page's buttons
        d.Scene(0, l + 0.8, l, -0.35);
    }

    // A square with its sides cut in three by dots, and two rectangles at 45 degrees across it with
    // their corners on the dots. What fraction is shaded? Nothing to drag: the dots fix it.
    [Puzzle("two-rectangles-in-a-square")]
    public static void TwoRectanglesInASquare(Lgf d)
    {
        string green = "#FF02C820";
        V[] one = [d.Pt(0, 2), d.Pt(1, 3), d.Pt(3, 1), d.Pt(2, 0)];
        V[] two = [d.Pt(0, 1), d.Pt(2, 3), d.Pt(3, 2), d.Pt(1, 0)];
        d.Region(d.Fill("Green", "#FF78E8A4", null), "Union", d.Polygon(null, one), d.Polygon(null, two));
        string line = d.Outline("GreenLine", green, 3);
        d.Polygon(line, one);
        d.Polygon(line, two);
        d.Box(d.Outline("Frame", Lgf.Ink, 3.5), V.At(0, 0), V.At(3, 3));
        for (int i = 0; i < 4; i++)
        {
            d.DotAt(V.At(i, 0), size: 12);
            d.DotAt(V.At(i, 3), size: 12);
        }

        foreach (int i in new[] { 1, 2 })
        {
            d.DotAt(V.At(0, i), size: 12);
            d.DotAt(V.At(3, i), size: 12);
        }

        // room at the bottom right for the page's buttons
        d.Scene(0, 3, 3, -0.3);
    }

    // A rectangle with semicircles: outside, on the left and the right side and on the top side's
    // right part (each 36); inside, one on the bottom side and one hanging from the top side's left
    // part, touching. The other two's total area? Nothing to drag: the touching fixes the rectangle up
    // to scale (the small one's diameter is the height times (sqrt 5 - 1) / 2).
    [Puzzle("five-semi-circles-around-a-rectangle")]
    public static void FiveSemiCirclesAroundARectangle(Lgf d)
    {
        const double H = 4;
        double s = H * (Math.Sqrt(5) - 1) / 2, w = s + H;
        string outer = d.Fill("Cyan", "#FFB6EDFD", "#FF02C0E6", 3);
        string inner = d.Fill("Blue", "#FFB5D6FB", "#FF4A1EE0", 3);
        d.Arc("CircleSegment", outer, d.Pt(0, H / 2), d.Pt(0, H), d.Pt(0, 0));
        d.Arc("CircleSegment", outer, d.Pt(w, H / 2), d.Pt(w, 0), d.Pt(w, H));
        d.Arc("CircleSegment", outer, d.Pt(s + H / 2, H), d.Pt(w, H), d.Pt(s, H));
        d.Arc("CircleSegment", inner, d.Pt(w / 2, 0), d.Pt(w, 0), d.Pt(0, 0));
        d.Arc("CircleSegment", inner, d.Pt(s / 2, H), d.Pt(0, H), d.Pt(s, H));
        d.Inked(d.Outline("Frame", Lgf.Ink, 3.5), d.Pt(0, 0), d.Pt(w, 0), d.Pt(w, H), d.Pt(0, H));
        string mark = d.ShapeText("CyanMark", "#FF00A8D8", 30);
        d.Label("36", V.At(-H / 4 - 0.1, H / 2), mark);
        d.Label("36", V.At(w + H / 4 + 0.1, H / 2), mark);
        d.Label("36", V.At(s + H / 2, H + H / 4 + 0.1), mark);
        d.Scene(-H / 2, 1.5 * H, w + H / 2, 0);
    }

    // A rectangle, the diagonal from its top right corner and a line from there to the bottom side
    // splitting the corner's right angle into three equal angles. What fraction is shaded (the triangle
    // between the two lines)? Nothing to drag: the angles fix the rectangle (sqrt 3 to 1).
    [Puzzle("three-equal-angles-in-a-rectangle")]
    public static void ThreeEqualAnglesInARectangle(Lgf d)
    {
        const double H = 3;
        double w = H * Math.Sqrt(3);
        string red = "#FFCE0204";
        var o = d.Pt(0, 0, "O");
        var v = d.Pt(w, H, "V");
        var p = d.Pt(w - H / Math.Sqrt(3), 0, "P");
        var left = d.Pt(0, H);
        var down = d.Pt(w, 0);
        d.Polygon(d.Fill("Red", "#FFE386A6", null), o, p, v);
        string line = d.Line("RedLine", red, 4);
        d.Segment(line, o, v);
        d.Segment(line, p, v);
        d.Polygon(d.Outline("Frame", red, 4), o, down, v, left);
        // she drew the three marks at about the same size, the middle one a little inside
        string mark = d.Line("Mark", red, 3);
        d.ArcMark(mark, v, left, o, 0.7);
        d.ArcMark(mark, v, o, p, 0.55);
        d.ArcMark(mark, v, p, down, 0.7);
        d.Scene(0, H, w, 0);
    }
}
