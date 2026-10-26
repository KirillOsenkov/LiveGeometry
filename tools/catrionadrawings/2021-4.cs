#nullable disable

static partial class Puzzles
{
    // Three squares of the same size: two side by side, the third tilted on the corner they share.
    // Lines from the row's outer bottom corners to the tilted square's top corner; the angle between
    // them? The top corner is on the circle around the shared corner through both bottom corners, so
    // dragging the tilted square's right corner turns it: the angle stays.
    [Puzzle("three-squares-iii")]
    public static void ThreeSquaresIII(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 4;
        var a = d.Pt(0, 0, "A");
        var b = d.Pt(2 * L, 0, "B");
        var m = d.Pt(L, L, "M");
        var r = d.ArcSlider(m, L, 3, 72, 19, (L, L), "R");
        var up = new V($"M.Y - R.Y", "R.X - M.X");
        var l = d.Pt(m + up, "Lc");
        var t = d.Pt(r + up, "T");
        d.Box(d.Fill("Green", "#FFBCEAC4"), V.At(0, 0), V.At(L, L));
        d.Box(d.Fill("Blue", "#FF67B2D1"), V.At(L, 0), V.At(2 * L, L));
        d.Polygon(d.Fill("Teal", "#FFA6E3E6"), m, r, t, l);
        string green = "#FF20935A";
        string line = d.Line("GreenLine", green, 3);
        d.Segment(line, a, t);
        d.Segment(line, t, b);
        d.Angle(line, t, a, b, 44);
        string ta = d.Num("distance(T, A)", "ta");
        string tb = d.Num("distance(T, B)", "tb");
        var bisector = new V($"(A.X - T.X) / {ta} + (B.X - T.X) / {tb}", $"(A.Y - T.Y) / {ta} + (B.Y - T.Y) / {tb}");
        d.Label("?", t + 1.2 * bisector, d.ShapeText("GreenMark", green, 34), 34);
        d.Scene(0, 9.2, 2 * L, 0);
    }

    // Four squares: a big one and its diagonal, a square on its bottom and one hanging from its top,
    // each with a corner on the diagonal, and a tilted one from the first's corner to the second's,
    // its other corners on the big square's top and right sides. What fraction is shaded? The corners
    // on the sides fix it (the hanging square's side is a fifth of the big one's, the bottom one's two
    // fifths): nothing to drag.
    [Puzzle("four-squares-viii")]
    public static void FourSquaresVIII(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 5;
        d.Segment(d.Line("Diagonal", "#FFF2C61A", 3), d.Pt(0, L), d.Pt(L, 0));
        string frame = d.Outline("Frame");
        d.Box(frame, V.At(0, 0), V.At(L, L));
        string yellow = d.Fill("Yellow", "#FFFFFE5C");
        d.Box(yellow, V.At(1, 0), V.At(3, 2));
        d.Box(yellow, V.At(1, 4), V.At(2, 5));
        d.Polygon(yellow, d.Pt(3, 2), d.Pt(5, 3), d.Pt(4, 5), d.Pt(2, 4));
        d.Scene(0, L, L, 0);
    }

    // Seven quarter circles of radius 1 on a grid of two rows of four squares (no grid drawn), inside
    // a semicircle through the grid's bottom left corner and its top right one. What fraction do they
    // cover? The semicircle is fixed by the grid (radius 2.5): nothing to drag.
    [Puzzle("seven-quarter-circles-inside-a-semi-circle")]
    public static void SevenQuarterCirclesInsideASemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 2.5, x0 = -R;
        var o = d.Pt(0, 0, "O");
        d.Arc("CircleSegment", d.Outline("Frame"), o, d.Pt(R, 0), d.Pt(-R, 0));
        V At(double x, double y) => d.Pt(x0 + x, y);
        void Quarter(string style, double cx, double cy, double fromX, double fromY, double toX, double toY) =>
            d.Arc("CircleSector", style, At(cx, cy), At(fromX, fromY), At(toX, toY));

        Quarter(d.Fill("Purple", "#FFBBA3DD"), 1, 0, 1, 1, 0, 0);
        Quarter(d.Fill("Red", "#FFF47784"), 1, 1, 1, 0, 2, 1);
        Quarter(d.Fill("Orange", "#FFFEE7B1"), 2, 1, 2, 2, 1, 1);
        Quarter(d.Fill("Yellow", "#FFFEFF98"), 2, 1, 3, 1, 2, 2);
        Quarter(d.Fill("Green", "#FFB8E9A5"), 3, 1, 2, 1, 3, 0);
        Quarter(d.Fill("LightBlue", "#FF8DC9FD"), 3, 1, 3, 0, 4, 1);
        Quarter(d.Fill("Blue", "#FF308CF0"), 4, 1, 4, 2, 3, 1);
        d.Scene(-R, R, R, 0);
    }

    // Two regular hexagons and an equilateral triangle: the small hexagon (area 5) sits on the big
    // one's top side at its right end, the triangle's top corner is the small one's top right corner
    // and its sides run through the big one's top left and right corners. The big one's area? Those
    // corners fix the hexagons (the small one's side is the big one's over sqrt 3); the triangle's size
    // is free: dragging its left corner along its side resizes it, and the hexagons stay.
    [Puzzle("two-hexagons-and-a-triangle")]
    public static void TwoHexagonsAndATriangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double B = 2, s = 2 / 1.7320508075688772, k = 0.7071067811865476;
        var big = new[] { d.Pt(-B, 0), d.Pt(0, 0), d.Pt(B / 2, -2 * S3), d.Pt(0, -4 * S3), d.Pt(-B, -4 * S3), d.Pt(-1.5 * B, -2 * S3) };
        var small = new[] { d.Pt(-s, 0), d.Pt(0, 0), d.Pt(s / 2, 1), d.Pt(0, 2), d.Pt(-s, 2), d.Pt(-1.5 * s, 1) };
        var t = d.Pt(0, 2, "T");
        var l = d.Slider(V.At(-4.2 * k, 2 - 4.2 * k), V.At(-5.8 * k, 2 - 5.8 * k), (5.26 - 4.2) / 1.6, (-5.26 * k, 2 - 5.26 * k), "L");
        var r = d.Pt(t + (l - t).Turn(60), "R");
        d.Polygon(d.Fill("Purple", "#FFA077C8"), big);
        d.Polygon(d.Fill("Red", "#FFD95460"), small);
        d.Polygon(d.Outline("Triangle", "#FF1E1460"), t, l, r);
        d.Scene(-3.8, 2, 1.45, -3.6);
    }

    // Two squares sharing a corner inside a semicircle: the small one's opposite corner is the
    // diameter's left end and the big one's its right end, and each has a corner on the arc. The
    // angle between them at the shared corner? A corner on the arc sees the diameter at a right
    // angle, so the shared corner is where the chords from those corners to the far ends cross, and
    // the corners on the arc are a quarter of the circle apart. Dragging the small square's corner
    // along the arc resizes both: the angle stays.
    [Puzzle("two-squares-inside-a-semi-circle")]
    public static void TwoSquaresInsideASemiCircle(Lgf d)
    {
        const double R = 4;
        var o = d.Pt(0, 0, "O");
        var w = d.Pt(-R, 0, "W");
        var e = d.Pt(R, 0, "E");
        var a = d.ArcSlider(o, R, 100, 170, 153.7, (0, 0), "A");
        var b = d.Pt(new V("A.Y", "-A.X"), "B");
        var v = d.Crossing(a, e, w, b, "Vs");
        var small = new[] { w, a, v, d.Pt(w + v - a) };
        var big = new[] { v, b, e, d.Pt(v + e - b) };
        string semicircle = d.Arc("CircleSegment", null, o, e, w, visible: false);
        string smallShape = d.Polygon(null, small);
        string bigShape = d.Polygon(null, big);
        string deep = d.Fill("Orange", "#FFFBA762", null);
        string light = d.Fill("LightOrange", "#FFFBCEAB", null);
        d.Region(deep, "Intersection", smallShape, semicircle);
        d.Region(deep, "Intersection", bigShape, semicircle);
        d.Region(light, "Difference", smallShape, semicircle);
        d.Region(light, "Difference", bigShape, semicircle);
        string orange = "#FFF29A3A";
        d.Arc("CircleSegment", d.Outline("Semicircle", orange), o, e, w);
        string frame = d.Outline("Frame");
        d.Inked(frame, small);
        d.Inked(frame, big);
        d.Angle(d.Line("OrangeLine", orange, 3), v, a, b, 24);
        string va = d.Num("distance(Vs, A)", "va");
        string vb = d.Num("distance(Vs, B)", "vb");
        var up = new V($"(A.X - Vs.X) / {va} + (B.X - Vs.X) / {vb}", $"(A.Y - Vs.Y) / {va} + (B.Y - Vs.Y) / {vb}");
        d.Label("?", v + 1.3 * up, d.Text("OrangeMark", orange, 30), 30);
        d.Scene(-R, R, R, -2.3);
    }

    // A line cut in three equal parts (4 each), black semicircles on them (the middle one up, the
    // outer ones down), and two blue semicircles on the chords from the line's ends to a point of the
    // middle semicircle. The total shaded area? Dragging that point along the middle semicircle
    // changes the blue ones: their total stays.
    [Puzzle("five-semi-circles")]
    public static void FiveSemiCircles(Lgf d)
    {
        d.Texture = "FeltTip";
        var a = d.Pt(0, 0, "A");
        var q1 = d.Pt(4, 0, "Q1");
        var m = d.Pt(6, 0, "M");
        var q2 = d.Pt(8, 0, "Q2");
        var b = d.Pt(12, 0, "B");
        var p = d.ArcSlider(m, 2, 25, 155, 59, (6, 0), "P");
        string blue = "#FF2340B4";
        string fill = d.Fill("Blue", "#FF85D2FE", blue);
        d.Arc("CircleSegment", fill, Mid(d, a, p), p, a);
        d.Arc("CircleSegment", fill, Mid(d, p, b), b, p);
        string ink = d.Line("Ink");
        d.Arc("CircleArc", ink, d.Pt(2, 0), a, q1);
        d.Arc("CircleArc", ink, m, q2, q1);
        d.Arc("CircleArc", ink, d.Pt(10, 0), q2, b);
        d.Segment(ink, a, q1, "TwoTicks");
        d.Segment(ink, q1, q2, "TwoTicks");
        d.Segment(ink, q2, b, "TwoTicks");
        d.Dimension(d.Line("Arrow", blue), d.Text("BlueMark", blue, 34), q1, q2, -0.85, "4", 0.6);
        d.Scene(0, 4.55, 12, -2);
    }

    // Four squares: a big one, one beside it on the same line, and on that one two squares side by
    // side (filling its width, the left one reaching the big one's top). The green quadrilateral
    // from the big one's corners to the small one's corners has area 7; the four squares' total? The
    // green is a third of the total for any sizes: dragging the corner where the two top squares meet
    // trades their sizes (the big one stays), and the ratio stays.
    [Puzzle("four-squares-vii")]
    public static void FourSquaresVII(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 5.4;
        var h = d.Slider(V.At(L + 1.55, L - 1.55), V.At(L + 2.5, L - 2.5), (2 - 1.55) / 0.95, (L + 2, L - 2), "H");
        string a = d.Num($"H.X - {Lgf.N(L)}", "a");
        string b = d.Num($"{Lgf.N(L)} - 2 * {a}", "b");
        var o = d.Pt(0, 0, "O");
        var t = d.Pt(L, L, "T");
        var k = d.Pt(new V($"{Lgf.N(L)} + {a}", $"{a} + 2 * {b}"), "K");
        var s = d.Pt(new V($"{Lgf.N(L)} + {a} + {b}", $"{a} + {b}"), "S");
        d.Polygon(d.Fill("Green", "#FFD1F9B1", "#FF0E6031"), o, t, k, s);
        string frame = d.Outline("Frame");
        d.Box(frame, V.At(0, 0), V.At(L, L));
        d.Box(frame, V.At(L, 0), s);
        d.Box(frame, new V(Lgf.N(L), $"{a} + {b}"), new V($"{Lgf.N(L)} + {a}", Lgf.N(L)));
        d.Box(frame, new V($"{Lgf.N(L)} + {a}", $"{a} + {b}"), k + new V(b, "0"));
        d.Scene(0, L, L + 3.4, 0);
    }

    // Four semicircles on parallel bases: the largest (diameter 8) hangs down; two on its base, up,
    // split it at a point; the fourth stands on a chord of the largest, its arc through that point.
    // Everything inside the first three and outside the fourth is shaded; its area? The point fixes
    // the chord's depth (the fourth's radius squared plus its depth squared is 16, and so is the
    // point's distance squared from the middle plus the depth squared, twice): dragging the point
    // along the base moves the chord, and the shaded area stays.
    [Puzzle("four-parallel-semi-circles")]
    public static void FourParallelSemiCircles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4;
        var o = d.Pt(0, 0, "O");
        var w = d.Pt(-R, 0, "W");
        var e = d.Pt(R, 0, "E");
        var p = d.Slider(V.At(-3, 0), V.At(-0.45, 0), (3 - 1.29) / 2.55, (-1.29, 0), "P");
        string h = d.Num($"sqrt(({Lgf.N(R * R)} - P.X ^ 2) / 2)", "h");
        string half = d.Num($"sqrt({Lgf.N(R * R)} - {h} ^ 2)", "half");
        var m = d.Pt(new V("0", $"-{h}"), "M");
        var c1 = d.Pt(new V($"-{half}", $"-{h}"), "C1");
        var c2 = d.Pt(new V(half, $"-{h}"), "C2");
        var right = d.Pt(new V($"(P.X + {Lgf.N(R)}) / 2", "0"), "Cr");
        var left = d.Pt(new V($"(P.X - {Lgf.N(R)}) / 2", "0"), "Cl");
        string largest = d.Arc("CircleSegment", null, o, w, e, visible: false);
        string big = d.Arc("CircleSegment", null, right, e, p, visible: false);
        string small = d.Arc("CircleSegment", null, left, p, w, visible: false);
        string middle = d.Arc("CircleSegment", null, m, c2, c1, visible: false);
        string all = d.Region(Clear2021_4(d), "Union", largest, big, small);
        d.Region(d.Fill("Pink", "#FFFFBCC2", null), "Difference", all, middle);
        string magenta = d.Line("Magenta", "#FFCB3093", 3);
        d.Arc("CircleArc", magenta, o, w, e);
        d.Arc("CircleArc", magenta, right, e, p);
        d.Arc("CircleArc", magenta, left, p, w);
        d.Arc("CircleArc", magenta, m, c2, c1);
        d.Segment(magenta, w, e);
        d.Segment(magenta, c1, c2);
        d.Scene(-R, 2.7, R, -R);
    }

    // Three overlapping regular hexagons: a big flat one; a small flat one whose long diagonal is the
    // big one's upper left side; and an upright one whose long diagonal runs from that side's top to
    // the big one's right corner. A part in one hexagon or all three is colored (tan in the big one
    // only, area 12), a part in two is white. The total red area? The corners fix the sizes (the
    // small one half the big one, the upright one sqrt 3 / 2 of it): nothing to drag.
    [Puzzle("three-overlapping-hexagons")]
    public static void ThreeOverlappingHexagons(Lgf d)
    {
        double r3 = Math.Sqrt(3);
        var big = new[] { d.Pt(0, 0), d.Pt(2, 2 * r3), d.Pt(6, 2 * r3), d.Pt(8, 0), d.Pt(6, -2 * r3), d.Pt(2, -2 * r3) };
        var small = new[] { d.Pt(3, r3), d.Pt(2, 2 * r3), d.Pt(0, 2 * r3), d.Pt(-1, r3), d.Pt(0, 0), d.Pt(2, 0) };
        var upright = new[] { d.Pt(5, 3 * r3), d.Pt(2, 2 * r3), d.Pt(2, 0), d.Pt(5, -r3), d.Pt(8, 0), d.Pt(8, 2 * r3) };
        string bigShape = d.Polygon(null, big);
        string smallShape = d.Polygon(null, small);
        string uprightShape = d.Polygon(null, upright);
        string red = d.Fill("Red", "#FFF2365A", null);
        d.Region(red, "Difference", uprightShape, bigShape);
        d.Region(red, "Difference", smallShape, bigShape);
        d.Region(red, "Intersection", smallShape, uprightShape, bigShape);
        d.Region(d.Fill("Tan", "#FFFCD5A0", null), "Difference", bigShape, smallShape, uprightShape);
        string frame = d.Outline("Frame");
        d.Polygon(frame, big);
        d.Polygon(frame, small);
        d.Polygon(frame, upright);
        d.Label("12", V.At(3.7, -2.4), d.ShapeText("ShapeMark", Lgf.Ink, 36), 36);
        d.Scene(-1, 3 * r3, 8, -2 * r3);
    }

    // Four squares: a big yellow one with a small yellow one on it at its right end, and two equal
    // white ones stacked beside them, as high as both. 2 from the big one's top right corner to the
    // right end of the line between the white ones; the yellow area? The big side and the small side
    // are the white side plus and minus the same amount, so the yellow is twice that distance
    // squared: dragging the line's right end around the corner (distance 2) trades the sizes, and the
    // yellow area stays.
    [Puzzle("four-squares-vi")]
    public static void FourSquaresVI(Lgf d)
    {
        d.Texture = "FeltTip";
        var p = d.Pt(0, 0, "P");
        var q = d.ArcSlider(p, 2, 326, 344, 336.5, (0, 0), "Q");
        string w = "Q.X", t = "(-Q.Y)";
        string big = d.Num($"{w} + {t}", "big");
        string small = d.Num($"{w} - {t}", "small");
        string yellow = d.Fill("Yellow", "#FFFFFE58");
        d.Box(yellow, new V($"-{big}", $"-{big}"), V.At(0, 0));
        d.Box(yellow, new V($"-{small}", "0"), new V("0", small));
        string paper = d.PaperFill();
        d.Box(paper, new V("0", $"-{big}"), q);
        d.Box(paper, new V("0", "Q.Y"), new V(w, small));
        string purple = "#FF6D1494";
        string arrow = d.Line("Arrow", purple);
        var middle = Mid(d, p, q);
        d.Vector(arrow, middle, p);
        d.Vector(arrow, middle, q);
        d.Label("2", middle + 0.45 * new V("-Q.Y / 2", "Q.X / 2"), d.Text("PurpleMark", purple, 30), 30);
        d.Scene(-2.6, 1.1, 1.85, -2.6);
    }

    // A regular hexagon and two squares side by side in a strip, the strip's corner at the hexagon's
    // top left corner and the next one on its bottom side. A line from the hexagon's bottom left
    // corner through the middle of the strip's right side; its angle with the hexagon's upper right
    // side? That line always rises at 45 degrees: dragging the strip's corner along the bottom side
    // turns and resizes the squares, and the line and the angle stay.
    [Puzzle("two-squares-on-a-hexagon")]
    public static void TwoSquaresOnAHexagon(Lgf d)
    {
        d.Texture = "FeltTip";
        const double H = 4, h = 4 * S3;
        var hexagon = new[] { d.Pt(-H / 2, h, "TL"), d.Pt(H / 2, h, "TR"), d.Pt(H, 0, "Rv"), d.Pt(H / 2, -h), d.Pt(-H / 2, -h, "BL"), d.Pt(-H, 0) };
        var a = hexagon[0];
        var b = d.Slider(V.At(-0.42 * H, -h), V.At(0.42 * H, -h), (0.42 - 0.08) / 0.84, (-0.08 * H, -h), "B");
        // half the strip's long side turned a quarter turn: the squares' other side
        var across = new V("(TL.Y - B.Y) / 2", "(B.X - TL.X) / 2");
        var left = Mid(d, a, b);
        var m = d.Pt(left + across, "Mr");
        var top = d.Pt(a + across);
        var bottom = d.Pt(b + across);
        string navy = "#FF275895";
        d.Polygon(d.Fill("LightBlue", "#FF97D6FD", null), hexagon);
        string blue = d.Fill("Blue", "#FF2994EB");
        d.Polygon(blue, a, top, m, left);
        d.Polygon(blue, left, m, bottom, b);
        d.Polygon(d.Outline("Hexagon", navy), hexagon);
        var start = hexagon[4];
        var end = d.Pt(start + V.At(2.3 * H * 0.7071067811865476, 2.3 * H * 0.7071067811865476));
        string green = "#FF1F846D";
        string line = d.Line("Green", green);
        d.Segment(line, start, end);
        var x = d.Crossing(start, end, hexagon[1], hexagon[2], "X");
        // in units, as the "?" beside it is, so the two keep their places at every size
        d.ArcMark(line, x, hexagon[2], end, 0.5);
        d.Label("?", x + V.At(0.95, 0.05), d.Text("GreenMark", green, 30), 30);
        d.Scene(-H, 4.35, 4.55, -h);
    }

    /// <summary>A shape style that draws nothing: for a region that only makes up another region</summary>
    private static string Clear2021_4(Lgf d)
    {
        string texture = d.Texture;
        d.Texture = null;
        string name = d.Fill("Clear", "#00000000", null);
        d.Texture = texture;
        return name;
    }
}
