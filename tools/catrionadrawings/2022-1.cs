#nullable disable

static partial class Puzzles
{
    // A semicircle with seven equally spaced dots on its arc (the ends included), chords from the right
    // end to the second and third dots from the left, and from the third to the fifth and the fifth to
    // the top. What fraction is shaded? Nothing to drag: the dots decide it.
    [Puzzle("subdivided-semi-circle")]
    public static void SubdividedSemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4;
        var o = d.Pt(0, 0, "O");
        var p = Enumerable.Range(0, 7).Select(i => d.Pt(R * Math.Cos(i * Math.PI / 6), R * Math.Sin(i * Math.PI / 6), "P" + i)).ToArray();
        // each pink part as one shape: two fills side by side would show a seam between their gradients in the dark theme
        string clear = Clear2022(d);
        string pink = d.Fill("Pink", "#FFF8AAE0", null);
        d.Region(pink, "Union", d.Polygon(clear, p[6], p[0], p[5]), d.Arc("CircleSegment", null, o, p[5], p[6], visible: false));
        d.Region(pink, "Union", d.Polygon(clear, p[4], p[0], p[2]), d.Arc("CircleSegment", null, o, p[0], p[2], visible: false));
        d.Arc("CircleSegment", pink, o, p[2], p[3]);
        string ink = d.Line("Ink");
        d.Arc("CircleSegment", d.Outline("Frame"), o, p[0], p[6]);
        d.Segment(ink, p[0], p[5]);
        d.Segment(ink, p[0], p[4]);
        d.Segment(ink, p[4], p[2]);
        d.Segment(ink, p[3], p[2]);
        foreach (var point in p)
        {
            d.DotAt(point);
        }

        d.Scene(-4.1, 4.1, 4.1, -0.1);
    }

    // Three squares: a big blue one, a small one in its bottom right corner, and a green one on the same
    // line whose right side runs up the small one's left side. The hatched parts (the green square's part
    // outside the big one, the big one's part outside the other two) have the same area, so the big
    // square's side is the hypotenuse of the other two. The angle at the green square's top right corner
    // between the lines to the big square's two far corners? Dragging the small square's corner along the
    // big one's diagonal resizes the small and the green square: the angle stays.
    [Puzzle("three-squares-vi")]
    public static void ThreeSquaresVI(Lgf d)
    {
        const double B = 5;
        var q = d.Pt(0, 0, "Q");
        var r = d.Pt(B, B, "R");
        var h = d.Slider(V.At(B - 1.8, 1.8), V.At(B - 4.2, 4.2), (3.05 - 1.8) / 2.4, (B - 3.05, 3.05), "H");
        string s = "H.Y";
        string g = d.Num($"sqrt({Lgf.N(B * B)} - {s} ^ 2)", "g");
        var p = d.Pt(new V("H.X", g), "P");
        var left = new V($"H.X - {g}", "0");
        string blueLine = "#FF1AB4EC";
        d.Box(d.Fill("Green", "#FFC8EBA8", null), left, new V("0", g));
        d.Polygon(d.Fill("Blue", "#FFA6E4F5", null), d.Pt(new V("0", g)), p, h, d.Pt(new V(Lgf.N(B), s)), r, d.Pt(0, B));
        d.Box(d.Outline("BigSquare", blueLine), V.At(0, 0), V.At(B, B));
        d.Box(d.Outline("SmallSquare", "#FF5CD6F5"), new V("H.X", "0"), new V(Lgf.N(B), s));
        d.Box(d.Outline("GreenSquare", "#FF7FCB5A"), left, p);
        string ink = d.Line("Ink");
        d.Segment(ink, q, p);
        d.Segment(ink, p, r);
        d.Angle(ink, p, q, r, 30);
        d.Label("?", p + V.At(0.62, -0.5));
        d.Scene(-2.1, 5.1, 5.1, -0.1);
    }

    // Four squares, two big and two small (the ticks), in a staircase, and an orange triangle on three of
    // their corners; its area is 4, the squares' total? The triangle's area is half the two sides' squares
    // added, so its corner on the big left square's bottom is a handle on the circle around its top right
    // corner (radius sqrt 8): dragging it trades the big squares' size for the small ones', and the
    // triangle's area stays 4.
    [Puzzle("four-squares-and-a-triangle-ii")]
    public static void FourSquaresAndATriangleII(Lgf d)
    {
        d.Texture = "FeltTip";
        var o = d.Pt(0, 0, "O");
        var t = d.ArcSlider(o, Math.Sqrt(8), 228, 248, 234.5, (0, 0), "T");
        // the big squares' side is a = -T.Y, the small ones' b = -T.X
        string a = d.Num("-T.Y", "a"), b = d.Num("-T.X", "b");
        var far = d.Pt(new V($"{a} + {b}", $"{a} - {b}"), "F");
        d.Polygon(d.Fill("Orange", "#FFFC9553"), t, o, far);
        string frame = d.Outline("Frame");
        d.Box(frame, new V("T.Y", "T.Y"), V.At(0, 0));
        d.Box(frame, new V("T.X", $"T.Y - {b}"), new V("0", "T.Y"));
        d.Box(frame, new V("0", "T.X"), new V(b, "0"));
        d.Box(frame, new V(b, "T.X"), far);
        string ink = d.Line("Ink");
        d.Segment(ink, d.Pt(new V("T.Y", "0")), o, "OneTick");
        d.Segment(ink, d.Pt(new V("0", $"T.Y - {b}")), d.Pt(new V("0", "T.Y")), "TwoTicks");
        d.Segment(ink, d.Pt(new V("0", "T.X")), d.Pt(new V(b, "T.X")), "TwoTicks");
        d.Segment(ink, d.Pt(new V(b, "T.X")), d.Pt(new V($"{a} + {b}", "T.X")), "OneTick");
        d.Scene(-2.4, 0.7, 4.0, -4.0);
    }

    // Five equilateral triangles in a row on one line, alternately up and down: three small ones
    // (purple, dark blue, light blue) and two big ones (green, dark green). Lines from the first corner
    // to the green one's top and from the purple one's far corner to the dark green one's; the angle
    // where they cross? Dragging the corner where the small ones meet the big ones along the bottom
    // (the row keeps its length) trades their sizes: the angle stays.
    [Puzzle("five-equilateral-triangles-ii")]
    public static void FiveEquilateralTrianglesII(Lgf d)
    {
        d.Texture = "FeltTip";
        const double W = 6;
        var v0 = d.Pt(0, 0, "V");
        var m = d.Slider(V.At(1.1, 0), V.At(3.4, 0), (1.86 - 1.1) / 2.3, (1.86, 0), "M");
        var e = d.Pt(W, 0, "E");
        string s = "M.X";
        string l = d.Num($"{Lgf.N(W)} - M.X", "l");
        var p = d.Pt(new V($"-{s} / 2", $"{s3} * {s}"), "P");
        var a = d.Pt(new V($"{s} / 2", $"{s3} * {s}"), "A");
        var b = d.Pt(new V($"1.5 * {s}", $"{s3} * {s}"), "B");
        var g = d.Pt(new V($"{s} + {l} / 2", $"{s3} * {l}"), "G");
        var k = d.Pt(new V($"{s} + 1.5 * {l}", $"{s3} * {l}"), "K");
        d.Polygon(d.Fill("Purple", "#FF764390"), p, a, v0);
        d.Polygon(d.Fill("DarkBlue", "#FF1A74A9"), v0, m, a);
        d.Polygon(d.Fill("LightBlue", "#FF36B8DF"), a, b, m);
        d.Polygon(d.Fill("Green", "#FF85D658"), m, e, g);
        d.Polygon(d.Fill("DarkGreen", "#FF56A092"), g, k, e);
        // where the line from V to G crosses the one from P to K
        string ex = d.Num("K.X - P.X", "ex"), ey = d.Num("K.Y - P.Y", "ey");
        string t = d.Num($"(P.Y * {ex} - P.X * {ey}) / (G.Y * {ex} - G.X * {ey})", "t");
        var x = d.Pt(new V($"{t} * G.X", $"{t} * G.Y"), "X");
        string ink = d.Line("Ink");
        d.Segment(ink, v0, g);
        d.Segment(ink, p, k);
        string blue = "#FF3D8FD6";
        d.Angle(d.Line("Blue", blue), x, g, p, 30);
        d.Label("?", x + V.At(-0.3, 0.72), d.Text("BlueMark", blue, 24));
        d.Scene(-1, 3.55, 8.2, -0.05);
    }

    // A rectangle with semicircles on its left side, its top side and part of its bottom side, and a
    // quarter circle around its bottom right corner as high as the rectangle, touching the small
    // semicircle; the three semicircles meet in one point. A part inside one of the shapes is red or
    // blue, a part inside two is white. Is more red or blue? The meeting point and the touching fix the
    // rectangle (its height is the small semicircle's diameter times the golden ratio): nothing to drag.
    [Puzzle("three-semi-circles-in-a-rectangle")]
    public static void ThreeSemiCirclesInARectangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double r = 1, H = r * 3.23606797749979, W = H + 2 * r;
        var o = d.Pt(0, 0, "O");
        var tl = d.Pt(0, H, "TL");
        var tr = d.Pt(W, H, "TR");
        var br = d.Pt(W, 0, "BR");
        var foot = d.Pt(2 * r, 0, "F");
        var leftCenter = d.Pt(0, H / 2);
        var bottomCenter = d.Pt(r, 0);
        var topCenter = d.Pt(W / 2, H);
        string left = d.Arc("CircleSegment", null, leftCenter, o, tl, visible: false);
        string bottom = d.Arc("CircleSegment", null, bottomCenter, foot, o, visible: false);
        string top = d.Arc("CircleSegment", null, topCenter, tl, tr, visible: false);
        string quarter = d.Arc("CircleSector", null, br, tr, foot, visible: false);
        string red = d.Fill("Red", "#FFE04F56", null);
        string blue = d.Fill("Blue", "#FF295B8C", null);
        d.Region(red, "Difference", left, bottom, top);
        d.Region(red, "Difference", quarter, top, bottom);
        d.Region(blue, "Difference", top, left, quarter, bottom);
        d.Region(blue, "Difference", bottom, left, top);
        string ink = d.Line("Ink");
        d.Arc("CircleArc", ink, leftCenter, o, tl);
        d.Arc("CircleArc", ink, bottomCenter, foot, o);
        d.Arc("CircleArc", ink, topCenter, tl, tr);
        d.Arc("CircleArc", ink, br, tr, foot);
        d.Polygon(d.Outline("Frame"), o, br, tr, tl);
        d.Scene(0, H, W, 0);
    }

    // A rectangle, its diagonal, and a circle touching its bottom, its right side and the diagonal, with a
    // line across the rectangle through the circle's center: 2 from the left side to the diagonal along
    // it, 1 on from there to the circle. The rectangle's area? The lengths fix it (radius 1.5, 6 by
    // 4.5): nothing to drag.
    [Puzzle("a-circle-in-a-rectangle")]
    public static void ACircleInARectangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double r = 1.5, W = 6, H = 4.5;
        var o = d.Pt(0, 0, "O");
        var br = d.Pt(W, 0, "BR");
        var tr = d.Pt(W, H, "TR");
        var tl = d.Pt(0, H, "TL");
        var west = d.Pt(0, r, "A");
        var cross = d.Pt(r * W / H, r, "X");
        var near = d.Pt(W - 2 * r, r, "N");
        var c = d.Pt(W - r, r, "C");
        var east = d.Pt(W, r, "E");
        string clear = Clear2022(d);
        d.Polygon(d.Fill("Yellow", "#FFFFFFA4", null), west, cross, tr, tl);
        d.Polygon(d.Fill("Lavender", "#FFBCADEA", null), o, cross, west);
        string upper = d.Arc("CircleSegment", d.Fill("DarkGreen", "#FF68955F", null), c, east, near);
        string lower = d.Arc("CircleSegment", d.Fill("DarkRed", "#FFBB5469", null), c, near, east);
        d.Region(d.Fill("LightGreen", "#FFBBE0B5", null), "Difference", d.Polygon(clear, cross, east, tr), upper);
        d.Region(d.Fill("Pink", "#FFEA8ECF", null), "Difference", d.Polygon(clear, o, br, east, cross), lower);
        string ink = d.Line("Ink");
        d.Segment(ink, o, tr);
        d.Segment(ink, west, east);
        d.Circle(d.Outline("Circle"), c, east);
        d.Polygon(d.Outline("Frame"), o, br, tr, tl);
        d.DotAt(c);
        // on the fills, which are mid-tones in the dark theme: ink there too, as the numbers are
        string arrow = d.Line("Arrow", Lgf.Ink, 2.5, dark: Lgf.Ink);
        string mark = d.ShapeText("ShapeMark", Lgf.Ink, 22);
        d.Dimension(arrow, mark, west, cross, 0.3, "2", 0.3);
        d.Dimension(arrow, mark, cross, near, -0.3, "1", 0.3);
        d.Scene(0, H, W, 0);
    }

    // Two squares: a big one, and a tilted one with a corner at the big one's bottom left corner and
    // the opposite corner on its top side. The tilted one's part outside the big one is 1, inside it 3;
    // the big one's part right of it? The numbers fix both squares (the tilted one has area 4, its
    // diagonal rises 3 for 1 across): nothing to drag.
    [Puzzle("two-squares-iii")]
    public static void TwoSquaresIII(Lgf d)
    {
        double k = 1 / Math.Sqrt(5);
        var b = d.Pt(0, 0, "B");
        var r = d.Pt(4 * k, 2 * k, "R");
        var f = d.Pt(2 * k, 6 * k, "F");
        var l = d.Pt(-2 * k, 4 * k, "L");
        var p = d.Pt(0, 5 * k, "P");
        var br = d.Pt(6 * k, 0);
        var tr = d.Pt(6 * k, 6 * k);
        var tl = d.Pt(0, 6 * k);
        d.Polygon(d.Fill("Orange", "#FFEEA780", null), b, p, l);
        d.Polygon(d.Fill("Pink", "#FFE393E1", null), b, r, f, p);
        d.Polygon(d.Fill("Blue", "#FF70C8F7", null), b, br, tr, f, r);
        string frame = d.Outline("Frame");
        d.Polygon(frame, b, br, tr, tl);
        d.Polygon(frame, b, r, f, l);
        d.Label("1", Mid(d, b, p, l), d.ShapeText("OrangeMark", "#FFD9540A", 40));
        d.Label("3", Mid(d, b, r, f, p), d.ShapeText("PinkMark", "#FFB8236E", 40));
        d.Label("?", Mid(d, br, tr, f, r), d.ShapeText("BlueMark", "#FF1F5FB8", 40));
        d.Scene(-2 * k, 6 * k, 6 * k, 0);
    }

    // Six congruent rectangles stacked in steps (the upright one as high as two lying ones, so they are 2
    // by 1). Lines from the bottom left rectangle's top left corner to the top one's top right corner and
    // from the second bottom one's bottom left corner to the top one's top left corner; the angle where
    // they cross? Nothing to drag: congruence fixes the rectangles.
    [Puzzle("six-rectangles")]
    public static void SixRectangles(Lgf d)
    {
        d.Texture = "FeltTip";
        string orange = d.Fill("Orange", "#FFFC8F62");
        foreach (var (x0, y0, x1, y1) in new[] { (0, 0, 2, 1), (2, 0, 4, 1), (2, 1, 4, 2), (4, 0, 5, 2), (3, 2, 5, 3), (3, 3, 5, 4) })
        {
            d.Box(orange, V.At(x0, y0), V.At(x1, y1));
        }

        var a = d.Pt(0, 1, "A");
        var b = d.Pt(5, 4, "B");
        var c = d.Pt(2, 0, "C");
        var e = d.Pt(3, 4, "E");
        var x = d.Pt(45.0 / 17, 44.0 / 17, "X");
        string magenta = "#FFB02E6E";
        string line = d.Line("Magenta", magenta);
        d.Segment(line, a, b);
        d.Segment(line, c, e);
        d.Angle(line, x, e, a, 30);
        d.Label("?", x + V.At(-0.65, 0.55), d.Text("MagentaMark", magenta, 30));
        d.Scene(0, 4, 5, 0);
    }

    // Four houses, each a square with an equilateral triangle on top, side by side on the diameter of a
    // semicircle (4 across), the outer two roofs touching the arc. Their total area? The arc fixes the
    // houses: nothing to drag.
    [Puzzle("four-houses-in-a-semi-circle")]
    public static void FourHousesInASemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 2;
        double s = R / Math.Sqrt(1.5 * 1.5 + (1 + S3) * (1 + S3));
        string blue = "#FF2F45B8";
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0, "E");
        var west = d.Pt(-R, 0, "W");
        d.Arc("CircleSegment", d.Outline("Frame", blue), o, east, west);
        var houses = new[]
        {
            ("Pink", "#FFFFD5DE", "PinkRoof", "#FFFE7795"),
            ("Green", "#FF91E074", "TealRoof", "#FF61B49D"),
            ("Lavender", "#FFDFD0FB", "PurpleRoof", "#FF7B3BAD"),
            ("Yellow", "#FFFFF260", "OrangeRoof", "#FFFEB76C"),
        };
        for (int i = 0; i < 4; i++)
        {
            double x = (i - 2) * s;
            var (wall, wallColor, roof, roofColor) = houses[i];
            d.Box(d.Fill(wall, wallColor), V.At(x, 0), V.At(x + s, s));
            d.Polygon(d.Fill(roof, roofColor), d.Pt(x, s), d.Pt(x + s, s), d.Pt(x + s / 2, s + S3 * s));
        }

        d.Dimension(d.Line("Arrow", blue), d.Text("BlueMark", blue, 26), west, east, -0.24, "4", 0.16);
        d.Scene(-R, R, R, -0.3);
    }

    /// <summary>A shape style that draws nothing: for a polygon that only makes up a Region (the kit has no hidden polygon)</summary>
    private static string Clear2022(Lgf d)
    {
        string texture = d.Texture;
        d.Texture = null;
        string name = d.Fill("Clear", "#00000000", null);
        d.Texture = texture;
        return name;
    }
}
