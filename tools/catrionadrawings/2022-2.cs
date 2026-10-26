#nullable disable

static partial class Puzzles
{
    // A square, its dashed diagonal, a light blue triangle on its left side whose long side runs from
    // the top left corner to the bottom, and a dark blue triangle from where that side crosses the
    // diagonal to the bottom right corner, its two marked angles equal. The angle at the top left
    // corner? Nothing to drag: the equal angles put the foot a side over √3 from the corner.
    [Puzzle("two-triangles-in-a-square")]
    public static void TwoTrianglesInASquare(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 6;
        double f = L / Math.Sqrt(3), p = L * f / (L + f);
        var bl = d.Pt(0, 0, "BL");
        var tl = d.Pt(0, L, "TL");
        var tr = d.Pt(L, L, "TR");
        var br = d.Pt(L, 0, "BR");
        var foot = d.Pt(f, 0, "F");
        var cross = d.Pt(p, p, "P");
        d.Polygon(d.Fill("LightBlue", "#FF3FA8DB"), tl, bl, foot);
        d.Polygon(d.Fill("DarkBlue", "#FF27639A"), cross, foot, br);
        d.Polygon(d.Outline("Frame"), bl, br, tr, tl);
        string ink = d.Line("Ink");
        Drawing2022b.Dashed(d, ink, bl, tr, 14);
        // her marks are big: plain arcs sized in units, so they keep her proportions on a phone too
        d.ArcMark(ink, tl, bl, foot, 0.9);
        foreach (double radius in new[] { 0.55, 0.68 })
        {
            d.ArcMark(ink, cross, foot, br, radius);
            d.ArcMark(ink, br, cross, foot, radius + 0.25);
        }

        d.Label("?", V.At(0.42, L - 1.4), d.ShapeText("ShapeMark", Lgf.Ink, 32));
        d.Scene(0, L, L, 0);
    }

    // A regular hexagon cut by the line from its left corner to its top right one and a line from its
    // top side to its bottom left corner; the triangles they cut off have areas 1 and 3. The whole
    // hexagon? Nothing to drag: the areas fix where the second line starts, a sixth of the side left
    // of the top side's middle.
    [Puzzle("subdivided-hexagon-iii")]
    public static void SubdividedHexagonIII(Lgf d)
    {
        const double s = 2, h = 2 * S3;
        var hexagon = HexagonFromLeft(d, V.At(-s, 0), Lgf.N(s));
        var (left, topRight, bottomLeft) = (hexagon[0], hexagon[2], hexagon[5]);
        var x = d.Pt(-s / 6, h, "X");
        var p = d.Crossing(left, topRight, x, bottomLeft, "P");
        d.Polygon(d.Fill("Orange", "#FFF4D2AA", null), x, topRight, p);
        d.Polygon(d.Fill("Purple", "#FFC08DC0", null), left, p, bottomLeft);
        string ink = d.Line("Ink");
        d.Segment(ink, left, topRight);
        d.Segment(ink, x, bottomLeft);
        d.Polygon(d.Outline("Frame"), hexagon);
        string mark = d.ShapeText("ShapeMark", Lgf.Ink, 34);
        d.Label("1", Mid(d, x, topRight, p), mark);
        d.Label("3", Mid(d, left, p, bottomLeft), mark);
        d.Scene(-s, h, s, -h);
    }

    // Three squares, two big ones side by side and a small one on the right one's top right, and two
    // triangles of area 9: one in the left big square under the line from its bottom left corner to the
    // small square's top left corner, one between that line, the small square's diagonal and the big
    // squares' top. The squares' total area? Nothing to drag: equal areas make the big squares' side
    // one and a half times the small one's.
    [Puzzle("three-squares-and-two-triangles")]
    public static void ThreeSquaresAndTwoTriangles(Lgf d)
    {
        const double a = 3, b = 2;
        var o = d.Pt(0, 0, "O");
        var top = d.Pt(0, a, "T");
        var q = d.Pt(a * (2 * a - b) / (a + b), a, "Q");
        var apex = d.Pt(2 * a - b, a + b, "A");
        var r = d.Pt(2 * a, a, "R");
        d.Polygon(d.Fill("Blue", "#FF5693D4", null), top, o, q);
        d.Polygon(d.Fill("LightBlue", "#FF52BBEB", null), q, apex, r);
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        d.Box(frame, V.At(0, 0), V.At(a, a));
        d.Box(frame, V.At(a, 0), V.At(2 * a, a));
        d.Box(frame, V.At(2 * a - b, a), V.At(2 * a, a + b));
        string navy = d.Line("Navy", "#FF1C1C8C", 3);
        d.Segment(navy, o, apex);
        d.Segment(navy, apex, r);
        d.Scene(0, a + b, 2 * a, 0);
    }

    // A right triangle with a red quarter circle in its right-angled corner and a semicircle on its
    // bottom side, both touching the long side, which the touching points cut into three lengths of 2.
    // The quarter circle's area? Nothing to drag: the lengths fix it (the radius 2√2, the semicircle's
    // center halfway along the bottom).
    [Puzzle("a-quarter-circle-and-a-semi-circle-in-a-triangle")]
    public static void AQuarterCircleAndASemiCircleInATriangle(Lgf d)
    {
        d.Texture = "FeltTip";
        double c = Math.Sqrt(2.0 / 3), s = Math.Sqrt(1.0 / 3);
        double w = 6 * c, h = 6 * s, R = 6 * c * s, r = w / 2 * s;
        var a = d.Pt(0, 0, "A");
        var corner = d.Pt(w, 0, "C");
        var b = d.Pt(w, h, "B");
        var quarterTop = d.Pt(w, R);
        var quarterFoot = d.Pt(w - R, 0);
        var m = d.Pt(w / 2, 0, "M");
        var semiRight = d.Pt(w / 2 + r, 0);
        var semiLeft = d.Pt(w / 2 - r, 0);
        d.Arc("CircleSector", d.Fill("Red", "#FFDF0202"), corner, quarterTop, quarterFoot);
        string ink = d.Line("Ink");
        d.Arc("CircleArc", ink, m, semiRight, semiLeft);
        d.Inked(d.Outline("Frame"), a, corner, b);
        var t1 = d.DotAt(V.At(2 * c, 2 * s), "T1");
        var t2 = d.DotAt(V.At(4 * c, 4 * s), "T2");
        string arrow = d.Line("Arrow", Lgf.Ink, 2);
        string mark = d.Text("Mark", Lgf.Ink, 32);
        d.Dimension(arrow, mark, a, t1, 0.36, "2", 0.27);
        d.Dimension(arrow, mark, t1, t2, 0.36, "2", 0.27);
        d.Dimension(arrow, mark, t2, b, 0.36, "2", 0.27);
        d.Scene(-0.2, h + 0.3, w, 0);
    }

    // Four squares, a big one and three small ones in a row beside it, and a yellow triangle from the
    // big one's top left corner to the last small one's bottom side; its long side runs through the
    // first small square's top left corner. What fraction is yellow? Nothing to drag: that corner on
    // the line makes the big square's side one and a half times the small ones'.
    [Puzzle("four-squares-and-a-triangle")]
    public static void FourSquaresAndATriangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double a = 3, b = 2;
        var top = d.Pt(0, a, "T");
        var foot = d.Pt(a + 2 * b, 0, "F");
        var end = d.Pt(a + 3 * b, 0, "E");
        d.Polygon(d.Fill("Yellow", "#FFFEFE14"), top, foot, end);
        string frame = d.Outline("Frame");
        d.Box(frame, V.At(0, 0), V.At(a, a));
        for (int i = 0; i < 3; i++)
        {
            d.Box(frame, V.At(a + i * b, 0), V.At(a + (i + 1) * b, b));
        }

        d.Scene(0, a, a + 3 * b, 0);
    }

    // Two equilateral triangles, a dark one standing on the ground and a green one on a side of it, its
    // far corner on the ground too; the line from the green one's top through the corner they share
    // down to the ground. The angle there? Dragging the dark one's top corner tilts it and resizes the
    // green one: the angle stays.
    [Puzzle("two-equilateral-triangles-iii")]
    public static void TwoEquilateralTrianglesIII(Lgf d)
    {
        d.Texture = "FeltTip";
        var a = d.Pt(0, 0, "A");
        var east = d.Pt(1, 0);
        var b = d.ArcSlider(a, 3, 70, 86, 79.6, (0, 0), "B");
        var p = d.Pt(new V($"0.5 * B.X + {s3} * B.Y", $"-{s3} * B.X + 0.5 * B.Y"), "P");
        var e = d.Crossing(b, p, a, east, "E");
        var c = d.Pt(b + (e - b).Turn(60), "C");
        var f = d.Crossing(c, p, a, east, "F");
        d.Polygon(d.Fill("Teal", "#FF036069"), a, b, p);
        d.Polygon(d.Fill("Green", "#FF3AC653"), b, e, c);
        string ink = d.Line("Ink");
        d.Segment(ink, d.Pt(-1.1, 0), d.Pt(6.2, 0));
        d.Segment(ink, c, f);
        d.ArcMark(ink, f, e, c, 0.38);
        d.Label("?", f + V.At(0.6, 0.42), d.Text("Mark", Lgf.Ink, 28));
        d.Scene(-1.1, 4.5, 6.2, 0);
    }

    // Two squares, the big one of side 12, the small one in the corner beside it; the shaded parts: a
    // quadrilateral from the big one's top corners to the small one's top corners and a thin triangle
    // from the big one's bottom left corner to the small one's top corners. Their area? Dragging the
    // small square's corner resizes it: the shaded area stays half the big square.
    [Puzzle("two-squares-ii")]
    public static void TwoSquaresII(Lgf d)
    {
        const double L = 6;
        var o = d.Pt(0, 0, "O");
        var tl = d.Pt(0, L, "TL");
        var tr = d.Pt(L, L, "TR");
        var s = d.Slider(V.At(L + 1.4, 1.4), V.At(L + 3.6, 3.6), (2.9 - 1.4) / 2.2, (L + 2.9, 2.9), "S");
        var m = d.Pt(new V(Lgf.N(L), "S.Y"), "M");
        string purple = d.Fill("Purple", "#FF9085D6", null);
        d.Polygon(purple, tl, tr, s, m);
        d.Polygon(purple, o, s, m);
        string frame = d.Outline("Frame");
        d.Box(frame, V.At(0, 0), V.At(L, L));
        d.Box(frame, V.At(L, 0), s);
        string line = d.Line("PurpleLine", "#FF4B22A8", 2.5);
        d.Segment(line, tl, m);
        d.Segment(line, tr, s);
        d.Segment(line, o, m);
        d.Segment(line, o, s);
        d.Dimension(d.Line("Arrow"), d.Text("Mark", Lgf.Ink, 36), o, tl, 1, "12", 0.8);
        d.Scene(-1.3, L, L + 3.2, 0);
    }

    // Three squares of area 4 in a row, and a pink right triangle with its right angle at the first
    // one's top left corner, its far corner at the last one's bottom right and its long side through the
    // middle one's top right corner. Its area? Nothing to drag: the squares fix it.
    [Puzzle("three-squares-and-a-triangle")]
    public static void ThreeSquaresAndATriangle(Lgf d)
    {
        d.Texture = "FeltTip";
        var a = d.Pt(0, 2, "A");
        var r = d.Pt(6, 0, "R");
        var t = d.Pt(1, 5, "T");
        d.Polygon(d.Fill("Pink", "#FFF897CE"), a, r, t);
        string frame = d.Outline("Frame");
        for (int i = 0; i < 3; i++)
        {
            d.Box(frame, V.At(2 * i, 0), V.At(2 * i + 2, 2));
        }

        d.Scene(0, 5, 6, 0);
    }

    // Two quarter circles side by side on the ground, meeting at their feet; a line from the yellow
    // one's top to the orange one's corner, and one from the orange one's top to where that line
    // crosses the yellow arc. The angle there? Dragging the point where they meet trades size between
    // them: the angle stays.
    [Puzzle("two-quarter-circles")]
    public static void TwoQuarterCircles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double W = 6;
        var yc = d.Pt(0, 0, "Yc");
        var oc = d.Pt(W, 0, "Oc");
        var g = d.Slider(V.At(2, 0), V.At(3.2, 0), (2.43 - 2) / 1.2, (2.43, 0), "G");
        var y = d.Pt(new V("0", "G.X"), "Y");
        var o = d.Pt(new V(Lgf.N(W), $"{Lgf.N(W)} - G.X"), "Ot");
        // where the line from Y to the orange corner crosses the yellow circle again
        string t = d.Num($"2 * G.X ^ 2 / ({Lgf.N(W * W)} + G.X ^ 2)", "t");
        var p = d.Pt(new V($"{t} * {Lgf.N(W)}", $"G.X * (1 - {t})"), "P");
        d.Arc("CircleSector", d.Fill("Yellow", "#FFFFFE04"), yc, g, y);
        d.Arc("CircleSector", d.Fill("Orange", "#FFFDA20A"), oc, o, g);
        string ink = d.Line("Ink");
        d.Segment(ink, y, oc);
        d.Segment(ink, o, p);
        d.ArcMark(ink, p, oc, o, 0.5);
        d.Label("?", p + V.At(0.85, 0.1), d.Text("Mark", Lgf.Ink, 30));
        d.Scene(0, 3.6, W, 0);
    }

    // A square of area 36 and a right triangle with its right angle at the square's top right corner,
    // a corner at the middle of its bottom side (the ticks) and its long side through its top left
    // corner. The triangle's area? Nothing to drag: the square fixes it.
    [Puzzle("a-square-and-a-triangle")]
    public static void ASquareAndATriangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 6;
        var o = d.Pt(0, 0, "O");
        var m = d.Pt(L / 2, 0, "M");
        var br = d.Pt(L, 0, "BR");
        var tr = d.Pt(L, L, "TR");
        var t = d.Pt(-2, 10, "T");
        d.Polygon(d.Fill("Teal", "#FF02927F"), t, m, tr);
        d.Box(d.Outline("Frame"), V.At(0, 0), V.At(L, L));
        string ink = d.Line("Ink");
        d.Segment(ink, o, m, "TwoTicks");
        d.Segment(ink, m, br, "TwoTicks");
        d.Scene(-2, 10, L, 0);
    }
}

file static class Drawing2022b
{
    /// <summary>A dashed line from a to b, as separate dashes (the dash style's pattern is a few pixels, hers is long): the given number of dashes, a dash at each end</summary>
    public static void Dashed(Lgf d, string style, V a, V b, int dashes, double share = 0.6)
    {
        double period = 1 / (dashes - 1 + share);
        for (int i = 0; i < dashes; i++)
        {
            double from = i * period, to = Math.Min(1, i * period + share * period);
            d.Segment(style, d.Pt(a + from * (b - a)), d.Pt(a + to * (b - a)));
        }
    }
}
