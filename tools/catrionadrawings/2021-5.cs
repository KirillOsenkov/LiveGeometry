#nullable disable

static partial class Puzzles
{
    // A rectangle 5 high, its diagonal from the top left corner to the bottom right one, and two green
    // semicircles: one on the top side ending at the top right corner, one on the bottom side starting
    // at the bottom left corner, each touching the diagonal at a dot; the dots cut the diagonal in three
    // equal parts. The total shaded area? Nothing to drag: touching at the thirds fixes the width (5√5/2)
    // and both radii (√5).
    [Puzzle("semi-circles-in-a-rectangle")]
    public static void SemiCirclesInARectangle(Lgf d)
    {
        const double H = 5;
        double w = 2.5 * Math.Sqrt(5), r = Math.Sqrt(5);
        var bl = d.Pt(0, 0);
        var br = d.Pt(w, 0);
        var tr = d.Pt(w, H);
        var tl = d.Pt(0, H);
        string green = d.Fill("Green", "#FF6DC05C", "#FF1A7F42");
        d.Arc("CircleSegment", green, d.Pt(w - r, H), d.Pt(w - 2 * r, H), tr);
        d.Arc("CircleSegment", green, d.Pt(r, 0), d.Pt(2 * r, 0), bl);
        string teal = d.Line("Teal", "#FF14967A");
        Draw2021.Around(d, teal, bl, br, tr, tl);
        d.Segment(teal, tl, br);
        foreach (var (x, y) in new[] { (0, H), (w, H), (w, 0), (w / 3, H * 2 / 3), (w * 2 / 3, H / 3) })
        {
            d.DotAt(V.At(x, y));
        }

        d.Dimension(d.Line("Arrow"), d.Text("Mark", Lgf.Ink, 36), bl, tl, 0.7, "5", 0.55);
        d.Scene(-1.1, H, w, 0);
    }

    // A regular hexagon with a purple triangle (area 6) from its left corner to its bottom left corner
    // and a point inside, and two equilateral red triangles of the same size: one from that point to
    // the top right corner and a third point on the horizontal middle line, the other from there to the
    // bottom right corner, poking out of the right side. The hexagon's area? Dragging the red triangles'
    // shared corner along the middle line swings both: the purple triangle's top follows and its area
    // stays a sixth of the hexagon.
    [Puzzle("three-triangles-in-a-hexagon")]
    public static void ThreeTrianglesInAHexagon(Lgf d)
    {
        var hexagon = HexagonFromLeft(d, V.At(-2, 0), "2");
        var (left, topRight, bottomRight, bottomLeft) = (hexagon[0], hexagon[2], hexagon[4], hexagon[5]);
        var q = d.Slider(V.At(0.1, 0), V.At(1.1, 0), 0.45, (0.55, 0), "Q");
        var p = d.Pt(q + (topRight - q).Turn(60), "P");
        var x = d.Pt(q + (bottomRight - q).Turn(60), "X");
        string purple = "#FF4A2A9A";
        d.Polygon(d.Fill("Purple", "#FFAC98DF", purple), left, p, bottomLeft);
        string red = d.Fill("Red", "#FFE46D83", "#FFD03A4C");
        d.Polygon(red, p, topRight, q);
        d.Polygon(red, q, x, bottomRight);
        // inked after coloring: the hexagon's side runs over the red triangle that pokes out
        d.Inked(d.Outline("Frame", "#FF3A3F4D", 3), hexagon);
        d.Label("6", Mid(d, left, p, bottomLeft), d.ShapeText("PurpleMark", purple, 34));
        d.Scene(-2, 2 * S3, 2.3, -2 * S3);
    }

    // A semicircle with seven equally spaced dots on its arc (the ends included): the chord between the
    // second and the sixth is 3. Shaded: the cap above the chord between the third and the fifth, and
    // the part below the long chord. The total shaded area? Nothing to drag: the chord fixes the radius
    // (√3).
    [Puzzle("striped-semi-circle")]
    public static void StripedSemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        double r = Math.Sqrt(3);
        var o = d.Pt(0, 0, "O");
        var p = Enumerable.Range(0, 7).Select(i => d.Pt(r * Math.Cos(i * Math.PI / 6), r * Math.Sin(i * Math.PI / 6), "P" + i)).ToArray();
        string blue = d.Fill("Blue", "#FF019AFE", null);
        d.Arc("CircleSegment", blue, o, p[2], p[4]);
        d.Region(blue, "Difference", d.Arc("CircleSegment", null, o, p[0], p[6], visible: false), d.Arc("CircleSegment", null, o, p[1], p[5], visible: false));
        string navy = "#FF3C4171";
        d.Arc("CircleSegment", d.Outline("Frame", navy, 3), o, p[0], p[6]);
        string line = d.Line("Navy", navy);
        d.Segment(line, p[1], p[5]);
        d.Segment(line, p[2], p[4]);
        foreach (var point in p)
        {
            d.DotAt(point);
        }

        string arrow = "#FF1844EA";
        d.Dimension(d.Line("Arrow", arrow), d.Text("BlueMark", arrow, 30), p[5], p[1], 0.08, "3", 0.24);
        d.Scene(-r, r, r, 0);
    }

    // A line 4 long with a semicircle on it, and a quarter circle around each end: all three arcs meet
    // in one point. Shaded: the left quarter circle's part outside the rest (purple), the right one's
    // (yellow), and where the two quarter circles overlap (orange). The total shaded area? Dragging the
    // meeting point along the semicircle trades size between the quarter circles: their radii's squares
    // still add up to 16, and the total stays 2π.
    [Puzzle("two-quarter-circles-and-a-semi-circle-ii")]
    public static void TwoQuarterCirclesAndASemiCircleII(Lgf d)
    {
        var a = d.Pt(0, 0, "A");
        var b = d.Pt(4, 0, "B");
        var m = d.Pt(2, 0, "M");
        var p = d.ArcSlider(m, 2, 62, 118, 82, (2, 0), "P");
        string ra = d.Num("distance(A, P)", "ra");
        string rb = d.Num("distance(B, P)", "rb");
        var aEnd = d.Pt(new V(ra, "0"));
        var aTop = d.Pt(new V("0", ra));
        var bTop = d.Pt(new V("4", rb));
        var bEnd = d.Pt(new V($"4 - {rb}", "0"));
        string left = d.Arc("CircleSector", null, a, aEnd, aTop, visible: false);
        string right = d.Arc("CircleSector", null, b, bTop, bEnd, visible: false);
        string semicircle = d.Arc("CircleSegment", null, m, b, a, visible: false);
        d.Region(d.Fill("Pink", "#FFE47C96", null), "Difference", left, semicircle, right);
        d.Region(d.Fill("Yellow", "#FFFEF361", null), "Difference", right, semicircle, left);
        d.Region(d.Fill("Orange", "#FFFE9F4D", null), "Intersection", left, right);
        string ink = d.Line("Ink");
        d.Arc("CircleArc", ink, a, aEnd, aTop);
        d.Arc("CircleArc", ink, b, bTop, bEnd);
        d.Arc("CircleArc", ink, m, b, a);
        d.Segment(ink, a, b);
        d.Segment(ink, a, aTop);
        d.Segment(ink, b, bTop);
        d.Dimension(ink, d.Text("Mark", Lgf.Ink, 30), a, b, -0.4, "4", 0.3);
        d.Scene(0, 3.45, 4, -0.5);
    }

    // Five quarter circles standing on one line in three columns: in the first, a small quarter circle
    // filling it and a bigger one around the same corner, which reaches the second column's left side
    // at the height of the second column's own small quarter circle; in the second, that one and a big
    // one, which reaches the third column's side at the height of the last quarter circle. Pink: the
    // second column's part of its big quarter circle outside the first one's big one. Purple: the first
    // column's band and the last quarter circle's part outside the second one's big one. More pink or
    // purple? Dragging the tops of the first and the second column changes every size but the first
    // column's width: the two colors stay equal.
    [Puzzle("five-quarter-circles")]
    public static void FiveQuarterCircles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double A = 2;
        var o = d.Pt(0, 0, "O");
        var t2 = d.Slider(V.At(0, 1.3 * A), V.At(0, 1.62 * A), 0.47, (0, 2.9), "T2");
        string w2 = d.Num($"sqrt(T2.Y ^ 2 - {A * A})", "w2");
        var t3 = d.Slider(new V(Lgf.N(A), $"1.45 * {w2}"), new V(Lgf.N(A), $"1.95 * {w2}"), 0.59, (A, 3.67), "T3");
        string r5 = d.Num($"sqrt(T3.Y ^ 2 - {w2} ^ 2)", "r5");
        string c = $"{Lgf.N(A)} + {w2}";
        var b = d.Pt(A, 0, "B");
        var cc = d.Pt(new V(c, "0"), "C");
        var q1End = d.Pt(A, 0);
        var q1Top = d.Pt(0, A);
        var q2End = d.Pt(new V("T2.Y", "0"));
        var q3End = d.Pt(new V($"{Lgf.N(A)} + T3.Y", "0"));
        var q4End = d.Pt(new V(c, "0"));
        var q4Top = d.Pt(new V(Lgf.N(A), w2));
        var q5End = d.Pt(new V($"{c} + {r5}", "0"));
        var q5Top = d.Pt(new V(c, r5));
        string q1 = d.Arc("CircleSector", null, o, q1End, q1Top, visible: false);
        string q2 = d.Arc("CircleSector", null, o, q2End, t2, visible: false);
        string q3 = d.Arc("CircleSector", null, b, q3End, t3, visible: false);
        string q5 = d.Arc("CircleSector", null, cc, q5End, q5Top, visible: false);
        // what lies right of a column, cut off its quarter circle
        string pastFirst = d.Polygon(null, b, d.Pt(new V("T2.Y + 1", "0")), d.Pt(new V("T2.Y + 1", "T2.Y")), d.Pt(new V(Lgf.N(A), "T2.Y")));
        string pastSecond = d.Polygon(null, cc, d.Pt(new V($"{c} + T3.Y", "0")), d.Pt(new V($"{c} + T3.Y", "T3.Y")), d.Pt(new V(c, "T3.Y")));
        string purple = d.Fill("Purple", "#FFAD8BDF", null);
        d.Region(purple, "Difference", q2, q1, pastFirst);
        d.Region(d.Fill("Pink", "#FFFEB8CF", null), "Difference", q3, q2, pastSecond);
        d.Region(purple, "Difference", q5, q3);
        string ink = d.Line("Ink", Lgf.Ink, 3);
        d.Arc("CircleArc", ink, o, q1End, q1Top);
        d.Arc("CircleArc", ink, o, q2End, t2);
        d.Arc("CircleArc", ink, b, q3End, t3);
        d.Arc("CircleArc", ink, b, q4End, q4Top);
        d.Arc("CircleArc", ink, cc, q5End, q5Top);
        d.Segment(ink, o, t2);
        d.Segment(ink, b, t3);
        d.Segment(ink, cc, q5Top);
        d.Segment(ink, o, q5End);
        d.Scene(0, 3.7, 7.1, 0);
    }

    // A regular hexagon cut by a line from its top left corner to a point on its bottom side and one from
    // there to its right corner: the dark green part is 2, the yellow 1, the light green? Nothing to
    // drag: the two areas put the point a quarter of the side from the bottom left corner.
    [Puzzle("subdivided-hexagon")]
    public static void SubdividedHexagon(Lgf d)
    {
        d.Texture = "FeltTip";
        var hexagon = HexagonFromLeft(d, V.At(-2, 0), "2");
        var (left, topLeft, topRight, right, bottomRight, bottomLeft) = (hexagon[0], hexagon[1], hexagon[2], hexagon[3], hexagon[4], hexagon[5]);
        var p = d.Pt(-0.5, -2 * S3, "P");
        d.Polygon(d.Fill("DarkGreen", "#FF5DC270", null), left, topLeft, p, bottomLeft);
        d.Polygon(d.Fill("LightGreen", "#FFC0F142", null), topLeft, topRight, right, p);
        d.Polygon(d.Fill("Yellow", "#FFFEFC05", null), p, bottomRight, right);
        string ink = d.Line("Ink", Lgf.Ink, 4);
        d.Segment(ink, topLeft, p);
        d.Segment(ink, p, right);
        d.Polygon(d.Outline("Frame", Lgf.Ink, 4), hexagon);
        string mark = d.ShapeText("ShapeMark", Lgf.Ink, 40);
        d.Label("2", V.At(-1.15, -0.2), mark);
        d.Label("?", V.At(0.35, 0.35), mark);
        d.Label("1", V.At(0.85, -1.3), mark);
        d.Scene(-2, 2 * S3, 2, -2 * S3);
    }

    // A red semicircle on a diameter of 6, a blue line from above its left end (straight up from it) to
    // its right end, and blue semicircles on that line: a big one on all of it through the red
    // diameter's left end, a small white one cut out of it up to where the red arc crosses the line, and
    // one on the rest on the other side. The total shaded area? Dragging the top of the blue line up and
    // down changes all three: the shaded area stays 9π (π/8 times twice the red diameter's square).
    [Puzzle("multiple-semi-circles-iii")]
    public static void MultipleSemiCirclesIII(Lgf d)
    {
        var a = d.Pt(-3, 0, "A");
        var e = d.Pt(3, 0, "E");
        var o = d.Pt(0, 0, "O");
        var x = d.Slider(V.At(-3, 2.5), V.At(-3, 5.5), (4.33 - 2.5) / 3, (-3, 4.33), "X");
        string s = d.Num("X.Y ^ 2 / (36 + X.Y ^ 2)", "s");
        var m = d.Pt(new V($"-3 + 6 * {s}", $"X.Y - X.Y * {s}"), "M");
        var big = Mid(d, x, e);
        var small = Mid(d, x, m);
        var middle = Mid(d, m, e);
        string none = null;
        string outer = d.Region(null, "Union", d.Arc("CircleSegment", none, big, x, e, visible: false), d.Arc("CircleSegment", none, middle, e, m, visible: false));
        d.Region(d.Fill("Blue", "#FFA7DBFD", null), "Difference", outer, d.Arc("CircleSegment", none, small, x, m, visible: false));
        string blue = d.Line("BlueLine", "#FF1F6BC4", 3);
        d.Arc("CircleArc", blue, big, x, e);
        d.Arc("CircleArc", blue, small, x, m);
        d.Arc("CircleArc", blue, middle, e, m);
        d.Segment(blue, x, e);
        string plum = "#FF803A71";
        string red = d.Line("Plum", plum, 3);
        d.Arc("CircleArc", red, o, e, a);
        d.Segment(red, a, e);
        d.Label("6", V.At(0, -0.6), d.ShapeText("PlumMark", plum, 34));
        d.Scene(-3.8, 4.4, 3.8, -1.6);
    }

    // A blue semicircle on a line 4 long; a yellow semicircle on the chord from its left end to the point
    // of the arc straight above 3 along, hanging below the chord (its arc crosses the diameter there),
    // and a red semicircle on the same center on the other side, touching the blue arc. The total shaded
    // area? Nothing to drag: the 3 and the 1 fix it (the yellow radius √3, the red one 1).
    [Puzzle("three-overlapping-semi-circles")]
    public static void ThreeOverlappingSemiCircles(Lgf d)
    {
        d.Texture = "FeltTip";
        double h = Math.Sqrt(3);
        var l = d.Pt(0, 0, "L");
        var k = d.Pt(3, 0, "K");
        var b = d.Pt(4, 0, "B");
        var o = d.Pt(2, 0, "O");
        var q = d.Pt(3, h, "Q");
        var m = d.Pt(1.5, h / 2, "M");
        d.Arc("CircleSegment", d.Fill("Yellow", "#FFFFF57F", null), m, l, q);
        d.Arc("CircleSegment", d.Fill("Red", "#FFE37485", null), m, d.Pt(1.5 + S3, h / 2 + 0.5), d.Pt(1.5 - S3, h / 2 - 0.5));
        d.Arc("CircleSegment", d.Outline("Frame", "#FF2F5FC0", 3), o, b, l);
        string blue = "#FF4F7FCC";
        string arrow = d.Line("Arrow", blue, 2.5, dark: "#FF1E3C80");
        string mark = d.Text("BlueMark", blue, 28, dark: "#FF1E3C80");
        d.Dimension(arrow, mark, l, k, -0.2, "3", 0.3);
        d.Dimension(d.Line("ArrowOut", blue), d.Text("BlueOut", blue, 28), k, b, -0.2, "1", 0.15);
        d.Scene(0, 2, 4, -0.9);
    }

    // Three quarter circles standing on one line: a big one around its right end through its left end,
    // and two of the same size around its two ends. The line from the right end to where the big arc
    // meets the left small one runs through where the small ones cross. The angle at the right end?
    // Nothing to drag: that crossing fixes the small radius (the big one over the golden ratio).
    [Puzzle("three-quarter-circles")]
    public static void ThreeQuarterCircles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 4;
        double r = L * (Math.Sqrt(5) - 1) / 2, px = r * r / (2 * L);
        var a = d.Pt(0, 0, "A");
        var b = d.Pt(L, 0, "B");
        var bigTop = d.Pt(L, L);
        var leftEnd = d.Pt(r, 0);
        var leftTop = d.Pt(0, r);
        var rightTop = d.Pt(L, r);
        var rightEnd = d.Pt(L - r, 0);
        var p = d.Pt(px, Math.Sqrt(r * r - px * px), "P");
        d.Region(d.Fill("Pink", "#FFFEB4D6", null), "Union", d.Arc("CircleSector", null, b, bigTop, a, visible: false), d.Arc("CircleSector", null, a, leftEnd, leftTop, visible: false));
        string magenta = d.Line("Magenta", "#FFB22FA8", 3);
        d.Arc("CircleArc", magenta, b, bigTop, a);
        d.Arc("CircleArc", magenta, a, leftEnd, leftTop);
        d.Arc("CircleArc", magenta, b, rightTop, rightEnd);
        d.Segment(magenta, a, b);
        d.Segment(magenta, a, leftTop);
        d.Segment(magenta, b, bigTop);
        string ink = d.Line("Ink");
        d.Segment(ink, p, b);
        d.Angle(d.Line("AngleMark", Lgf.Ink, 3), b, p, a, 30);
        d.Label("?", b + V.At(-0.3, 0.62), d.ShapeText("Mark", Lgf.Ink, 24));
        // room below: the page's buttons sit over the bottom right corner, where the angle is
        d.Scene(0, L, L, -0.45);
    }

    // A regular nonagon and a regular hexagon inside it on the same bottom side; the shaded part between
    // them. The line from the hexagon's right corner through the nonagon's corner after the bottom side
    // meets the bottom line 2 from that corner. The shaded part's perimeter? Nothing to drag: the angles
    // make that 2 the common side.
    [Puzzle("a-hexagon-and-a-nonagon")]
    public static void AHexagonAndANonagon(Lgf d)
    {
        const double s = 2;
        V[] Regular(int sides)
        {
            var corners = new List<V>();
            double x = -s / 2, y = 0;
            for (int i = 0; i < sides; i++)
            {
                corners.Add(d.Pt(x, y));
                x += s * Math.Cos(2 * Math.PI * i / sides);
                y += s * Math.Sin(2 * Math.PI * i / sides);
            }

            return corners.ToArray();
        }

        var nonagon = Regular(9);
        var hexagon = Regular(6);
        string ink = Lgf.Ink;
        d.Polygon(d.Fill("Green", "#FF97DC7B", ink, 3.5), nonagon);
        d.Polygon(d.PaperFill("Hole", ink, 3.5), hexagon);
        var e = d.Pt(s / 2 + 2 * s * Math.Cos(40 * Math.PI / 180), 0, "E");
        string green = "#FF1FA37E";
        string line = d.Line("GreenLine", green, 2.5);
        d.Segment(line, hexagon[2], nonagon[2]);
        d.Segment(line, nonagon[2], e);
        d.Segment(line, nonagon[1], e);
        d.Dimension(d.Line("Arrow", "#FF3CC2A0", 2), d.Text("GreenMark", green, 26), nonagon[2], e, 0.35, "2", 0.25);
        // the nonagon's circumradius: its top corner is that plus the inradius above the bottom side
        double radius = s / (2 * Math.Sin(Math.PI / 9));
        d.Scene(-radius * Math.Cos(Math.PI / 18), radius * (1 + Math.Cos(Math.PI / 9)), s / 2 + 2 * s * Math.Cos(40 * Math.PI / 180) + 0.15, 0);
    }

    // A circle and four quarter circles around one point inside it, like a pinwheel: each has its corner
    // straight across from that point and the far end of its arc on the circle, on a diagonal through
    // the point. What fraction of the circle do they cover? Dragging the point inside the circle changes
    // all four: the squares of the radii always add up to twice the circle's, so they cover half.
    [Puzzle("four-quarter-circles-in-a-circle")]
    public static void FourQuarterCirclesInACircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4;
        var o = d.Pt(0, 0, "O");
        var rim = d.Pt(R, 0, "Rim");
        var reach = d.Pt(0.45 * R, 0, "Reach");
        string area = d.Circle(null, o, reach);
        var c = d.HandleIn(area, -0.33, 0.89, "C");
        string k = d.Num($"C.X ^ 2 + C.Y ^ 2 - {R * R}", "k");
        var colors = new[] { ("Blue", "#FF0193FE", -1, 1), ("Yellow", "#FFFCEA01", 1, 1), ("Red", "#FFEE6A81", 1, -1), ("Purple", "#FF9F81CE", -1, -1) };
        foreach (var (name, color, ex, ey) in colors)
        {
            string b = $"({ex} * C.X + {ey} * C.Y)";
            string r = d.Num($"(sqrt({b} ^ 2 - 2 * {k}) - {b}) / 2", "r" + name);
            var end = d.Pt(c + new V($"{ex} * {r}", $"{ey} * {r}"));
            // the corner is across from the point horizontally on the left and right ones, vertically on the others
            var corner = ex * ey < 0 ? d.Pt(c + new V($"{ex} * {r}", "0")) : d.Pt(c + new V("0", $"{ey} * {r}"));
            // each turns counterclockwise from the point to the far end
            d.Arc("CircleSector", d.Fill(name, color, Lgf.Ink, 2), corner, c, end);
        }

        d.Circle(d.Outline("Circle", Lgf.Ink, 3.5), o, rim);
        d.Scene(-R, R, R, -R);
    }
}

file static class Draw2021
{
    /// <summary>A closed outline as segments, one per side: a polygon's sides in a line style, drawn over arcs</summary>
    public static void Around(Lgf d, string lineStyle, params V[] points)
    {
        for (int i = 0; i < points.Length; i++)
        {
            d.Segment(lineStyle, points[i], points[(i + 1) % points.Length]);
        }
    }
}
