#nullable disable

static partial class Puzzles
{
    // Five squares: three of side 5 in an L (two on the bottom, one on the right one), and a pink
    // rectangle of two squares with a corner at the L's bottom left corner, its top side through the
    // L's top right corner, the side the two pink squares share starting where the column's left side
    // crosses that top side. The total shaded area? Nothing to drag: that fixes the tilt at 15°.
    [Puzzle("five-squares")]
    public static void FiveSquares(Lgf d)
    {
        d.Texture = "FeltTip";
        const double s = 5;
        double h = Math.Sqrt(50), c = Math.Cos(Math.PI / 12), n = Math.Sin(Math.PI / 12);
        var a = d.Pt(0, 0, "A");
        var b = d.Pt(2 * h * c, 2 * h * n, "B");
        var e = d.Pt(2 * h * c - h * n, 2 * h * n + h * c, "C");
        var f = d.Pt(-h * n, h * c, "D");
        d.Polygon(d.Fill("Pink", "#FFFEBACE", Lgf.Ink, 3), a, b, e, f);
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        d.Box(frame, V.At(0, 0), V.At(s, s));
        d.Box(frame, V.At(s, 0), V.At(2 * s, s));
        d.Box(frame, V.At(s, s), V.At(2 * s, 2 * s));
        string ink = d.Line("Ink", Lgf.Ink, 3);
        d.Segment(ink, d.Pt(h * c, h * n), d.Pt(h * c - h * n, h * n + h * c));
        d.Dimension(d.Line("Arrow"), d.Text("Mark", Lgf.Ink, 30), d.Pt(s, 2 * s), d.Pt(2 * s, 2 * s), 0.75, "5", 0.45);
        d.Scene(-h * n, 2 * s + 1.1, 2 * h * c, 0);
    }

    // Five congruent rectangles in a semicircle of diameter 10: one lying on the diameter from its left
    // end, one standing on it, one lying on top of that from its right side, the two top corners on the
    // arc (which fixes the rectangles: L = (5 + √5) / 2 by w = (3√5 - 5) / 2); two more leaning like
    // dominoes, a foot on the diameter, the first against the standing one, the second against the
    // first. Their total area? Dragging the leaning ones' feet tilts them: the rectangles stay.
    [Puzzle("five-rectangles-in-a-semi-circle")]
    public static void FiveRectanglesInASemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 5;
        double r5 = Math.Sqrt(5), L = (5 + r5) / 2, w = (3 * r5 - 5) / 2, xv = -R + L;
        string blue = "#FF1AA8DE";
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0, "E");
        var west = d.Pt(-R, 0, "W");
        string orange = d.Fill("Orange", "#FFFEDCA8", Lgf.Ink, 3);
        V[] Rectangle(double x0, double y0, double x1, double y1) => [d.Pt(x0, y0), d.Pt(x1, y0), d.Pt(x1, y1), d.Pt(x0, y1)];
        // over the semicircle's diameter, which they stand on (a polygon is drawn under every arc)
        d.Inked(orange, Rectangle(-R, 0, xv, w));
        d.Inked(orange, Rectangle(xv - w, w, xv, w + L));
        d.Inked(orange, Rectangle(xv, L, xv + L, L + w));
        string l = Lgf.N(L), wn = Lgf.N(w);
        // the first domino: its foot P on the diameter, its top left corner Q against the standing one
        var p = d.Slider(V.At(0.75, 0), V.At(1.9, 0), (1.4 - 0.75) / 1.15, (1.4, 0), "P");
        var q = d.Pt(new V(Lgf.N(xv), $"sqrt({l} ^ 2 - (P.X - {Lgf.N(xv)}) ^ 2)"), "Q");
        string ux = d.Num($"(Q.X - P.X) / {l}", "ux"), uy = d.Num($"(Q.Y - P.Y) / {l}", "uy");
        var pn = d.Pt(p + new V($"{wn} * {uy}", $"-{wn} * {ux}"), "Pn");
        var qn = d.Pt(q + new V($"{wn} * {uy}", $"-{wn} * {ux}"), "Qn");
        d.Inked(orange, p, pn, qn, q);
        // the second: its foot K on the diameter, its top left corner S on the first one's upper side
        var k = d.Slider(new V("P.X + 2.2", "0"), new V("min(P.X + 3.2, 4.4)", "0"), 0.66, (4.13, 0), "K");
        string b = d.Num($"(Pn.X - K.X) * {ux} + (Pn.Y - K.Y) * {uy}", "b");
        string c = d.Num($"(Pn.X - K.X) ^ 2 + (Pn.Y - K.Y) ^ 2 - {l} ^ 2", "c");
        string t = d.Num($"-{b} + sqrt({b} ^ 2 - {c})", "t");
        var sp = d.Pt(pn + t * new V(ux, uy), "S");
        string vx = d.Num($"(S.X - K.X) / {l}", "vx"), vy = d.Num($"(S.Y - K.Y) / {l}", "vy");
        var kn = d.Pt(k + new V($"{wn} * {vy}", $"-{wn} * {vx}"), "Kn");
        var sn = d.Pt(sp + new V($"{wn} * {vy}", $"-{wn} * {vx}"), "Sn");
        d.Inked(orange, k, kn, sn, sp);
        d.Arc("CircleSegment", d.Outline("Semicircle", blue, 3), o, east, west);
        d.Dimension(d.Line("Arrow", blue), d.Text("BlueMark", blue, 28), west, east, -0.55, "10", 0.55);
        d.Scene(-R, R, R, -0.8);
    }

    // Two squares stacked and an equilateral triangle with its top at their top left corner and its
    // base on their bottom line; a line from their bottom left corner to where the triangle's right
    // side leaves them. The angle there? Nothing to drag: the squares fix the triangle.
    [Puzzle("two-squares-and-an-equilateral-triangle-ii")]
    public static void TwoSquaresAndAnEquilateralTriangleII(Lgf d)
    {
        const double s = 2;
        double half = 2 * s / Math.Sqrt(3);
        string teal = "#FF1F6B4E";
        var o = d.Pt(0, 0, "O");
        var a = d.Pt(0, 2 * s, "A");
        var left = d.Pt(-half, 0, "L");
        var right = d.Pt(half, 0, "R");
        var x = d.Pt(s, 2 * s - Math.Sqrt(3) * s, "X");
        d.Polygon(d.Fill("DarkGreen", "#FF66B887", teal), left, right, a);
        string light = d.Fill("LightGreen", "#FFB0F07A", "#FF5E9E3C");
        d.Box(light, V.At(0, 0), V.At(s, s));
        d.Box(light, V.At(0, s), V.At(s, 2 * s));
        // inked after coloring: the triangle's side runs over the squares
        d.Polygon(d.Outline("Triangle", teal), left, right, a);
        string ink = "#FF2B2622";
        string line = d.Line("Ink", ink, 2.5, dark: ink);
        d.Segment(line, o, x);
        d.Angle(line, x, o, a, 30);
        d.Label("?", x + V.At(-0.62, 0.32), d.ShapeText("Mark", ink, 32));
        d.Scene(-half, 2 * s, half, 0);
    }

    // A regular hexagon (a corner at the top), every side halved (the ticks), and lines joining the
    // middles of the sides two apart; what fraction is shaded? Nothing to drag: the hexagon decides it.
    [Puzzle("subdivided-hexagon-ii")]
    public static void SubdividedHexagonII(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 2;
        var corners = Enumerable.Range(0, 6).Select(i => d.Pt(R * Math.Cos((90 + 60 * i) * Math.PI / 180), R * Math.Sin((90 + 60 * i) * Math.PI / 180), "C" + i)).ToArray();
        var middles = Enumerable.Range(0, 6).Select(i => Mid(d, corners[i], corners[(i + 1) % 6])).ToArray();
        // corners: top, upper left, lower left, bottom, lower right, upper right; middles[i] after corners[i]
        var top = d.Pt(0, R / 2, "Xt");
        var bottom = d.Pt(0, -R / 2, "Xb");
        d.Polygon(d.Fill("Blue", "#FF40A1FD", null), corners[1], middles[0], top, middles[1]);
        d.Polygon(d.Fill("Yellow", "#FFFFFF64", null), middles[5], corners[5], middles[4], top);
        d.Polygon(d.Fill("Red", "#FFEA596D", null), middles[1], bottom, middles[2], corners[2]);
        d.Polygon(d.Fill("Green", "#FF74C057", null), bottom, middles[4], corners[4], middles[3]);
        string ink = d.Line("Ink", Lgf.Ink, 3);
        d.Segment(ink, middles[0], middles[4]);
        d.Segment(ink, middles[5], middles[1]);
        d.Segment(ink, middles[1], middles[3]);
        d.Segment(ink, middles[4], middles[2]);
        for (int i = 0; i < 6; i++)
        {
            d.Segment(ink, corners[i], middles[i], "TwoTicks");
            d.Segment(ink, middles[i], corners[(i + 1) % 6], "TwoTicks");
        }

        d.Scene(-R * S3, R, R * S3, -R);
    }

    // Four congruent right triangles in a row on a line, legs alternately standing and lying (a up b
    // across, b across a up, b up a across, a up b across), a line from the first one's top to the
    // last one's; the angle where it crosses the third one's long side? Dragging the first one's top
    // changes the short leg: the triangles stay congruent and the angle stays.
    [Puzzle("four-congruent-triangles")]
    public static void FourCongruentTriangles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double b = 4;
        var p = d.Slider(V.At(1.8, b), V.At(3.4, b), 0.75, (3, b), "P");
        string a = "P.X", bn = Lgf.N(b);
        var o = d.Pt(0, 0, "O");
        var a1 = d.Pt(new V(a, "0"), "A1");
        var a2 = d.Pt(new V($"{a} + {bn}", "0"), "A2");
        var a3 = d.Pt(new V($"2 * {a} + {bn}", "0"), "A3");
        var a4 = d.Pt(new V($"2 * {a} + 2 * {bn}", "0"), "A4");
        var q = d.Pt(new V($"{a} + {bn}", a), "Q");
        var t = d.Pt(new V($"{a} + {bn}", bn), "T");
        var k = d.Pt(new V($"2 * {a} + {bn}", a), "K");
        d.Polygon(d.Fill("Purple", "#FFD5B6F1", Lgf.Ink, 3), o, a1, p);
        d.Polygon(d.Fill("Pink", "#FFFEC3D2", Lgf.Ink, 3), a1, a2, q);
        d.Polygon(d.Fill("Orange", "#FFFEDCB2", Lgf.Ink, 3), a2, a3, t);
        d.Polygon(d.Fill("Brown", "#FFEEA288", Lgf.Ink, 3), a3, a4, k);
        string purple = "#FF7A4CCB";
        string line = d.Line("PurpleLine", purple, 3);
        d.Segment(line, p, k);
        var x = d.Crossing(p, k, t, a3, "X");
        d.Angle(line, x, k, t, 30);
        d.Label("?", x + V.At(0.6, 0.8), d.Text("PurpleMark", purple, 30));
        d.Scene(0, b + 0.5, 2 * 3 + 2 * b, 0);
    }

    // Two circles, the blue one's center on the green one; a triangle from a point on the green
    // circle to the circles' lower crossing to a point on the blue one, its third side through their
    // upper crossing. The angle at the lower crossing is 50°, the one on the blue circle? The 50°
    // fixes the circles (the upper and lower crossing are 100° apart seen from the blue center).
    // Dragging the triangle's corner on the green circle swings the line through the upper crossing:
    // the angles stay.
    [Puzzle("two-circles")]
    public static void TwoCircles(Lgf d)
    {
        d.Texture = null;
        const double r = 4;
        V OnGreen(double degrees, string name) => d.Pt(r * Math.Cos(degrees * Math.PI / 180), r * Math.Sin(degrees * Math.PI / 180), name);
        var g = d.Pt(0, 0, "G");
        var gRim = OnGreen(180, "Gr");
        var t = OnGreen(58, "T");
        var b = OnGreen(258, "B");
        var o = OnGreen(338, "O");
        var a = d.ArcSlider(g, r, 75, 135, 99, (0, 0), "A");
        string length = d.Num("distance(A, T)", "at");
        string ux = d.Num($"(T.X - A.X) / {length}", "ux"), uy = d.Num($"(T.Y - A.Y) / {length}", "uy");
        string along = d.Num($"-2 * ((T.X - O.X) * {ux} + (T.Y - O.Y) * {uy})", "k");
        var c = d.Pt(t + along * new V(ux, uy), "C");
        // the circles only make up the regions (a region is drawn under every circle)
        string clear = Clear2021(d);
        string green = d.Circle(clear, g, gRim);
        string blue = d.Circle(clear, o, b);
        d.Region(d.Fill("Green", "#FFC7F2AD", null), "Difference", green, blue);
        d.Region(d.Fill("Blue", "#FFA9E3FF", null), "Difference", blue, green);
        d.Region(d.Fill("Lens", "#FFBAE9E1", null), "Intersection", green, blue);
        d.Circle(d.Outline("GreenLine", "#FF79B04A", 3), g, gRim);
        d.Circle(d.Outline("BlueLine", "#FF2A93D5", 3), o, b);
        string ink = d.Line("Ink", "#FF2B2B2B", 2.5, dark: "#FF2B2B2B");
        d.Segment(ink, a, b);
        d.Segment(ink, b, c);
        d.Segment(ink, c, a);
        d.Angle(ink, b, c, a, 32);
        d.Angle(ink, c, a, b, 32);
        d.DotAt(o);
        string mark = d.ShapeText("Mark", "#FF2B2B2B", 30);
        d.Label("50°", InAngle2021(d, b, c, a, 1.6), mark);
        d.Label("?", InAngle2021(d, c, a, b, 1.25), mark);
        d.Scene(-r, r, 3.709 + 5.142, -1.498 - 5.142);
    }

    // Three congruent rectangles: one lying, one standing at its right end, one lying on top of that
    // to the right; lines from the top right corner to the bottom one's top left corner and to the
    // standing one's bottom right corner. The angle between them? The first line runs through the
    // standing one's top left corner, which makes the rectangles 2 by 1: nothing to drag.
    [Puzzle("three-congruent-rectangles-v")]
    public static void ThreeCongruentRectanglesV(Lgf d)
    {
        d.Texture = "FeltTip";
        string orange = d.Fill("Orange", "#FFFE853A", "#FFF04A45", 4);
        d.Box(orange, V.At(0, 0), V.At(2, 1));
        d.Box(orange, V.At(2, 0), V.At(3, 2));
        d.Box(orange, V.At(2, 2), V.At(4, 3));
        var corner = d.Pt(4, 3, "K");
        var left = d.Pt(0, 1, "A");
        var foot = d.Pt(3, 0, "F");
        string ink = d.Line("Ink", Lgf.Ink, 4);
        d.Segment(ink, left, corner);
        d.Segment(ink, foot, corner);
        d.Angle(d.Line("Mark", Lgf.Ink, 3), corner, left, foot, 40);
        d.Label("?", V.At(3.42, 2.45), d.ShapeText("QuestionMark", Lgf.Ink, 34));
        d.Scene(0, 3, 4, 0);
    }

    // Two congruent rectangles, the upper one standing on the lower one's top, half a length to the
    // right, and a tilted square on the lower one's bottom left corner and the upper one's top left
    // and bottom right corners; what fraction of the square is shaded (its part outside both
    // rectangles)? Nothing to drag: the square makes the rectangles 2 by 1.
    [Puzzle("two-rectangles-and-a-square")]
    public static void TwoRectanglesAndASquare(Lgf d)
    {
        d.Texture = "FeltTip";
        var o = d.Pt(0, 0, "O");
        var bottom = d.Pt(2, -1, "B");
        var right = d.Pt(3, 1, "R");
        var top = d.Pt(1, 2, "T");
        string square = d.Polygon(null, o, bottom, right, top);
        string lower = d.Polygon(null, o, d.Pt(2, 0), d.Pt(2, 1), d.Pt(0, 1));
        string upper = d.Polygon(null, d.Pt(1, 1), d.Pt(3, 1), d.Pt(3, 2), top);
        d.Region(d.Fill("Purple", "#FFAA8ACD", null), "Difference", square, lower, upper);
        string purple = d.Outline("Rectangle", "#FF4A2FA0", 4);
        d.Box(purple, V.At(0, 0), V.At(2, 1));
        d.Box(purple, V.At(1, 1), V.At(3, 2));
        d.Polygon(d.Outline("Square", Lgf.Ink, 4), o, bottom, right, top);
        d.Scene(0, 2, 3, -1);
    }

    // Three squares: a big upright one, a green one turned on its top right corner whose left side
    // runs through the big one's bottom left corner, and a yellow one on that corner with its right
    // side along the same line. A line from the green one's top left corner through the big one's
    // top left corner to the yellow one's top side; the angle there? Dragging where the green one's
    // side crosses the big one's top turns the green square, dragging the yellow one's top left
    // corner resizes it: the angle stays 135°.
    [Puzzle("three-squares-ii")]
    public static void ThreeSquaresII(Lgf d)
    {
        d.Texture = "FeltTip";
        const double s = 4;
        string sn = Lgf.N(s);
        var a = d.Pt(0, 0, "A");
        var c = d.Pt(s, s, "C");
        var k = d.Slider(V.At(0.12 * s, s), V.At(0.32 * s, s), (0.238 - 0.12) / 0.2, (0.238 * s, s), "K");
        string length = d.Num($"sqrt(K.X ^ 2 + {sn} ^ 2)", "ak");
        string n = d.Num($"K.X / {length}", "n"), cs = d.Num($"{sn} / {length}", "c");
        string g = d.Num($"{sn} * ({cs} - {n})", "g");
        // the yellow square's top left corner slides along the line from A at 45° to its sides, its
        // side from 0.56 to 0.7 of s(sin + cos) (between half and all of it the line meets its top side), where the line through the big square's corner meets its top side
        string reach = d.Num($"{sn} * ({n} + {cs})", "reach");
        var diagonal = new V($"{n} - {cs}", $"{cs} + {n}");
        double home = 0.62 * 4 * 1.2044;
        var y1 = d.Slider($"0.56 * {reach}" * diagonal, $"0.7 * {reach}" * diagonal, (0.62 - 0.56) / 0.14, (home * -0.7412, home * 1.2044), "Y");
        string y = d.Num("sqrt(Y.X ^ 2 + Y.Y ^ 2) / sqrt(2)", "y");
        var y2 = d.Pt(y * new V(n, cs), "Y2");
        var y4 = d.Pt(y * new V($"-{cs}", n), "Y4");
        var g1 = d.Pt(c + g * new V($"-{cs}", n), "G1");
        var g4 = d.Pt(g1 + g * new V($"-{n}", $"-{cs}"), "G4");
        var g3 = d.Pt(c + g * new V($"-{n}", $"-{cs}"), "G3");
        var p = d.Pt(y2 + d.Num($"{reach} - {y}", "t") * new V($"-{cs}", n), "P");
        d.Polygon(d.Fill("Yellow", "#FFFFFF92", null), y1, y2, a, y4);
        d.Polygon(d.Fill("Green", "#FFAEEC9D", null), g1, c, g3, g4);
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        d.Box(frame, V.At(0, 0), V.At(s, s));
        d.Polygon(frame, y1, y2, a, y4);
        d.Polygon(frame, g1, c, g3, g4);
        string greenLine = "#FF10A872";
        string line = d.Line("GreenLine", greenLine, 3);
        d.Segment(line, g1, p);
        d.Angle(line, p, g1, y1, 26);
        d.Label("?", InAngle2021(d, p, g1, y1, 0.62), d.Text("GreenMark", greenLine, 28));
        d.Scene(-2.95, 5.2, s, 0);
    }

    // Two equilateral triangles sharing their top corner, the small one's other corner on the big
    // one's base; the big one has area 8. The blue triangle, from the small one's far corner to the
    // shared top and the big one's bottom right corner? Dragging the corner on the base resizes the
    // small triangle: the blue triangle changes and its area stays the big one's.
    [Puzzle("two-overlapping-triangles-ii")]
    public static void TwoOverlappingTrianglesII(Lgf d)
    {
        d.Texture = "FeltTip";
        const double s = 4;
        double h = 2 * S3 * s / 2;
        var a = d.Pt(0, h, "A");
        var left = d.Pt(-s / 2, 0, "P");
        var right = d.Pt(s / 2, 0, "R");
        var b = d.Slider(V.At(-1.6, 0), V.At(1, 0), (1.6 - 0.96) / 2.6, (-0.96, 0), "B");
        var l = d.Pt(a + (b - a).Turn(-60), "L");
        d.Polygon(d.Fill("Blue", "#FF8EC8FD", null), l, a, right);
        d.Segment(d.Line("LightBlue", "#FF439DD4", 3), l, right);
        d.Polygon(d.Outline("Big", Lgf.Ink, 3.5), left, right, a);
        d.Polygon(d.Outline("Small", "#FF2B2E9C", 3.5), l, a, b);
        d.Scene(-3.6, h, s / 2, 0);
    }

    /// <summary>A shape style that draws nothing: for a circle or polygon that only makes up a Region</summary>
    private static string Clear2021(Lgf d)
    {
        string texture = d.Texture;
        d.Texture = null;
        string name = d.Fill("Clear", "#00000000", null);
        d.Texture = texture;
        return name;
    }

    /// <summary>A point inside the angle at the vertex, on its bisector, the given distance from the vertex: where its label goes</summary>
    private static V InAngle2021(Lgf d, V vertex, V a, V b, double distance)
    {
        var toA = a - vertex;
        var toB = b - vertex;
        string la = d.Num($"sqrt({toA.X} ^ 2 + {toA.Y} ^ 2)"), lb = d.Num($"sqrt({toB.X} ^ 2 + {toB.Y} ^ 2)");
        var sum = d.Pt(new V($"{toA.X} / {la} + {toB.X} / {lb}", $"{toA.Y} / {la} + {toB.Y} / {lb}"));
        string ls = d.Num($"sqrt({sum.X} ^ 2 + {sum.Y} ^ 2)");
        return d.Pt(vertex + new V($"{Lgf.N(distance)} * {sum.X} / {ls}", $"{Lgf.N(distance)} * {sum.Y} / {ls}"));
    }
}
