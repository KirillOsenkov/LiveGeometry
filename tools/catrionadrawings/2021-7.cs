#nullable disable

static partial class Puzzles
{
    // A semicircle, a chord from its left end to a point on the arc, and two semicircles of area 4: one
    // on the chord from that point to the right end, bulging toward the first chord, and one on the
    // diameter, centered, touching the first chord and the tilted semicircle. The big semicircle's area?
    // Nothing to drag: two equal semicircles fix it (the chord rises 1 in 2, the radius is √5 times
    // theirs).
    [Puzzle("two-semi-circles-in-a-semi-circle")]
    public static void TwoSemiCirclesInASemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4;
        double r = R / Math.Sqrt(5);
        var o = d.Pt(0, 0, "O");
        var a = d.Pt(-R, 0, "A");
        var b = d.Pt(R, 0, "B");
        var c = d.Pt(R * 0.6, R * 0.8, "C");
        var m = d.Pt(R * 0.8, R * 0.4, "M");
        string green = d.Fill("Green", "#FF74D67C");
        d.Arc("CircleSegment", green, o, d.Pt(r, 0), d.Pt(-r, 0));
        d.Arc("CircleSegment", green, m, c, b);
        d.Segment(d.Line("Teal", "#FF1D7F78", 2), a, c);
        // inked after coloring: the diameter runs over the tilted semicircle where it dips below
        d.Arc("CircleSegment", d.Outline("Frame"), o, b, a);
        string mark = d.ShapeText("ShapeMark", Lgf.Ink, 40);
        d.Label("4", V.At(0, 0.64), mark);
        d.Label("4", V.At(2.57, 1.08), mark);
        d.Scene(-R, R, R, -0.2);
    }

    // A regular hexagon and six equilateral triangles in a zigzag band along its long diagonal, three
    // up and three down, their bases a third of the diagonal each. What fraction do they cover? Nothing
    // to drag: the diagonal fixes them.
    [Puzzle("six-triangles-in-a-hexagon")]
    public static void SixTrianglesInAHexagon(Lgf d)
    {
        d.Texture = "FeltTip";
        const double s = 3, h = 2 * S3;
        var hexagon = HexagonFromLeft(d, V.At(-s, 0), Lgf.N(s));
        d.Polygon(d.PaperFill("Paper", Lgf.Ink, 3), hexagon);
        var a = d.Pt(-1, 0);
        var b = d.Pt(1, 0);
        var up1 = d.Pt(-2, h);
        var up2 = d.Pt(0, h);
        var down1 = d.Pt(0, -h);
        var down2 = d.Pt(2, -h);
        d.Polygon(d.Fill("Yellow", "#FFFFFFAE", Lgf.Ink, 3), hexagon[0], a, up1);
        d.Polygon(d.Fill("LightGreen", "#FFC9FCC1", Lgf.Ink, 3), up1, up2, a);
        d.Polygon(d.Fill("Green", "#FF90D580", Lgf.Ink, 3), a, b, up2);
        d.Polygon(d.Fill("Blue", "#FF39A6FC", Lgf.Ink, 3), a, b, down1);
        d.Polygon(d.Fill("LightBlue", "#FFA1DDFE", Lgf.Ink, 3), down1, down2, b);
        d.Polygon(d.Fill("Purple", "#FFD4BEF9", Lgf.Ink, 3), b, hexagon[3], down2);
        d.Scene(-s, h * 1.5, s, -h * 1.5);
    }

    // Two equilateral triangles: the yellow one standing on its base, the orange one with a corner at
    // the middle of that base and its far side through the yellow one's top. The orange line from the
    // yellow one's right corner to the orange one's right corner; the angle there? The orange one's
    // right corner is on the circle through the base's middle, its right end and the top, so it is
    // dragged along that circle: the orange triangle turns and resizes, and the angle stays.
    [Puzzle("two-overlapping-triangles")]
    public static void TwoOverlappingTriangles(Lgf d)
    {
        const double a = 2;
        var m = d.Pt(0, 0, "M");
        var left = d.Pt(-a, 0, "A");
        var right = d.Pt(a, 0, "B");
        var top = d.Pt(0, 2 * a * S3, "P");
        var center = d.Pt(a / 2, a * S3, "C");
        var r = d.ArcSlider(center, a, 12, 76, 36, (a / 2, a * S3), "R");
        var q = d.Pt(V.Of("R").Turn(60), "Q");
        string yellowShape = d.Polygon(null, left, right, top);
        string orangeShape = d.Polygon(null, m, r, q);
        d.Region(d.Fill("Yellow", "#FFFEFF5C", null), "Difference", yellowShape, orangeShape);
        d.Region(d.Fill("Orange", "#FFFCA00A", null), "Difference", orangeShape, yellowShape);
        d.Region(d.Fill("LightOrange", "#FFFCD857", null), "Intersection", yellowShape, orangeShape);
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        // inked after coloring: over the regions (a plain polygon would be drawn under them)
        d.Inked(frame, left, right, top);
        d.Inked(frame, m, r, q);
        string ink = d.Line("Ink", Lgf.Ink, 3);
        d.Segment(ink, left, m, "TwoTicks");
        d.Segment(ink, m, right, "TwoTicks");
        string orange = "#FFE6550E";
        string line = d.Line("OrangeLine", orange, 2.5);
        d.Segment(line, right, r);
        d.Angle(line, r, m, right, 50);
        // the question mark on the angle's bisector, wherever the corner is
        string toM = d.Num("distance(R, M)"), toB = d.Num("distance(R, B)");
        var bisector = new V($"(M.X - R.X) / {toM} + (B.X - R.X) / {toB}", $"(M.Y - R.Y) / {toM} + (B.Y - R.Y) / {toB}");
        d.Label("?", r + 0.45 * bisector, d.Text("OrangeMark", orange, 26));
        d.Scene(-2.1, 3.75, 2.8, -0.15);
    }

    // Four congruent squares, all tilted the same way, in a rectangle: two side by side from the left
    // side down to the bottom, two stacked on the bottom up to the top and right sides, the lower one's
    // left side along the second square's right side. A line from the bottom left corner through the top
    // corner the first two share; the angle it makes with the top? Dragging the first square's corner up
    // and down the left side tilts all four (the rectangle follows): the line stays at the same angle.
    [Puzzle("four-squares-inside-a-rectangle")]
    public static void FourSquaresInsideARectangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double s = 2;
        var o = d.Pt(0, 0, "O");
        var p1 = d.Slider(V.At(0, 2 * s * Math.Sin(4 * Math.PI / 180)), V.At(0, 2 * s * Math.Sin(28 * Math.PI / 180)), 0.4, (0, 1.0), "P");
        string sn = d.Num($"P.Y / {Lgf.N(2 * s)}", "sn");
        string c = d.Num($"sqrt(1 - {sn} ^ 2)", "cs");
        var u = new V($"{Lgf.N(s)} * {c}", $"-{Lgf.N(s)} * {sn}");
        var v = new V($"{Lgf.N(s)} * {sn}", $"{Lgf.N(s)} * {c}");
        var a1 = d.Pt(p1 + u);
        var a2 = d.Pt(p1 + u + v, "T");
        var a3 = d.Pt(p1 + v);
        var q2 = d.Pt(p1 + 2 * u, "Q");
        var b2 = d.Pt(p1 + 2 * u + v);
        var b3 = d.Pt(q2 + $"{sn} / {c}" * v, "B");
        var c1 = d.Pt(b3 + u);
        var c2 = d.Pt(b3 + u + v);
        var c3 = d.Pt(b3 + v);
        var c4 = d.Pt(b3 + u + 2 * v);
        var c5 = d.Pt(b3 + 2 * v, "F");
        string w = d.Num($"{c4.X}", "w");
        var tl = d.Pt(new V("0", "F.Y"));
        var tr = d.Pt(new V(w, "F.Y"));
        var br = d.Pt(new V(w, "0"));
        var x = d.Pt(new V("F.Y", "F.Y"), "X");
        string pink = d.Fill("Pink", "#FFFEC2CC", "#FF4A2441");
        d.Polygon(pink, p1, a1, a2, a3);
        d.Polygon(pink, a1, q2, b2, a2);
        d.Polygon(pink, b3, c1, c2, c3);
        d.Polygon(pink, c3, c2, c4, c5);
        d.Polygon(d.Outline("Frame", "#FF4A2441"), o, br, tr, tl);
        string magenta = "#FFEA4182";
        string line = d.Line("PinkLine", magenta, 2.5);
        d.Segment(line, o, x);
        d.Angle(line, x, tl, o, 40);
        d.Label("?", x + V.At(-0.85, -0.38), d.Text("PinkMark", magenta, 28));
        d.Scene(0, 4.6, 7.3, 0);
    }

    // A semicircle and a circle crossing it twice. Lines from the ends of the diameter through those
    // crossings meet the circle again at the triangles' far corners, and cross each other at their
    // shared corner. Which shaded triangle is bigger? The circle's center and its top are dragged
    // (moving and resizing the circle): the far corners stay level with each other, so the two
    // triangles stay equal.
    [Puzzle("two-triangles-in-a-circle-and-semi-circle")]
    public static void TwoTrianglesInACircleAndSemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4;
        var o = d.Pt(0, 0, "O");
        var a = d.Pt(-R, 0, "A");
        var b = d.Pt(R, 0, "B");
        var c = d.Slider(V.At(1.4, 3.09), V.At(2.4, 3.09), 0.6, (2, 3.09), "C");
        var t = d.Slider(new V("C.X", "C.Y + 2.3"), new V("C.X", "C.Y + 3"), 0.5, (2, 5.74), "T");
        // where the circle crosses the semicircle: E on the left, F on the right
        string r = d.Num("T.Y - C.Y", "r");
        string dc = d.Num("sqrt(C.X ^ 2 + C.Y ^ 2)", "dc");
        string along = d.Num($"({Lgf.N(R * R)} - {r} ^ 2 + {dc} ^ 2) / (2 * {dc})", "al");
        string across = d.Num($"sqrt({Lgf.N(R * R)} - {along} ^ 2)", "ac");
        var foot = new V($"{along} * C.X / {dc}", $"{along} * C.Y / {dc}");
        var side = new V($"-{across} * C.Y / {dc}", $"{across} * C.X / {dc}");
        var e = d.Pt(foot + side, "E");
        var f = d.Pt(foot - side, "F");
        var p = Drawing2021_7.SecondCrossing(d, c, a, e, "P");
        var q = Drawing2021_7.SecondCrossing(d, c, b, f, "Q");
        var x = d.Crossing(a, q, b, p, "X");
        string blueLine = "#FF2A4FC8", grayLine = "#FF555566";
        d.Polygon(d.Fill("Blue", "#FF5FD6FE", blueLine), a, p, x);
        d.Polygon(d.Fill("Green", "#FF94FA40", grayLine), x, q, b);
        string green = d.Line("GreenLine", "#FF2EA15C", 2.5);
        d.Arc("CircleArc", green, o, b, a);
        d.Segment(green, a, b);
        d.Circle(green, c, t);
        d.DotAt(o);
        d.Scene(-R, 5.9, R + 0.7, 0);
    }

    // A rectangle, a point on its top joined to the bottom corners, and a line from the bottom left
    // corner to the right line; the three marked angles (two at the top point, one at the bottom left
    // corner) are equal. What fraction is shaded? Equal angles make the big triangle isosceles (its
    // left side as long as the bottom) and the line its median. Dragging the top right corner changes
    // the rectangle's width (the top point follows): the shaded part stays a quarter.
    [Puzzle("triangle-in-a-rectangle")]
    public static void TriangleInARectangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double H = 4;
        var a = d.Pt(0, 0, "A");
        var topRight = d.Slider(V.At(H * 1.003, H), V.At(H * 1.6, H), (4.26 / H - 1.003) / 0.597, (4.26, H), "TR");
        var b = d.Pt(new V("TR.X", "0"), "B");
        var t = d.Pt(new V($"sqrt(B.X ^ 2 - {Lgf.N(H * H)})", Lgf.N(H)), "T");
        var x = d.Pt(new V("(T.X + B.X) / 2", Lgf.N(H / 2)), "X");
        var topLeft = d.Pt(0, H, "TL");
        d.Polygon(d.Fill("Purple", "#FFAD8ECB", null), a, b, x);
        string ink = d.Line("Ink", Lgf.Ink, 3);
        d.Polygon(d.Outline("Frame", Lgf.Ink, 3), a, b, topRight, topLeft);
        d.Segment(ink, a, t);
        d.Segment(ink, t, b);
        d.Segment(ink, a, x);
        string pink = d.Line("Pink", "#FFEC8AAA", 2.5);
        d.Angle(pink, t, a, b, 46);
        d.Angle(pink, t, b, topRight, 46);
        d.Angle(pink, a, x, topLeft, 46);
        d.Scene(0, H, 4.26, 0);
    }

    // Two semicircles on the bottom line, a big one and a small one inside it, a line across both,
    // and a semicircle on each one's chord along that line. Yellow: the big chord's semicircle outside
    // the big semicircle; red: the small chord's semicircle outside the small semicircle, and the big
    // semicircle's part under the line outside the small one. More red or yellow? Dragging the line's
    // left end along the big arc moves the line up and down, dragging the small semicircle's right
    // foot resizes it: the red and the yellow stay equal.
    [Puzzle("multiple-semi-circles-ii")]
    public static void MultipleSemiCirclesII(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4, leftFoot = -3.18;
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0);
        var west = d.Pt(-R, 0);
        var l = d.ArcSlider(o, R, 152.45, 168.5, 154.85, (0, 0), "L");
        var f = d.Slider(V.At(0.8, 0), V.At(2.2, 0), (1.62 - 0.8) / 1.4, (1.62, 0), "F");
        var lr = d.Pt(new V("-L.X", "L.Y"), "Lr");
        var bigMiddle = d.Pt(new V("0", "L.Y"), "Bm");
        string c = d.Num($"(F.X + {Lgf.N(leftFoot)}) / 2", "c");
        string r = d.Num($"(F.X - {Lgf.N(leftFoot)}) / 2", "r");
        string b = d.Num($"sqrt({r} ^ 2 - L.Y ^ 2)", "b");
        var smallCenter = d.Pt(new V(c, "0"), "Sc");
        var smallFoot = d.Pt(V.At(leftFoot, 0), "Sf");
        var chordCenter = d.Pt(new V(c, "L.Y"), "Cc");
        var chordLeft = d.Pt(new V($"{c} - {b}", "L.Y"), "Cl");
        var chordRight = d.Pt(new V($"{c} + {b}", "L.Y"), "Cr");
        string big = d.Arc("CircleSegment", null, o, east, west, visible: false);
        string bigChord = d.Arc("CircleSegment", null, bigMiddle, lr, l, visible: false);
        string small = d.Arc("CircleSegment", null, smallCenter, f, smallFoot, visible: false);
        string smallChord = d.Arc("CircleSegment", null, chordCenter, chordRight, chordLeft, visible: false);
        string above = d.Polygon(null, d.Pt(new V(Lgf.N(-R), "L.Y")), d.Pt(new V(Lgf.N(R), "L.Y")), d.Pt(R, R), d.Pt(-R, R));
        d.Region(d.Fill("Yellow", "#FFFFFE30", null), "Difference", bigChord, big);
        // two red regions (they only touch at points): a Union of the two hung the player
        string red = d.Fill("Red", "#FFFB2E24", null);
        d.Region(red, "Difference", smallChord, small);
        d.Region(red, "Difference", big, small, above);
        string ink = d.Line("Ink", Lgf.Ink, 3);
        d.Arc("CircleArc", ink, o, east, west);
        d.Arc("CircleArc", ink, bigMiddle, lr, l);
        d.Arc("CircleArc", ink, smallCenter, f, smallFoot);
        d.Arc("CircleArc", ink, chordCenter, chordRight, chordLeft);
        d.Segment(ink, west, east);
        d.Segment(ink, l, lr);
        d.Scene(-R, 1.7 + Math.Sqrt(R * R - 1.7 * 1.7), R, 0);
    }

    // A blue square and a light blue rectangle as high as it, side by side, and a tilted square of area
    // 25 with a corner at their bottom left corner, its top side through the blue square's top right
    // corner and its next corner on the rectangle's right side. The square and the rectangle together?
    // Dragging the tilted square's top left corner turns it: the blue square and the rectangle change,
    // and their area stays 25.
    [Puzzle("square-overlapping-a-square-and-a-rectangle")]
    public static void SquareOverlappingASquareAndARectangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 5;
        var a = d.Pt(0, 0, "A");
        var p1 = d.ArcSlider(a, L, 50, 85, 74, (0, 0), "P");
        // P = L (sin t, cos t), the square turned by t: its sides along (cos t, -sin t) and (sin t, cos t);
        // its top side is L from A, so the blue square's side s has s (sin t + cos t) = L
        var p3 = d.Pt(new V("P.Y", "-P.X"), "P3");
        var p2 = d.Pt(new V("P.X + P.Y", "P.Y - P.X"), "P2");
        string s = d.Num($"{Lgf.N(L * L)} / (P.X + P.Y)", "s");
        string navy = "#FF0A0A80";
        d.Box(d.Fill("Blue", "#FF0470F2", navy, 3), V.At(0, 0), new V(s, s));
        d.Box(d.Fill("LightBlue", "#FF89CAFD", navy, 3), new V(s, "0"), new V("P2.X", s));
        d.Inked(d.Outline("Square", Lgf.Ink, 3.5), a, p1, p2, p3);
        d.Scene(0, 4.8, 6.2, -1.4);
    }

    // A square on the diameter of a semicircle, and a chord of length 6 from the square's left side to
    // its right side through its center (halving it), its ends on the arc. The semicircle's area? The
    // chord's ends are a quarter turn apart around the center, so the radius is 6 / √2. Dragging the
    // chord's right end along the arc tilts the chord and resizes the square: the semicircle stays.
    [Puzzle("chord-splitting-a-square-in-half")]
    public static void ChordSplittingASquareInHalf(Lgf d)
    {
        d.Texture = "FeltTip";
        double r = Math.Sqrt(18);
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(r, 0);
        var west = d.Pt(-r, 0);
        // the right end E = (p + m, p - m), the left end F = (m - p, p + m), the square's side 2p
        var e = d.ArcSlider(o, r, 8, 44, 27.9, (0, 0), "E");
        var f = d.Pt(new V("-E.Y", "E.X"), "F");
        var bottomLeft = d.Pt(new V("-E.Y", "0"));
        var bottomRight = d.Pt(new V("E.X", "0"));
        var topRight = d.Pt(new V("E.X", "E.X + E.Y"));
        var topLeft = d.Pt(new V("-E.Y", "E.X + E.Y"));
        string outline = "#FF3A3020";
        d.Polygon(d.Fill("Green", "#FFB8E178", null), f, e, topRight, topLeft);
        d.Polygon(d.Fill("LightGreen", "#FFF0FFB0", null), bottomLeft, bottomRight, e, f);
        d.Polygon(d.Outline("Square", outline), bottomLeft, bottomRight, topRight, topLeft);
        string line = d.Line("Line", outline);
        d.Arc("CircleArc", line, o, east, west);
        d.Segment(line, west, east);
        string blue = "#FF3466B0";
        d.Segment(d.Line("Blue", blue, 3), f, e);
        d.Label("6", Mid(d, e, f) + new V("0.075 * (E.Y - F.Y)", "-0.075 * (E.X - F.X)"), d.ShapeText("BlueMark", blue, 36));
        d.Scene(-r, 5.75, r, 0);
    }

    // A black square and an orange rectangle 4 wide with a corner at the middle of the square's bottom,
    // the next corner on its left side and the far one on its top. The shaded area? The rectangle is
    // twice as long as wide. Dragging the corner on the left side along the circle of radius 4 around
    // the middle tilts the rectangle and resizes the square: the rectangle stays 4 by 8.
    [Puzzle("rectangle-tilted-over-a-square")]
    public static void RectangleTiltedOverASquare(Lgf d)
    {
        var m = d.DotAt(V.At(0, 0), "M");
        var p = d.ArcSlider(m, 4, 154, 174, 160, (0, 0), "P");
        // the square's half side s = -P.X; the long side runs from M along (P.Y, s), twice as long
        var q = d.Pt(new V("2 * P.Y", "-2 * P.X"), "Q");
        var r = d.Pt(new V("P.X + 2 * P.Y", "P.Y - 2 * P.X"), "R");
        var bottomLeft = d.Pt(new V("P.X", "0"));
        var bottomRight = d.Pt(new V("-P.X", "0"));
        var topRight = d.Pt(new V("-P.X", "-2 * P.X"));
        var topLeft = d.Pt(new V("P.X", "-2 * P.X"));
        string orange = "#FFF8420A";
        d.Polygon(d.Fill("Orange", "#FFFDC666", orange, 3), m, q, r, p);
        // inked after coloring: the square's sides run over the rectangle
        d.Inked(d.Outline("Square", Lgf.Ink, 3), bottomLeft, bottomRight, topRight, topLeft);
        string ink = d.Line("Ink", Lgf.Ink, 3);
        d.Segment(ink, bottomLeft, m, "TwoTicks");
        d.Segment(ink, m, bottomRight, "TwoTicks");
        d.Dimension(d.Line("Arrow", orange, 2.5), d.Text("OrangeMark", orange, 30), r, q, 0.5, "4", 0.4);
        d.Scene(-3.8, 9.6, 3.8, -0.2);
    }

    // A quadrilateral on a base halved by a line from the top left corner; its diagonal from the bottom
    // left corner crosses that line. Angles 50, 70, 80 and 60 given; the angle at the top right corner
    // between the top side and the diagonal? Nothing to drag: the angles fix it (both top corners are
    // on the circle on the base, 100° and 60° round from its right end).
    [Puzzle("multiple-angles")]
    public static void MultipleAngles(Lgf d)
    {
        const double R = 4;
        var m = d.Pt(0, 0, "M");
        var a = d.Pt(-R, 0, "A");
        var b = d.Pt(R, 0, "B");
        var t = d.Pt(R * Math.Cos(100 * Math.PI / 180), R * Math.Sin(100 * Math.PI / 180), "T");
        var s = d.Pt(R / 2, R * S3, "S");
        var x = d.Crossing(t, m, a, s, "X");
        string ink = d.Line("Ink");
        d.Segment(ink, a, m, "TwoTicks");
        d.Segment(ink, m, b, "TwoTicks");
        d.Segment(ink, a, t);
        d.Segment(ink, t, s);
        d.Segment(ink, s, b);
        d.Segment(ink, t, m);
        d.Segment(ink, a, s);
        foreach (var point in new[] { a, m, b })
        {
            d.DotAt(point);
        }

        var marks = new[]
        {
            ("Blue", "#FF0341CC", t, a, m, "50°", V.At(-1.15, 3.0)),
            ("Red", "#FFFB1231", x, t, s, "70°", V.At(-0.03, 2.77)),
            ("Green", "#FF0BA80B", s, t, a, "?", V.At(1.03, 3.28)),
            ("Orange", "#FFFD981F", m, t, a, "80°", V.At(-0.8, 0.45)),
            ("Purple", "#FF230278", b, a, s, "60°", V.At(3.0, 0.5)),
        };
        foreach (var (name, color, vertex, from, to, text, at) in marks)
        {
            d.Angle(d.Wedge(name + "Wedge", color), vertex, from, to, 28);
            d.Label(text, at, d.Text(name + "Mark", color, 27));
        }

        d.Scene(-R, R, R, 0);
    }

    // A rectangle in a circle, semicircles on its top and bottom sides bulging in, and three equal
    // circles: one between the semicircles, touching both, and one each side between the rectangle and
    // the circle. What fraction is shaded? Nothing to drag: equal circles make the rectangle 3 by 4.
    [Puzzle("three-circles-and-two-semi-circles-in-a-circle")]
    public static void ThreeCirclesAndTwoSemiCirclesInACircle(Lgf d)
    {
        const double w = 3, h = 4, R = 5, r = 1;
        var o = d.Pt(0, 0, "O");
        var topLeft = d.Pt(-w, h);
        var topRight = d.Pt(w, h);
        var bottomRight = d.Pt(w, -h);
        var bottomLeft = d.Pt(-w, -h);
        var top = d.Pt(0, h);
        var bottom = d.Pt(0, -h);
        string pink = d.Fill("Pink", "#FFFBD5F5", null);
        d.Arc("CircleSegment", pink, top, topLeft, topRight);
        d.Arc("CircleSegment", pink, bottom, bottomRight, bottomLeft);
        string ink = d.Line("Ink", Lgf.Ink, 3);
        d.Arc("CircleArc", ink, top, topLeft, topRight);
        d.Arc("CircleArc", ink, bottom, bottomRight, bottomLeft);
        string circle = d.Fill("Circle", "#FFFBD5F5", Lgf.Ink, 3);
        foreach (double x in new[] { -w - r, 0, w + r })
        {
            d.Circle(circle, d.Pt(x, 0), d.Pt(x + r, 0));
        }

        d.Circle(d.Outline("Frame", Lgf.Ink, 3), o, d.Pt(R, 0));
        // inked after coloring: the rectangle's top and bottom run along the semicircles' diameters
        d.Inked(d.Outline("Rectangle", "#FFD82AA0", 3), bottomLeft, bottomRight, topRight, topLeft);
        d.Scene(-R, R, R, -R);
    }

}

file static class Drawing2021_7
{
    /// <summary>Where the line from a point through a point on a circle (around the center) meets the circle again</summary>
    public static V SecondCrossing(Lgf d, V center, V from, V on, string name)
    {
        string length = d.Num($"sqrt(({on.X} - {from.X}) ^ 2 + ({on.Y} - {from.Y}) ^ 2)");
        var w = new V($"({on.X} - {from.X}) / {length}", $"({on.Y} - {from.Y}) / {length}");
        string s = d.Num($"-2 * (({on.X} - {center.X}) * {w.X} + ({on.Y} - {center.Y}) * {w.Y})");
        return d.Pt(on + s * w, name);
    }
}
