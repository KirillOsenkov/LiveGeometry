#nullable disable

static partial class Puzzles
{
    // Three regular hexagons: a small one and a middle one inside the large one, the small one at its
    // left corner and the middle one at its right corner, meeting on the long diagonal (so their sides
    // add up to the large one's). The shaded trapezoid runs from the small one's left side to the
    // middle one's right side; what fraction of the large hexagon is it? Dragging the corner where
    // the two meet along the diagonal trades their sizes: the trapezoid changes, it stays half.
    [Puzzle("three-regular-hexagons-iii")]
    public static void ThreeRegularHexagonsIII(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 4;
        var big = HexagonFromLeft(d, V.At(-L, 0), Lgf.N(L));
        var m = d.Slider(V.At(-3.4, 0), V.At(0.8, 0), 1.4 / 4.2, (-2, 0), "M");
        string a = d.Num($"(M.X + {Lgf.N(L)}) / 2", "a");
        string b = d.Num($"{Lgf.N(L)} - {a}", "b");
        var small = HexagonFromLeft(d, big[0], a);
        var middle = HexagonFromLeft(d, m, b);
        V[] shaded = [small[1], middle[2], middle[4], small[5]];
        d.Polygon(d.Fill("Orange", "#FFFDC78B", null), shaded);
        string red = d.Outline("Red", "#FFE04E52");
        d.Polygon(red, small);
        d.Polygon(red, middle);
        d.Polygon(d.Outline("OrangeLine", "#FFF85530"), shaded);
        d.Polygon(d.Outline("Frame"), big);
        d.Scene(-L, 2 * S3 * L / 2, L, -2 * S3 * L / 2);
    }

    // Three squares: a green one and a blue one (area 12) side by side on a line, and a tilted one whose
    // diagonal runs from the green one's bottom left corner to the blue one's top right corner, its side
    // through the green one's top right corner. The green area? Nothing to drag: the side through that
    // corner makes the green square's side √2 times the blue one's.
    [Puzzle("two-squares-overlapping-a-square-ii")]
    public static void TwoSquaresOverlappingASquareII(Lgf d)
    {
        d.Texture = "FeltTip";
        double b = 2, g = 2 * Math.Sqrt(2);
        var o = d.Pt(0, 0, "O");
        var p = d.Pt(g + b, b, "P");
        var t = d.Pt(g / 2, g / 2 + b, "T");
        var q = d.Pt(g / 2 + b, -g / 2, "Q");
        d.Box(d.Fill("Green", "#FFA8E6AD"), V.At(0, 0), V.At(g, g));
        d.Box(d.Fill("Blue", "#FFA0D0FD"), V.At(g, 0), V.At(g + b, b));
        d.Polygon(d.Outline("Frame"), o, q, p, t);
        d.Label("?", V.At(0.47 * g, 0.52 * g), d.ShapeText("GreenMark", "#FF3AA04C", 40));
        d.Label("12", V.At(g + 0.45 * b, 0.52 * b), d.ShapeText("BlueMark", "#FF0A1AA8", 40));
        d.Scene(0, g / 2 + b, g + b, -g / 2);
    }

    // Three congruent triangles on a line: the red one and the purple one on it, the pink one between
    // them standing on its corner, with its long side over the red one's short side. Each has its 60°
    // angle at the middle point, so the line from the pink one's top to the purple one's far corner
    // makes an isosceles triangle with two long sides. The angle there? Dragging the purple one's top
    // along its side changes the shape of all three: the angle stays.
    [Puzzle("three-congruent-triangles-ii")]
    public static void ThreeCongruentTrianglesII(Lgf d)
    {
        d.Texture = "FeltTip";
        const double c = 6;
        var middle = d.Pt(0, 0, "C");
        var a = d.Pt(-c, 0, "A");
        var f = d.Pt(c, 0, "F");
        var top = d.Pt(-c / 2, c * S3, "D");
        var e = d.Slider(V.At(0.4 * c / 2, 0.4 * c * S3), V.At(0.9 * c / 2, 0.9 * c * S3), (0.78 - 0.4) / 0.5, (0.78 * c / 2, 0.78 * c * S3), "E");
        var b = d.Pt(new V("-E.X", "E.Y"), "B");
        d.Polygon(d.Fill("Red", "#FFE87A83"), a, middle, b);
        d.Polygon(d.Fill("Pink", "#FFFEC4C2"), top, middle, e);
        d.Polygon(d.Fill("Purple", "#FFB397CB"), middle, f, e);
        string indigo = "#FF4B2795";
        string line = d.Line("Indigo", indigo);
        d.Segment(line, top, f);
        d.Angle(line, f, top, middle, 50);
        d.Label("?", f + V.At(-1.55, 0.42), d.ShapeText("IndigoMark", indigo, 28));
        d.Scene(-c, c * S3, c, 0);
    }

    // A quarter circle and two equilateral triangles side by side on its bottom line, their tops on the
    // arc, 2 apart. The quarter circle's area? Seen from the triangles' shared corner the tops are 60°
    // apart, and the two tops are always 60° apart around the center too: the distance between them
    // is the radius. Dragging the big triangle's top along the arc trades the triangles' sizes; the
    // small one's top follows 60° round, and the quarter circle stays.
    [Puzzle("two-triangles-overlapping-a-quarter-circle")]
    public static void TwoTrianglesOverlappingAQuarterCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4;
        const string sqrt3 = "1.7320508075688772";
        var q = d.Pt(0, 0, "Q");
        var east = d.Pt(R, 0);
        var north = d.Pt(0, R);
        var p = d.ArcSlider(q, R, 72, 88, 81.7, (0, 0), "P");
        var p2 = d.Pt(V.Of("P").Turn(-60), "P2");
        string big = d.Num($"2 * P.Y / {sqrt3}", "big");
        var v = d.Pt(new V($"P.X + {big} / 2", "0"), "V");
        var left = d.Pt(new V($"V.X - {big}", "0"), "Lf");
        var right = d.Pt(new V(big, "0"), "Rt");
        string blue = d.Fill("Blue", "#FFA8DCFE");
        d.Polygon(blue, left, v, p);
        d.Polygon(blue, v, right, p2);
        d.Arc("CircleSector", d.Outline("Quarter", "#FF08108C"), q, east, north);
        d.Dimension(d.Line("Arrow"), d.Text("Mark", Lgf.Ink, 36), p, p2, 0, "2", 0.42);
        d.Scene(-2, R, 4.6, 0);
    }

    // Three congruent rectangles in a square: one upright on the left, one lying along the bottom
    // right, one upright on top of that at the right. Lines from the top right one's top left corner to
    // the left one's top left corner and to the bottom one's bottom left corner; the angle between
    // them? Dragging the left one's top right corner changes the rectangles' shape (the square stays):
    // the angle stays.
    [Puzzle("three-congruent-rectangles-iv")]
    public static void ThreeCongruentRectanglesIV(Lgf d)
    {
        d.Texture = "FeltTip";
        const double S = 6;
        var h = d.Slider(V.At(1, 5), V.At(2.8, 3.2), (2.03 - 1) / 1.8, (2.03, 3.97), "H");
        string w = "H.X", t = "H.Y";
        // the yellow stays bright in the dark theme: ink outlines there too, so the rectangles part
        string yellow = d.Fill("Yellow", "#FFFEFF78", darkOutline: Lgf.Ink);
        d.Box(yellow, V.At(0, 0), V.Of("H"));
        d.Box(yellow, new V(w, "0"), new V(Lgf.N(S), w));
        d.Box(yellow, new V(t, w), V.At(S, S));
        var top = d.Pt(new V(t, Lgf.N(S)), "T");
        var a = d.Pt(new V("0", t), "A");
        var b = d.Pt(new V(w, "0"), "B");
        string blue = "#FF2014CC";
        // across the paper and the yellow: a blue between light and deep shows on both in the dark theme
        string line = d.Line("Blue", blue, 3, dark: "#FF6B7BFF");
        d.Segment(line, top, a);
        d.Segment(line, top, b);
        d.Angle(line, top, a, b, 55);
        // on the bisector: the angle is always 45°, so the two unit vectors add up to 2 cos 22.5° long
        string ta = d.Num("distance(T, A)", "ta");
        string tb = d.Num("distance(T, B)", "tb");
        var bisector = new V($"(A.X - T.X) / {ta} + (B.X - T.X) / {tb}", $"(A.Y - T.Y) / {ta} + (B.Y - T.Y) / {tb}");
        d.Label("?", top + (1.25 / 1.8477590650225735) * bisector, d.Text("BlueMark", blue, 30, dark: "#FF8A97FF"));
        d.Scene(0, S, S, 0);
    }

    // A semicircle and three congruent right triangles: two on its diameter sharing their short leg,
    // the left one from the end of the diameter, and the third with its hypotenuse (10) along the top
    // from their shared corner to the arc, its right angle on the right one's hypotenuse. The
    // semicircle's area? The point on the arc is where the right one's legs would be measured from
    // the center: the radius is the hypotenuse. Dragging that point along the arc changes the
    // triangles' shape; the semicircle stays.
    [Puzzle("three-congruent-triangles-in-a-semi-circle")]
    public static void ThreeCongruentTrianglesInASemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 10;
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0);
        var west = d.Pt(-R, 0, "W");
        var q = d.ArcSlider(o, R, 18, 52, 38, (0, 0), "Q");
        var foot = d.Pt(new V("Q.X - 10", "0"), "F");
        var p = d.Pt(new V("Q.X - 10", "Q.Y"), "P");
        var b = d.Pt(new V("2 * Q.X - 10", "0"), "B");
        var x = d.Pt(new V("P.X + Q.X * Q.X / 10", "Q.Y - Q.X * Q.Y / 10"), "X");
        // one shape: three fills side by side would show seams between their gradients in the dark theme
        d.Region(d.Fill("Orange", "#FFFECC9B", null), "Union", d.Polygon(null, west, foot, p), d.Polygon(null, p, foot, b), d.Polygon(null, p, q, x));
        string frame = d.Outline("Frame");
        d.Polygon(frame, west, foot, p);
        d.Polygon(frame, p, foot, b);
        d.Polygon(frame, p, q, x);
        d.Arc("CircleSegment", frame, o, east, west);
        string orange = "#FFEE9A55";
        d.Dimension(d.Line("OrangeArrow", orange), d.Text("OrangeMark", orange, 32), p, q, 0.55, "10", 0.75);
        d.Scene(-R, R, R, 0);
    }

    // An equilateral triangle and a circle through its left and top corners and a point on its base;
    // the line from the top to that point is 9. The circle's area? The angle at the left corner is 60°,
    // so that chord is the radius times √3 wherever the point is. Dragging the left corner along the
    // circle (the base stays level) changes the triangle; the circle stays.
    [Puzzle("triangle-over-a-circle")]
    public static void TriangleOverACircle(Lgf d)
    {
        d.Texture = "FeltTip";
        double r = 3 * Math.Sqrt(3);
        const string sqrt3 = "1.7320508075688772";
        var o = d.Pt(0, 0, "O");
        var rim = d.Pt(r, 0);
        var a = d.ArcSlider(o, r, 216, 250, 228.3, (0, 0), "A");
        string s = d.Num($"-A.X - {sqrt3} * A.Y", "s");
        var x = d.Pt(new V("-A.X", "A.Y"), "X");
        var b = d.Pt(new V($"A.X + {s}", "A.Y"), "B");
        var c = d.Pt(new V($"A.X + {s} / 2", $"A.Y + {s3} * {s}"), "C");
        d.Polygon(d.Fill("Green", "#FFC0E597"), a, b, c);
        string teal = "#FF087665";
        d.Circle(d.Line("Circle", teal, 3), o, rim);
        string green = "#FF169A45";
        d.Segment(d.Line("GreenLine", green, 2), c, x);
        d.Label("9", Mid(d, c, x) + V.At(-0.5, 0.15), d.ShapeText("GreenMark", green, 30));
        d.Scene(-r, r, 6.72, -r);
    }

    // A big semicircle with three equal semicircles on its diameter, the outer two up, the middle one
    // down and cut out; the red semicircle stands on a chord of the big one and its arc runs through
    // the middle one's ends. The red one has area 10; the yellow area? Nothing to drag: the equal
    // semicircles fix it (the red one's radius is √5 small radii, the yellow is twice it).
    [Puzzle("multiple-semi-circles-iv")]
    public static void MultipleSemiCirclesIV(Lgf d)
    {
        d.Texture = "FeltTip";
        const double r = 1;
        double red = Math.Sqrt(5) * r;
        var o = d.Pt(0, 0, "O");
        var w3 = d.Pt(-3 * r, 0);
        var w1 = d.Pt(-r, 0);
        var e1 = d.Pt(r, 0);
        var e3 = d.Pt(3 * r, 0);
        var leftCenter = d.Pt(-2 * r, 0);
        var rightCenter = d.Pt(2 * r, 0);
        string big = d.Arc("CircleSegment", null, o, w3, e3, visible: false);
        string left = d.Arc("CircleSegment", null, leftCenter, w1, w3, visible: false);
        string right = d.Arc("CircleSegment", null, rightCenter, e3, e1, visible: false);
        d.Region(d.Fill("Yellow", "#FFFFFF8E", null), "Union", big, left, right);
        d.Arc("CircleSegment", d.PaperFill("Hole"), o, w1, e1);
        string ink = d.Line("Ink");
        d.Arc("CircleArc", ink, o, w3, e3);
        d.Arc("CircleArc", ink, leftCenter, w1, w3);
        d.Arc("CircleArc", ink, rightCenter, e3, e1);
        // the diameter runs across the yellow, which stays bright in the dark theme: ink there too
        d.Segment(d.Line("Diameter", Lgf.Ink, 2.5, dark: Lgf.Ink), w3, e3);
        var redCenter = d.Pt(0, -2 * r);
        var redEast = d.Pt(red, -2 * r);
        var redWest = d.Pt(-red, -2 * r);
        // mostly on the yellow, bright in the dark theme: a red between light and deep there
        string crimson = d.Line("Red", "#FFA3183F", 2.5, dark: "#FFD0285A");
        d.Arc("CircleArc", crimson, redCenter, redEast, redWest);
        d.Segment(crimson, redWest, redEast);
        d.Scene(-3 * r, r, 3 * r, -3 * r);
    }

    // Four squares: a pink one, a tilted one from its bottom left corner (its next corner on the pink
    // one's top side), a small one hanging from the pink one's bottom right corner down to the tilted
    // one's corner, and a light blue one beside it. Pink is the pink square outside the tilted one,
    // purple the tilted one outside the other three. Which is more? Dragging the tilted square's
    // bottom corner down the line of the pink one's right side tilts it (the small and the light blue
    // square follow): the pink and the purple stay equal.
    [Puzzle("four-squares-v")]
    public static void FourSquaresV(Lgf d)
    {
        d.Texture = "FeltTip";
        const double A = 6;
        var o = d.Pt(0, 0, "O");
        var corner = d.Pt(A, 0, "K");
        var topRight = d.Pt(A, A);
        var topLeft = d.Pt(0, A);
        var p4 = d.Slider(V.At(A, -0.9), V.At(A, -2.9), (2.05 - 0.9) / 2, (A, -2.05), "P");
        string b = "(-P.Y)";
        string c = d.Num($"1.25 * {b}", "c");
        var p2 = d.Pt(new V(b, Lgf.N(A)));
        var p3 = d.Pt(new V($"{Lgf.N(A)} + {b}", $"{Lgf.N(A)} - {b}"));
        V[] small = [d.Pt(new V($"{Lgf.N(A)} - {b}", "P.Y")), p4, corner, d.Pt(new V($"{Lgf.N(A)} - {b}", "0"))];
        V[] blue = [corner, d.Pt(new V(Lgf.N(A), $"-{c}")), d.Pt(new V($"{Lgf.N(A)} + {c}", $"-{c}")), d.Pt(new V($"{Lgf.N(A)} + {c}", "0"))];
        V[] pink = [o, corner, topRight, topLeft];
        V[] tilted = [o, p4, p3, p2];
        string pinkShape = d.Polygon(null, pink);
        string tiltedShape = d.Polygon(null, tilted);
        d.Region(d.Fill("Pink", "#FFFDBCD4", null), "Difference", pinkShape, tiltedShape);
        d.Region(d.Fill("Purple", "#FF8565C3", null), "Difference", tiltedShape, pinkShape, d.Polygon(null, small), d.Polygon(null, blue));
        d.Polygon(d.Outline("LightBlue", "#FF7FC8F2"), blue);
        d.Polygon(d.Outline("Navy", "#FF262DAE"), small);
        d.Polygon(d.Outline("Magenta", "#FFE8509F"), pink);
        d.Polygon(d.Outline("DarkPurple", "#FF3A1470", 3), tilted);
        d.Scene(0, A, A + 1.25 * 2.05, -1.25 * 2.05);
    }

    // Three squares: two blue ones of area 18 sharing a corner, the left one from the big square's
    // bottom left corner, the right one with its next corners on the big square's top and right sides.
    // The red lengths (from those corners to the big square's corners) add up to? The shared corner is
    // on the big square's diagonal, so the red lengths add up to its two coordinates. Dragging the right
    // blue square's top corner turns it about the shared corner; the big square grows or shrinks, the
    // red lengths change, their sum stays.
    [Puzzle("two-squares-overlapping-a-square")]
    public static void TwoSquaresOverlappingASquare(Lgf d)
    {
        d.Texture = "FeltTip";
        double s = 3 * Math.Sqrt(2), m = 3;
        var o = d.Pt(0, 0, "O");
        var shared = d.Pt(m, m, "M");
        var t = d.ArcSlider(shared, s, 52, 110, 69.5, (m, m), "T");
        var r = d.Pt(new V("T.Y", "6 - T.X"), "R");
        var u = d.Pt(new V("T.X + T.Y - 3", "T.Y + 3 - T.X"), "U");
        string blue = d.Fill("Blue", "#FF64A9FD", "#FF2E48CC");
        d.Polygon(blue, o, shared, d.Pt(0, 2 * m), d.Pt(-m, m));
        d.Polygon(blue, shared, r, u, t);
        var topLeft = d.Pt(new V("0", "T.Y"), "TL");
        var topRight = d.Pt(new V("T.Y", "T.Y"), "TR");
        var bottomRight = d.Pt(new V("T.Y", "0"), "BR");
        string ink = d.Line("Ink", Lgf.Ink, 3);
        d.Segment(ink, o, topLeft);
        d.Segment(ink, o, bottomRight);
        d.Segment(ink, t, topRight);
        d.Segment(ink, topRight, r);
        string red = d.Line("Red", "#FFEA5050", 3);
        d.Segment(red, topLeft, t);
        d.Segment(red, r, bottomRight);
        d.Scene(-m, 6.98, 8.46, 0);
    }

    // A square, its diagonal from the top left corner, and a zig-zag from the top right corner to a
    // point on the left side, a point on the right side and the bottom left corner, crossing the
    // diagonal at its quarters (the ticks). What fraction is shaded? Nothing to drag: the quarters put
    // the zig-zag's corners a third of the way up and down.
    [Puzzle("zig-zag-inside-a-square")]
    public static void ZigZagInsideASquare(Lgf d)
    {
        d.Texture = "FeltTip";
        const double S = 6;
        var bl = d.Pt(0, 0);
        var br = d.Pt(S, 0);
        var tr = d.Pt(S, S);
        var tl = d.Pt(0, S);
        var left = d.Pt(0, 2 * S / 3);
        var right = d.Pt(S, S / 3);
        var c1 = d.Pt(S / 4, 3 * S / 4);
        var c2 = d.Pt(S / 2, S / 2);
        var c3 = d.Pt(3 * S / 4, S / 4);
        string pink = d.Fill("Pink", "#FFFC4A70", null);
        d.Polygon(pink, tl, left, c1);
        d.Polygon(pink, c1, tr, right, c2);
        d.Polygon(pink, left, c2, c3, bl);
        d.Polygon(pink, c3, br, right);
        string ink = d.Line("Ink", Lgf.Ink, 3);
        d.Segment(ink, tr, left);
        d.Segment(ink, left, right);
        d.Segment(ink, right, bl);
        d.Segment(ink, tl, c1, "TwoTicks");
        d.Segment(ink, c1, c2, "TwoTicks");
        d.Segment(ink, c2, c3, "TwoTicks");
        d.Segment(ink, c3, br, "TwoTicks");
        d.Polygon(d.Outline("Frame", Lgf.Ink, 3), bl, br, tr, tl);
        d.Scene(0, S, S, 0);
    }
}
