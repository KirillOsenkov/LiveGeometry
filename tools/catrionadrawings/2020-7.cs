#nullable disable

static partial class Puzzles
{
    // A purple square and a red rectangle sharing its diagonal: the rectangle's corners are the
    // square's top left and bottom right ones, its long sides leave the square a quarter of the way
    // down the right side and up the left. 75% of the square is shaded; what percentage of the
    // rectangle? Nothing to drag: the 75% puts the crossings at the quarters and fixes the rectangle.
    [Puzzle("a-square-and-a-rectangle-overlapping")]
    public static void ASquareAndARectangleOverlapping(Lgf d)
    {
        const double s = 4;
        var a = d.Pt(0, s);
        var c = d.Pt(s, 0);
        var p = d.Pt(s, 0.75 * s);
        var q = d.Pt(0, 0.25 * s);
        d.Polygon(d.Fill("Pink", "#FFFDB8FD", null), a, q, c, p);
        d.Box(d.Outline("Purple", "#FF6D087F", 3), V.At(0, 0), V.At(s, s));
        // a vivid red in the dark theme: its light tint would be as pink as the purple's
        d.Polygon(d.Outline("Red", "#FFF80A10", 3, dark: "#FFFF4A4A"), a, d.Pt(20 * s / 17, 12 * s / 17), c, d.Pt(-3 * s / 17, 5 * s / 17));
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(-3 * s / 17, s, 20 * s / 17, -0.4);
    }

    // Three equal circles touching each other and two sides of a regular hexagon each (it has a
    // corner at the top), the hexagon in a circle. What fraction is shaded? Nothing to drag: the
    // circles' radius is √3/4 of the big one's.
    [Puzzle("three-circles-in-a-hexagon-in-a-circle")]
    public static void ThreeCirclesInAHexagonInACircle(Lgf d)
    {
        const double R = 4;
        double r = R * Math.Sqrt(3) / 4;
        var o = d.Pt(0, 0, "O");
        var hexagon = Enumerable.Range(0, 6).Select(i => d.Pt(R * Math.Cos((90 + 60 * i) * Math.PI / 180), R * Math.Sin((90 + 60 * i) * Math.PI / 180))).ToArray();
        d.Polygon(d.Outline("Hexagon", "#FF12C52E", 3), hexagon);
        d.Circle(d.Outline("Frame", Lgf.Ink, 3.5), o, hexagon[0]);
        string green = d.Fill("Green", "#FFDCFEA8", width: 3.5);
        foreach (double angle in new[] { 90.0, 210, 330 })
        {
            double x = R / 2 * Math.Cos(angle * Math.PI / 180), y = R / 2 * Math.Sin(angle * Math.PI / 180);
            d.Circle(green, d.Pt(x, y), d.Pt(x + r, y));
        }

        d.Scene(-R, R, R, -R);
    }

    // Two small circles in a row and two lines from the first one's center touching the second; the
    // big circle touches both lines and the second circle. The small ones have area 1; the big one?
    // Nothing to drag: the lines are 30° from the axis, and the big circle has three times the radius.
    [Puzzle("three-circles")]
    public static void ThreeCircles(Lgf d)
    {
        var o = d.Pt(0, 0, "O");
        string cyan = d.Line("Cyan", "#FF21CFE8", 3);
        double c = Math.Cos(Math.PI / 6), s = Math.Sin(Math.PI / 6);
        d.Segment(cyan, o, d.Pt(7.9 * c, 7.9 * s));
        d.Segment(cyan, o, d.Pt(7.9 * c, -7.9 * s));
        string frame = d.Outline("Frame", Lgf.Ink, 3.5);
        d.Circle(frame, o, d.Pt(1, 0));
        d.Circle(frame, d.Pt(2, 0), d.Pt(3, 0));
        var big = d.Pt(6, 0);
        d.Circle(d.Fill("Blue", "#FFC8EFFB", "#FF1E6FF8", 3.5), big, d.Pt(9, 0));
        d.DotAt(o, "Od", 11);
        d.Label("?", big, d.ShapeText("BlueMark", "#FF1E6FF8", 48));
        d.Scene(-1, 7.9 * s, 9, -7.9 * s);
    }

    // Three squares of areas 4, 36 and 9 side by side on the base of a triangle, whose sides run
    // through their top corners. The triangle's area? Nothing to drag: the squares fix it.
    [Puzzle("three-squares-in-a-triangle")]
    public static void ThreeSquaresInATriangle(Lgf d)
    {
        string frame = d.Outline("Frame", Lgf.Ink, 3.5);
        d.Box(frame, V.At(0, 0), V.At(2, 2));
        d.Box(frame, V.At(2, 0), V.At(8, 6));
        d.Box(frame, V.At(8, 0), V.At(11, 3));
        // the left side through (0, 2) and (2, 6), the right one through (8, 6) and (11, 3)
        d.Polygon(d.Outline("Red", "#FFD0080A", 3.5, dark: "#FFFF4A4A"), d.Pt(-1, 0), d.Pt(14, 0), d.Pt(4, 10));
        string mark = d.Text("Mark", Lgf.Ink, 30);
        d.Label("4", V.At(1, 1), mark);
        d.Label("36", V.At(5, 3), d.Text("Big", Lgf.Ink, 40));
        d.Label("9", V.At(9.5, 1.5), mark);
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(-1, 10, 14, -1.2);
    }

    // Four squares: a big one (side 10) between a left one and a right one, all on one line 22 long,
    // and a small one on the left one, filling the gap to the big one's top. The shaded quadrilateral
    // joins the small one's bottom left corner, the big one's top right and bottom left corners and
    // the right one's top right corner. Its area? Dragging the big square's bottom left corner along
    // the line trades the left square's size for the right one's: the shaded area stays 100.
    [Puzzle("four-squares")]
    public static void FourSquares(Lgf d)
    {
        d.Texture = "FeltTip";
        var foot = d.Slider(V.At(5.4, 0), V.At(9.3, 0), (6.9 - 5.4) / 3.9, (6.9, 0), "A");
        string a = "A.X";
        var p1 = d.Pt(new V($"2 * {a} - 10", a));
        var p2 = d.Pt(new V($"{a} + 10", "10"));
        var p3 = d.Pt(new V("22", $"12 - {a}"));
        string teal = "#FF0AB9AD";
        d.Polygon(d.Fill("Teal", "#FF05E0B6", teal), foot, p3, p2, p1);
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        d.Box(frame, V.At(0, 0), new V(a, a));
        d.Box(frame, new V(a, "0"), new V($"{a} + 10", "10"));
        d.Box(frame, new V($"{a} + 10", "0"), new V("22", $"12 - {a}"));
        d.Box(frame, new V($"2 * {a} - 10", a), new V(a, "10"));
        string arrow = d.Line("Arrow", teal);
        string mark = d.Text("TealMark", teal, 28);
        d.Dimension(arrow, mark, d.Pt(0, 0), d.Pt(0, 10), 1, "10", 1.5);
        d.Dimension(arrow, mark, d.Pt(0, 0), d.Pt(22, 0), -1, "22", 2);
        d.Scene(-1.6, 10, 22, -1.6);
    }

    // A semicircle, a chord from the left end at 30° and a chord whose arc (red) is as long as the
    // arc the first chord leaves on the right (red). What fraction of the semicircle is shaded (between
    // the chords)? Dragging the second chord's end along the arc slides it round: the shaded part
    // changes shape, its area stays the 120° segment minus the 60° one.
    [Puzzle("arcs-in-a-semi-circle")]
    public static void ArcsInASemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4;
        var o = d.Pt(0, 0, "O");
        var e = d.Pt(R, 0, "E");
        var w = d.Pt(-R, 0, "W");
        var c = d.Pt(R / 2, R * S3, "C");
        var a = d.ArcSlider(o, R, 126, 174, 142, (0, 0), "A");
        var b = d.Pt(V.Of("A").Turn(-60), "B");
        string big = d.Arc("CircleSegment", null, o, c, w, visible: false);
        string cut = d.Arc("CircleSegment", null, o, b, a, visible: false);
        d.Region(d.Fill("Yellow", "#FFFFEE02", null), "Difference", big, cut);
        string ink = d.Line("Ink", Lgf.Ink, 3);
        string red = d.Line("Red", "#FFC0100E", 3);
        d.Segment(ink, w, e);
        d.Segment(ink, w, c);
        d.Segment(ink, a, b);
        d.Arc("CircleArc", ink, o, a, w);
        d.Arc("CircleArc", red, o, b, a);
        d.Arc("CircleArc", ink, o, c, b);
        d.Arc("CircleArc", red, o, e, c);
        d.ArcMark(d.Line("MarkLine", Lgf.Ink, 2.5), w, e, c, 0.8);
        d.Label("30°", V.At(-R + 2.0, 0.42), d.Text("Mark", Lgf.Ink, 22));
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(-R, R, R, -0.5);
    }

    // A triangle cut from a point inside into three isosceles triangles (blue, purple, red) and a white
    // quadrilateral: all four segments from the point are equal, so the bottom corners and the two
    // points on the sides are on a circle around it. The angle at the top is 50°, the white angle at
    // the point 35°; the angle at the bottom? Dragging the point on the left side round that circle
    // (the one on the right follows 35° behind) changes the blue and the purple triangles: the top
    // angle stays 50° and the asked angle stays (half the far arc minus the near arc is the top angle).
    [Puzzle("three-isoscelese-triangles-inside-a-triangle")]
    public static void ThreeIsosceleseTrianglesInsideATriangle(Lgf d)
    {
        const double r = 5;
        var p = d.Pt(0, 0, "P");
        var a = d.Pt(r * Math.Cos(202.5 * Math.PI / 180), r * Math.Sin(202.5 * Math.PI / 180), "A");
        var b = d.Pt(r * Math.Cos(337.5 * Math.PI / 180), r * Math.Sin(337.5 * Math.PI / 180), "B");
        var q = d.ArcSlider(p, r, 89, 122, 96.5, (0, 0), "Q");
        var right = d.Pt(V.Of("Q").Turn(-35), "R");
        var t = d.Crossing(a, q, b, right, "T");
        d.Polygon(d.Fill("Blue", "#FF7EBBEE", null), a, q, p);
        d.Polygon(d.Fill("Purple", "#FFBC76D7", null), p, right, b);
        d.Polygon(d.Fill("Red", "#FFF36C7E", null), a, b, p);
        string ink = d.Line("Ink", Lgf.Ink, 3);
        d.Polygon(d.Outline("Frame", Lgf.Ink, 3), a, b, t);
        d.Segment(ink, p, q);
        d.Segment(ink, p, right);
        d.Segment(ink, p, a);
        d.Segment(ink, p, b);
        string red = "#FFE53A1E";
        string redLine = d.Line("RedLine", red, 2.5);
        string redText = d.Text("RedMark", red, 24);
        d.Angle(redLine, t, a, b, 30);
        d.Angle(redLine, p, right, q, 30);
        d.Angle(d.Line("MarkLine", Lgf.Ink, 2.5), p, a, b, 22);
        string ta = d.Num("distance(T, A)", "ta");
        string tb = d.Num("distance(T, B)", "tb");
        // on the bisector of the 50° angle, 1.5 from the top: the two unit vectors add up to 2 cos 25° long
        var bisector = new V($"(A.X - T.X) / {ta} + (B.X - T.X) / {tb}", $"(A.Y - T.Y) / {ta} + (B.Y - T.Y) / {tb}");
        d.Label("50°", t + (1.55 / 1.8126155740733) * bisector, redText);
        // Q and R are on the circle: Q + R is along the bisector of the 35° angle, 2 r cos 17.5° long
        d.Label("35°", Lgf.N(2.1 / (2 * r * Math.Cos(17.5 * Math.PI / 180))) * (q + right), redText);
        d.Label("?", V.At(0, -1.25), d.ShapeText());
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(-4.62, 9.6, 4.62, -2.6);
    }

    // A semicircle, a rectangle in it with its top corners on the arc, and a square on the left end of
    // the diameter whose next corner is the rectangle's top left one and whose far side runs through
    // the center. The pink overlap is 15; the total yellow area? Nothing to drag: the far side through
    // the center puts the shared corner at (-3/5, 4/5) of the radius.
    [Puzzle("square-and-rectangle-overlapping-in-a-semi-circle")]
    public static void SquareAndRectangleOverlappingInASemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 5;
        var o = d.Pt(0, 0, "O");
        var e = d.Pt(R, 0, "E");
        var w = d.Pt(-R, 0, "W");
        V[] square = [w, d.Pt(-3, 4, "K"), d.Pt(1, 2), d.Pt(-1, -2)];
        V[] rectangle = [d.Pt(-3, 0), d.Pt(3, 0), d.Pt(3, 4), square[1]];
        string squareShape = d.Polygon(null, square);
        string rectangleShape = d.Polygon(null, rectangle);
        d.Region(d.Fill("Yellow", "#FFF5E804", null), "Union", squareShape, rectangleShape);
        d.Region(d.Fill("Magenta", "#FFEE02AA", null), "Intersection", squareShape, rectangleShape);
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        d.Polygon(frame, rectangle);
        d.Polygon(frame, square);
        d.Arc("CircleSegment", frame, o, e, w);
        d.DotAt(o, "Od", 13);
        d.Label("15", V.At(-1.55, 1.75), d.ShapeText("ShapeMark", Lgf.Ink, 36));
        d.Scene(-R, R, R, -2);
    }

    // Three small squares of area 4, one stacked on the left one of a pair (a little to the left), and
    // a tilted cyan square with a corner at the pair's middle bottom corner and a side through the top
    // square's top left corner. The total shaded area (the cyan square over the small ones)? Dragging
    // where its right side crosses the pair's top turns it (it grows to keep its side through that
    // corner): the top square stays inside it, and its right angle at the middle of the pair always
    // takes a quarter of the 4 by 4 square around that corner: the shaded area stays 8.
    [Puzzle("three-stacked-squares-and-a-fourth-square")]
    public static void ThreeStackedSquaresAndAFourthSquare(Lgf d)
    {
        d.Texture = "FeltTip";
        var m = d.Pt(0, 0, "M");
        var h = d.Slider(V.At(0.72, 2), V.At(1.2, 2), (0.8 - 0.72) / 0.48, (0.8, 2), "H");
        string length = d.Num("sqrt(H.X ^ 2 + 4)", "l");
        var along = new V($"H.X / {length}", $"2 / {length}");
        var across = new V($"-2 / {length}", $"H.X / {length}");
        // the side through the top square's top left corner (-2.5, 4)
        string s = d.Num($"(5 + 4 * H.X) / {length}", "s");
        var right = d.Pt(s * along);
        var left = d.Pt(s * across);
        var top = d.Pt(s * along + s * across);
        string cyan = d.Polygon(null, m, right, top, left);
        string small = d.Region(null, "Union",
            d.Polygon(null, d.Pt(-2.5, 2), d.Pt(-0.5, 2), d.Pt(-0.5, 4), d.Pt(-2.5, 4)),
            d.Polygon(null, d.Pt(-2, 0), m, d.Pt(0, 2), d.Pt(-2, 2)),
            d.Polygon(null, m, d.Pt(2, 0), d.Pt(2, 2), d.Pt(0, 2)));
        d.Region(d.Fill("Cyan", "#FF10E3EC", null), "Intersection", cyan, small);
        d.Polygon(d.Outline("CyanLine", "#FF1EBDC0", 2), m, right, top, left);
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        d.Box(frame, V.At(-2.5, 2), V.At(-0.5, 4));
        d.Box(frame, V.At(-2, 0), V.At(0, 2));
        d.Box(frame, V.At(0, 0), V.At(2, 2));
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(-3.6, 5.2, 2.2, -0.5);
    }

    // A parallelogram with sides 5 and diagonal 10; the angle a between the diagonal and the base and
    // the angle b outside the base's far corner add up to a right angle. Its area? Nothing to drag:
    // the condition makes tan a = 1/2.
    [Puzzle("parallelogram-area")]
    public static void ParallelogramArea(Lgf d)
    {
        double k = Math.Sqrt(5);
        var a = d.Pt(0, 0, "A");
        var b = d.Pt(3 * k, 0, "B");
        var c = d.Pt(4 * k, 2 * k, "C");
        var top = d.Pt(k, 2 * k, "D");
        var beyond = d.Pt(3 * k + 1.6, 0);
        d.Polygon(d.Fill("Lime", "#FFDAFC3A", width: 3.5), a, b, c, top);
        string green = "#FF0ED018";
        // over the lime, which stays bright in the dark theme: deep there
        string line = d.Line("Green", green, 3, dark: "#FF0A6E12");
        d.Segment(line, a, c);
        string outside = d.Line("Extension", green, 3);
        d.Segment(outside, b, beyond);
        d.ArcMark(d.Line("MarkOnFill", green, 2.5, dark: "#FF0A6E12"), a, b, c, 1.0);
        d.ArcMark(outside, b, beyond, c, 0.65);
        d.Label("10", V.At(4.35, 2.8), d.ShapeText("GreenMark", "#FF0A9A14", 36));
        d.Label("5", V.At(0.7, 2.75), d.Text("Mark", Lgf.Ink, 36));
        d.Label("a", V.At(1.5, 0.32), d.ShapeText("GreenLetter", "#FF0A9A14", 24));
        d.Label("b", b + V.At(1.0, 0.36), d.Text("GreenText", "#FF0A9A14", 24));
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(0, 2 * k, 4 * k, -0.6);
    }

    // Three squares: 32 and 18 corner to corner, and a tilted one (25) whose diagonal joins the 32's top
    // left corner to the 18's. How long is the red line, from the tilted square's far corner to the
    // corner the other two share? Nothing to drag: the areas fix it.
    [Puzzle("three-squares")]
    public static void ThreeSquares(Lgf d)
    {
        double a = 4 * Math.Sqrt(2), b = 3 * Math.Sqrt(2);
        var p = d.Pt(a, a, "P");
        var q = d.Pt(0, a, "Q");
        var u = d.Pt(a, a + b, "U");
        var v = d.Pt(a / 2 - b / 2, a + b / 2 + a / 2, "V");
        var low = d.Pt(a / 2 + b / 2, a + b / 2 - a / 2);
        string cyan = d.Fill("Cyan", "#FF38EEFB", width: 4);
        d.Box(cyan, V.At(0, 0), V.At(a, a));
        d.Box(cyan, V.At(a, a), V.At(a + b, a + b));
        d.Polygon(cyan, q, low, u, v);
        string red = "#FFDA0318";
        string arrow = d.Line("Red", red, 3);
        var middle = Mid(d, v, p);
        d.Vector(arrow, middle, v);
        d.Vector(arrow, middle, p);
        d.Label("?", V.At(3.65, 8.6), d.ShapeText("RedMark", red, 32));
        string blue = d.ShapeText("BlueMark", "#FF0A28E0", 36);
        d.Label("25", V.At(1.9, 7.3), blue);
        d.Label("18", V.At(a + b / 2, a + b / 2), blue);
        d.Label("32", V.At(a / 2, a / 2), blue);
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(0, a + b / 2 + a / 2, a + b, -0.8);
    }

    // Three squares: a black one, a green one (area 16) on its bottom left corner, and a blue one on its
    // top right corner sharing the green one's far corner, with its next corner on the black one's
    // bottom side. The blue area? That corner on the bottom puts the shared corner on the black square's
    // other diagonal, so the blue side is the green one's diagonal. Dragging the shared corner round the
    // green one's turns it and resizes the black square: the blue square stays 32.
    [Puzzle("three-squares-iv")]
    public static void ThreeSquaresIV(Lgf d)
    {
        double diagonal = 4 * Math.Sqrt(2);
        var o = d.Pt(0, 0, "O");
        var x = d.ArcSlider(o, diagonal, 72, 84, 77, (0, 0), "X");
        string s = d.Num("X.X + X.Y", "s");
        V[] green = [o, d.Pt(new V("(X.X + X.Y) / 2", "(X.Y - X.X) / 2")), x, d.Pt(new V("(X.X - X.Y) / 2", "(X.Y + X.X) / 2"))];
        V[] blue = [x, d.Pt(new V(s, s)), d.Pt(new V("2 * X.X + X.Y", "X.X")), d.Pt(new V("2 * X.X", "0"))];
        string greenShape = d.Polygon(d.Fill("Green", "#FF3FE972", null), green);
        string blueShape = d.Polygon(d.Fill("Blue", "#FF3E78DD", null), blue);
        d.Region(d.Fill("Teal", "#FF66CAB6", null), "Intersection", greenShape, blueShape);
        d.Polygon(d.Outline("GreenLine", "#FF10D870", 3), green);
        d.Polygon(d.Outline("BlueLine", "#FF3D33C4", 3), blue);
        d.Box(d.Outline("Frame", Lgf.Ink, 4), V.At(0, 0), new V(s, s));
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(-2.15, 6.8, 8.5, -0.6);
    }

    // A rectangle, the midpoints of its sides (the ticks), and four lines: from two opposite corners to
    // the midpoints of the far sides, and from the top and bottom midpoints to the side ones. What
    // fraction is shaded (the parallelogram they make)? Dragging the top right corner changes the
    // rectangle: the fraction stays.
    [Puzzle("parallelogram-in-a-rectangle")]
    public static void ParallelogramInARectangle(Lgf d)
    {
        d.Texture = "FeltTip";
        var bl = d.Pt(0, 0, "BL");
        var tr = d.HandleIn(d.Polygon(null, d.Pt(3.8, 3.8), d.Pt(6.9, 3.8), d.Pt(6.9, 6.4), d.Pt(3.8, 6.4)), 6, 6, "H");
        var br = d.Pt(new V("H.X", "0"), "BR");
        var tl = d.Pt(new V("0", "H.Y"), "TL");
        var left = d.Pt(new V("0", "H.Y / 2"), "L");
        var right = d.Pt(new V("H.X", "H.Y / 2"), "R");
        var top = d.Pt(new V("H.X / 2", "H.Y"), "Tm");
        var bottom = d.Pt(new V("H.X / 2", "0"), "Bm");
        var upper = d.Crossing(tl, right, top, left, "X1");
        var lower = d.Crossing(br, left, bottom, right, "X2");
        d.Polygon(d.Fill("Orange", "#FFF55A02", null), left, upper, right, lower);
        string ink = d.Line("Ink", Lgf.Ink, 3.5);
        d.Segment(ink, tl, right);
        d.Segment(ink, br, left);
        d.Segment(ink, top, left);
        d.Segment(ink, bottom, right);
        foreach (var (from, middle, to) in new[] { (bl, bottom, br), (br, right, tr), (tr, top, tl), (tl, left, bl) })
        {
            d.Segment(ink, from, middle, "TwoTicks");
            d.Segment(ink, middle, to, "TwoTicks");
        }

        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(0, 6, 6, -0.7);
    }
}
