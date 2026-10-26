#nullable disable

static partial class Puzzles
{
    // A square in a square: a line from each corner of the big one, turned 15° from its side, to the
    // foot of the next one, the four making a pinwheel of right triangles around the blue square. The
    // blue square covers half the big one, which fixes the turn: the angle at the bottom left corner?
    // Nothing to drag.
    [Puzzle("a-square-in-a-square")]
    public static void ASquareInASquare(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 4;
        double c = Math.Cos(15 * Math.PI / 180);
        // the corners counterclockwise from the bottom left, and the right angle of each corner's triangle
        // with the next one: c L along the line turned 15° from the side
        (double X, double Y)[] at = [(0, 0), (L, 0), (L, L), (0, L)];
        var corners = at.Select(p => d.Pt(p.X, p.Y)).ToArray();
        var feet = Enumerable.Range(0, 4).Select(k =>
        {
            double angle = (15 + 90 * k) * Math.PI / 180;
            return d.Pt(at[k].X + L * c * Math.Cos(angle), at[k].Y + L * c * Math.Sin(angle));
        }).ToArray();
        string navy = "#FF22355E";
        d.Polygon(d.Fill("Blue", "#FF03A7EC", navy), feet);
        string line = d.Line("Navy", navy, 3);
        for (int k = 0; k < 4; k++)
        {
            d.Segment(line, corners[k], corners[(k + 1) % 4]);
            d.Segment(line, corners[k], feet[k]);
        }

        d.ArcMark(line, corners[0], corners[1], feet[0], 1.4);
        d.Label("?", V.At(1.8, 0.28), d.Text("NavyMark", navy, 30));
        d.Scene(0, L, L, 0);
    }

    // Two rectangles side by side on one line, the blue one 1 by 5 and the yellow one 2 by 3, sharing a
    // bottom corner; the angle there between their diagonals? Nothing to drag: the lengths fix it.
    [Puzzle("two-rectangles")]
    public static void TwoRectangles(Lgf d)
    {
        var o = d.Pt(0, 0, "O");
        var blueTop = d.Pt(-1, 5, "B");
        var yellowTop = d.Pt(2, 3, "Y");
        d.Box(d.Fill("Blue", "#FFA5D8F3", "#FF4FB0D8", 2), V.At(-1, 0), V.At(0, 5));
        d.Box(d.Fill("Yellow", "#FFFFFFA4", "#FFF2DA3A", 2), V.At(0, 0), V.At(2, 3));
        // on the fills, which are mid-tones in the dark theme: ink there too
        string ink = d.Line("Ink", Lgf.Ink, 3.5, dark: Lgf.Ink);
        d.Segment(ink, blueTop, o);
        d.Segment(ink, o, yellowTop);
        d.Angle(d.Line("AngleMark", Lgf.Ink, 2.5, dark: Lgf.Ink), o, yellowTop, blueTop, 30);
        d.Label("?", V.At(0.12, 0.85), d.ShapeText("ShapeMark", Lgf.Ink, 28));
        string mark = d.Text("Mark", Lgf.Ink, 30);
        d.Label("5", V.At(-1.45, 2.5), mark);
        d.Label("3", V.At(2.42, 1.6), mark);
        d.Label("1", V.At(-0.5, -0.4), mark);
        d.Label("2", V.At(1, -0.4), mark);
        d.Scene(-1.7, 5, 2.7, -0.7);
    }

    // An isosceles triangle lying on its side, its apex at the left, cut into four by a zigzag from the top
    // side to the bottom and back: the first piece of the top side, the three zigzag lines and the short
    // right side all equal (the ticks). The angle at the apex? The equal lengths make the angles there
    // x, 2x, 3x, 4x along the zigzag, so x is 20°: nothing to drag.
    [Puzzle("subdivided-triangle-iii")]
    public static void SubdividedTriangleIII(Lgf d)
    {
        d.Texture = "FeltTip";
        const double u = 2.1;
        double x = 20 * Math.PI / 180, side = u * (1 + 2 * Math.Cos(2 * x) + 2 * Math.Cos(4 * x));
        double cx = Math.Cos(x / 2), sx = Math.Sin(x / 2);
        // along the top side and along the bottom side from the apex
        V Up(double length) => V.At(length * cx, length * sx);
        V Down(double length) => V.At(length * cx, -length * sx);
        var a = d.Pt(0, 0, "A");
        var t = d.Pt(Up(side), "T");
        var b = d.Pt(Down(side), "B");
        var p = d.Pt(Up(u), "D");
        var e = d.Pt(Down(2 * u * Math.Cos(x)), "E");
        var f = d.Pt(Up(u + 2 * u * Math.Cos(2 * x)), "F");
        d.Polygon(d.Fill("Yellow", "#FFFEF15B", null), a, p, e);
        d.Polygon(d.Fill("Orange", "#FFFE8850", null), p, f, e);
        d.Polygon(d.Fill("Red", "#FFEB545A", null), e, f, b);
        d.Polygon(d.Fill("Pink", "#FFFFB2D7", null), f, t, b);
        // the outline light on the dark paper, the zigzag across the fills (mid-tones there) in ink
        string frame = d.Line("Frame");
        d.Segment(frame, a, p, "TwoTicks");
        d.Segment(frame, p, t);
        d.Segment(frame, t, b, "TwoTicks");
        d.Segment(frame, b, a);
        string ink = d.Line("Ink", Lgf.Ink, 2.5, dark: Lgf.Ink);
        d.Segment(ink, p, e, "TwoTicks");
        d.Segment(ink, e, f, "TwoTicks");
        d.Segment(ink, f, b, "TwoTicks");
        d.ArcMark(ink, a, b, t, 0.5);
        d.Label("?", V.At(0.74, -0.03), d.ShapeText("ShapeMark", Lgf.Ink, 26));
        d.Scene(0, side * sx, side * cx, -side * sx);
    }

    // Three squares: a blue one, and two purple ones stacked into a tilted rectangle standing on the blue
    // one's bottom right corner, the line between them through the blue one's top right corner. A line
    // from the rectangle's top left corner to the blue one's top left corner; the angle there? The left
    // end M of the line between the purple squares sees the blue one's right side at a right angle, so it
    // is on the circle on that side: dragging it along the circle tilts and resizes the purple squares,
    // and the angle stays.
    [Puzzle("three-squares-v")]
    public static void ThreeSquaresV(Lgf d)
    {
        d.Texture = "FeltTip";
        const double b = 4;
        var o = d.Pt(0, 0, "O");
        var c = d.Pt(b, 0, "C");
        var topRight = d.Pt(b, b, "E");
        var topLeft = d.Pt(0, b, "D");
        var center = d.Pt(b, b / 2, "K");
        var m = d.ArcSlider(center, b / 2, 128, 160, 145.6, (b, b / 2), "M");
        // across the rectangle: the side from C to M turned a quarter clockwise
        var across = new V("M.Y - C.Y", "C.X - M.X");
        var right = d.Pt(c + across, "R");
        var middleRight = d.Pt(m + across, "N");
        var apex = d.Pt(new V("2 * M.X - C.X", "2 * M.Y - C.Y"), "A");
        var top = d.Pt(apex + across, "T");
        string blueSquare = d.Polygon(d.Fill("Blue", "#FF629DFE", null), o, c, topRight, topLeft);
        string rectangle = d.Polygon(d.Fill("Purple", "#FFD4ABF6", null), c, right, top, apex);
        d.Region(d.Fill("Overlap", "#FF6982F7", null), "Intersection", blueSquare, rectangle);
        string frame = d.Outline("Frame", Lgf.Ink, 3.5);
        d.Inked(frame, o, c, topRight, topLeft);
        d.Inked(frame, c, right, top, apex);
        d.Segment(d.Line("Ink", Lgf.Ink, 3.5), m, middleRight);
        string purple = "#FF4A0E94";
        string line = d.Line("PurpleLine", purple, 2.5);
        d.Segment(line, apex, topLeft);
        d.Angle(line, apex, topLeft, c, 40);
        d.Label("?", 0.6 * apex + 0.3 * topLeft + 0.1 * c, d.Text("PurpleMark", purple, 30));
        d.Scene(0, 8.4, 7.6, 0);
    }

    // An equilateral triangle, a big semicircle on its bottom side (touching its left side) and a small
    // one on its right side, the two touching. Lines from the touching point to the big one's left end
    // and the small one's top end; the angle between them? Dragging the touching point along the big
    // semicircle moves and resizes the small one (its center where the line from the big one's center
    // through the touching point meets the right side): the angle stays.
    [Puzzle("two-semi-circles-in-an-equilateral-triangle")]
    public static void TwoSemiCirclesInAnEquilateralTriangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double side = 6, x1 = 2.6, r1 = x1 * S3;
        var left = d.Pt(0, 0, "V0");
        var right = d.Pt(side, 0, "V1");
        var apex = d.Pt(side / 2, side * S3, "V2");
        var o1 = d.Pt(x1, 0, "O1");
        var p1 = d.Pt(x1 - r1, 0, "P1");
        var q1 = d.Pt(x1 + r1, 0, "Q1");
        var t = d.ArcSlider(o1, r1, 48, 69, 67, (x1, 0), "T");
        var o2 = d.Crossing(o1, t, right, apex, "O2");
        string r2 = d.Num("distance(O2, T)", "r2");
        var up = new V("-0.5", s3);
        var p2 = d.Pt(o2 + r2 * up, "P2");
        var q2 = d.Pt(o2 - r2 * up, "Q2");
        string rim = "#FF2C98D4";
        d.Arc("CircleSegment", d.Fill("Blue", "#FF8ACAFE", rim, 2.5), o1, q1, p1);
        d.Arc("CircleSegment", d.Fill("LightBlue", "#FFA4E0FE", rim, 2.5), o2, p2, q2);
        d.Inked(d.Outline("Frame", "#FF4A4FB5", 3), left, right, apex);
        string ink = d.Line("Ink", Lgf.Ink, 3.5);
        d.Segment(ink, p1, t);
        d.Segment(ink, t, p2);
        d.Angle(d.Line("AngleMark"), t, p2, p1, 34);
        // the two sides' directions add up to the bisector's, of length 1 (the angle is 120°)
        string toP1 = d.Num("distance(T, P1)"), toP2 = d.Num("distance(T, P2)");
        var bisector = new V($"(P1.X - T.X) / {toP1} + (P2.X - T.X) / {toP2}", $"(P1.Y - T.Y) / {toP1} + (P2.Y - T.Y) / {toP2}");
        d.Label("?", t + 0.85 * bisector, d.Text("Mark", Lgf.Ink, 30));
        d.Scene(0, side * S3, side, 0);
    }

    // Three regular hexagons: an orange and a red one of the same size sharing a side, and a purple one
    // hanging from the red one's right corner, the red one's lower right side along its upper left side.
    // Lines from the orange one's top left corner to the purple one's right corner and from the orange
    // one's right corner to the purple one's top right corner; the angle where they cross? Dragging the
    // purple one's top right corner resizes it: the angle stays.
    [Puzzle("three-regular-hexagons-iv")]
    public static void ThreeRegularHexagonsIV(Lgf d)
    {
        const double s = 2;
        var orange = HexagonFromLeft(d, V.At(-s, 0), Lgf.N(s));
        var red = HexagonFromLeft(d, V.At(s / 2, S3 * s), Lgf.N(s));
        var corner = red[3];
        var h = d.Slider(V.At(2.5 * s + 1.2 * s, S3 * s), V.At(2.5 * s + 1.8 * s, S3 * s), (1.54 - 1.2) / 0.6, (2.5 * s + 1.54 * s, S3 * s), "H");
        string p = d.Num($"H.X - {Lgf.N(2.5 * s)}", "p");
        var purple = HexagonFromLeft(d, corner + new V($"-0.5 * {p}", $"-{s3} * {p}"), p);
        d.Polygon(d.Fill("Orange", "#FFFDE3B5", "#FFFB775B", 3), orange);
        d.Polygon(d.Fill("Red", "#FFFDCAD1", "#FFD8575B", 3), red);
        d.Polygon(d.Fill("Purple", "#FFDDC9EF", "#FF684792", 3), purple);
        var x = d.Crossing(orange[1], purple[3], orange[3], h, "X");
        string ink = d.Line("Ink", Lgf.Ink, 3.5);
        d.Segment(ink, orange[1], purple[3]);
        d.Segment(ink, orange[3], h);
        d.Angle(d.Line("AngleMark", Lgf.Ink, 2.5), x, h, orange[1], 30);
        d.Label("?", x + V.At(-0.15, 0.9), d.ShapeText("ShapeMark", Lgf.Ink, 30));
        // around the purple hexagon at its largest
        d.Scene(-s, 2 * S3 * s, 2.5 * s + 2.7 * s, S3 * s - 3.6 * S3 * s);
    }

    // Five regular polygons: the red square (area 12) with the yellow one beside it, a green equilateral
    // triangle on the red one's top left corner, a purple one from the green one's top to the red one's
    // top right corner, and the blue square standing on the line of the red one's top, its top left
    // corner the purple triangle's third corner. The blue square's area? Its side is always the height
    // of an equilateral triangle on the red square's side: dragging the green triangle's top resizes it
    // and moves the purple triangle and the blue square, whose size stays.
    [Puzzle("five-regular-polygons")]
    public static void FiveRegularPolygons(Lgf d)
    {
        d.Texture = "FeltTip";
        const double a = 4, y = 2 * a / 3, b = S3 * a;
        var topLeft = d.Pt(0, a, "TL");
        var r = d.Pt(a, a, "R");
        var g = d.Slider(V.At(0.25 * a, a + 0.5 * a * S3), V.At(0.4 * a, a + 0.8 * a * S3), (0.62 - 0.5) / 0.3, (0.31 * a, a + 0.62 * a * S3), "G");
        // the green triangle's side is twice the top's distance right of the red square's corner
        var foot = d.Pt(new V("2 * G.X", Lgf.N(a)), "F");
        var p = d.Pt(new V($"{Lgf.N(a / 2)} + 2 * G.X", Lgf.N(a + b)), "P");
        d.Box(d.Fill("Red", "#FFE84F7F", Lgf.Ink, 4), V.At(0, 0), V.At(a, a));
        d.Box(d.Fill("Yellow", "#FFFFF877", Lgf.Ink, 4), V.At(a, a - y), V.At(a + y, a));
        d.Polygon(d.Fill("Green", "#FF48C027", Lgf.Ink, 4), topLeft, foot, g);
        d.Polygon(d.Fill("Purple", "#FFBB8BF0", Lgf.Ink, 4), g, r, p);
        d.Box(d.Fill("Blue", "#FF037CF7", Lgf.Ink, 4), new V("P.X", Lgf.N(a)), new V($"P.X + {Lgf.N(b)}", "P.Y"));
        d.Label("12", V.At(a / 2, a / 2), d.ShapeText("ShapeMark", Lgf.Ink, 40));
        d.Label("?", new V($"P.X + {Lgf.N(b / 2)}", Lgf.N(a + b / 2)), d.ShapeText("ShapeMark", Lgf.Ink, 40));
        d.Scene(0, a + b, a / 2 + 0.8 * a + b, 0);
    }

    // Three squares, 4 and 3 side by side and a small one on the big one's top left corner, and a yellow
    // rectangle with a corner where the two big ones meet, the next corner at the small one's top right
    // corner and its far side through the 3 square's top right corner. Its area? It is twice the triangle
    // of those three corners, 24 for any small square: dragging the small square's corner resizes it and
    // turns the rectangle, whose area stays.
    [Puzzle("three-squares-and-a-rectangle")]
    public static void ThreeSquaresAndARectangle(Lgf d)
    {
        d.Texture = "FeltTip";
        var c = d.Pt(4, 0, "C");
        var q = d.Slider(V.At(0.3, 4.3), V.At(1.8, 5.8), (1 - 0.3) / 1.5, (1, 5), "Q");
        string length = d.Num("distance(C, Q)", "l");
        string width = d.Num($"24 / {length}", "w");
        // across the rectangle: the side from C to Q turned a quarter clockwise, as long as the width
        var across = new V($"{width} * (Q.Y - C.Y) / {length}", $"{width} * (C.X - Q.X) / {length}");
        string red = "#FFDF472B";
        d.Polygon(d.Fill("Yellow", "#FFFEEC70", red, 3), c, q, d.Pt(q + across), d.Pt(c + across));
        string frame = d.Outline("Frame", Lgf.Ink, 3.5);
        d.Box(frame, V.At(0, 0), V.At(4, 4));
        d.Box(frame, V.At(4, 0), V.At(7, 3));
        d.Box(frame, V.At(0, 4), q);
        string arrow = d.Line("Arrow", Lgf.Ink, 2.5);
        string mark = d.Text("Mark", Lgf.Ink, 30);
        d.Dimension(arrow, mark, c, d.Pt(0, 0), 0.5, "4", 0.4);
        d.Dimension(arrow, mark, d.Pt(7, 0), c, 0.5, "3", 0.4);
        d.Scene(0, 7.2, 7.8, -0.8);
    }

    // A right isosceles triangle, a point on its long side 8 from the right angle, and three quarter
    // circles: two around that point out to the long side's ends (on the far side of it, sharing the
    // radius square to it) and one around the right angle from end to end. The shaded total? The two
    // around the point have the squares of their radii adding up to twice 8 squared. Dragging the point
    // around the right angle at 8 resizes the triangle: the total stays.
    [Puzzle("three-quarter-circles-ii")]
    public static void ThreeQuarterCirclesII(Lgf d)
    {
        d.Texture = "FeltTip";
        var c = d.Pt(0, 0, "C");
        var x = d.ArcSlider(c, 8, 110, 160, 121.7, (0, 0), "X");
        string r = d.Num("X.Y - X.X", "R");
        var a = d.Pt(new V($"-{r}", "0"), "A");
        var b = d.Pt(new V("0", r), "B");
        // across the long side, away from the right angle
        var up = new V($"-{0.5 * Math.Sqrt(2)}", Lgf.N(0.5 * Math.Sqrt(2)));
        var k = d.Pt(x + d.Num("distance(X, A)", "xa") * up, "K");
        var k2 = d.Pt(x + d.Num("distance(X, B)", "xb") * up, "K2");
        string purple = "#FF6A35C2";
        string lavender = d.Fill("Lavender", "#FFC7ADE1", null);
        d.Region(lavender, "Union", d.Arc("CircleSector", null, x, k, a, visible: false), d.Arc("CircleSector", null, x, b, k2, visible: false));
        string line = d.Line("Purple", purple, 3);
        d.Arc("CircleArc", line, x, k, a);
        d.Arc("CircleArc", line, x, b, k2);
        d.Arc("CircleArc", line, c, b, a);
        d.Segment(line, a, c);
        d.Segment(line, c, b);
        d.Segment(line, a, b);
        d.Segment(line, x, k);
        d.Segment(line, x, k2);
        var middle = Mid(d, c, x);
        string arrow = d.Line("Arrow", Lgf.Ink, 3);
        d.Vector(arrow, middle, c);
        d.Vector(arrow, middle, x);
        d.Label("8", middle + 0.15 * new V("-X.Y", "X.X"), d.Text("Mark", Lgf.Ink, 34));
        d.Scene(-14, 15, 0.3, 0);
    }

    // Two equilateral triangles: a blue one standing on a corner, its right side upright, and a green one
    // from its top corner down that side. Lines from the blue one's left corner to the green one's bottom
    // corner and from the middle of the blue one's lower left side to the middle of the green one's
    // lower right side (the ticks); the angle where they cross? Dragging the green one's bottom corner up
    // and down the blue one's side resizes it: the angle stays.
    [Puzzle("two-equilateral-triangles-ii")]
    public static void TwoEquilateralTrianglesII(Lgf d)
    {
        d.Texture = "FeltTip";
        const double s = 6;
        var l = d.Pt(0, 0, "L");
        var t = d.Pt(S3 * s, s / 2, "T");
        var b = d.Pt(S3 * s, -s / 2, "B");
        var g = d.Slider(V.At(S3 * s, s / 2 - 0.4 * s), V.At(S3 * s, s / 2 - 0.95 * s), (0.75 - 0.4) / 0.55, (S3 * s, s / 2 - 0.75 * s), "G");
        string side = d.Num($"{Lgf.N(s / 2)} - G.Y", "g");
        var r = d.Pt(t + new V($"{s3} * {side}", $"-0.5 * {side}"), "R");
        var m = d.DotAt(Mid(d, l, b), "D");
        var e = d.DotAt(Mid(d, g, r), "E");
        string navy = "#FF1B2AA0", green = "#FF3BAA6A";
        d.Polygon(d.Fill("Blue", "#FF9DD7FC", navy, 4), l, t, b);
        d.Polygon(d.Fill("Green", "#FFC4F0A0", green, 4), t, r, g);
        string navyLine = d.Line("Navy", navy, 4);
        d.Segment(navyLine, l, m, "OneTick");
        d.Segment(navyLine, m, b, "OneTick");
        d.Segment(navyLine, t, b);
        string greenLine = d.Line("GreenLine", green, 4);
        d.Segment(greenLine, g, e, "TwoTicks");
        d.Segment(greenLine, e, r, "TwoTicks");
        var x = d.Crossing(l, g, m, e, "X");
        // on the fills, which are mid-tones in the dark theme: ink there too
        string ink = d.Line("Ink", Lgf.Ink, 4.5, dark: Lgf.Ink);
        d.Segment(ink, l, g);
        d.Segment(ink, m, e);
        d.Angle(d.Line("AngleMark", Lgf.Ink, 3, dark: Lgf.Ink), x, e, l, 30);
        d.Label("?", x + V.At(0, 1.15), d.ShapeText("ShapeMark", Lgf.Ink, 32));
        d.Scene(0, s / 2, S3 * s + S3 * 0.95 * s, -s / 2);
    }

    // Two semicircles on one center, the big one up and the small one down, and two quarter circles
    // around a point T on the big one's arc, out to the small one's ends A and B (the triangle ABT
    // between them). Inside exactly one shape is colored, inside two is white: yellow outside the big
    // semicircle, blue inside it and in the small one. More yellow or blue? The quarter circles' radii
    // squared add up to twice the semicircles' (TA² + TB² = 2(R² + r²)), so they are the same: dragging
    // T along the arc, or B along the diameter (the small semicircle's size), keeps them equal.
    [Puzzle("two-semi-circles-and-two-quarter-circles")]
    public static void TwoSemiCirclesAndTwoQuarterCircles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4;
        var m = d.Pt(0, 0, "M");
        var east = d.Pt(R, 0, "P2");
        var west = d.Pt(-R, 0, "P1");
        var t = d.ArcSlider(m, R, 86, 112, 101.9, (0, 0), "T");
        var b = d.Slider(V.At(1.8, 0), V.At(3, 0), (2.5 - 1.8) / 1.2, (2.5, 0), "B");
        var a = d.Pt(new V("-B.X", "0"), "A");
        // the quarter circles' other ends: TA turned a quarter clockwise, TB a quarter counterclockwise
        var k1 = d.Pt(new V("T.X + A.Y - T.Y", "T.Y - A.X + T.X"), "K1");
        var k2 = d.Pt(new V("T.X - B.Y + T.Y", "T.Y + B.X - T.X"), "K2");
        string big = d.Arc("CircleSegment", null, m, east, west, visible: false);
        string left = d.Arc("CircleSector", null, t, k1, a, visible: false);
        string right = d.Arc("CircleSector", null, t, b, k2, visible: false);
        string yellow = d.Fill("Yellow", "#FFFEF994", null);
        d.Region(yellow, "Difference", left, big);
        d.Region(yellow, "Difference", right, big);
        string blue = d.Fill("Blue", "#FF4AAAFD", null);
        d.Region(blue, "Difference", big, left, right);
        d.Arc("CircleSegment", blue, m, a, b);
        string gray = d.Line("Gray", "#FF5A5A5A", 2.5);
        d.Arc("CircleArc", gray, m, east, west);
        d.Arc("CircleArc", gray, t, k1, a);
        d.Arc("CircleArc", gray, t, b, k2);
        d.Segment(gray, t, k1);
        d.Segment(gray, t, k2);
        string navy = d.Line("Navy", "#FF1C4FA8", 2.5);
        d.Arc("CircleArc", navy, m, a, b);
        d.Segment(navy, t, a);
        d.Segment(navy, t, b);
        d.Segment(navy, west, a, "TwoTicks");
        d.Segment(navy, a, b);
        d.Segment(navy, b, east, "TwoTicks");
        d.Scene(-5.2, 7.5, 4.7, -2.6);
    }
}
