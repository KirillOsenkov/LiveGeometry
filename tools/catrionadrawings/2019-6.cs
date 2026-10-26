#nullable disable

static partial class Puzzles
{
    // A blue square; a pink square of area 4 standing on its corner on the top side, its diagonal
    // along the side from the top left corner, and a second one of area 4 with its diagonal from that
    // square's bottom corner to the big square's bottom left corner. The blue triangle from the first
    // square's right corner to the big square's right side? Nothing to drag: the second square's
    // diagonal (2√2) fixes the big square's side at √2 + √6. (Her picture has the big square a little
    // larger than that.)
    [Puzzle("two-squares-overlapping-a-third-square")]
    public static void TwoSquaresOverlappingAThirdSquare(Lgf d)
    {
        d.Texture = "FeltTip";
        double r2 = Math.Sqrt(2), L = r2 + Math.Sqrt(6), m = (L - r2) / 2;
        var bl = d.Pt(0, 0);
        var br = d.Pt(L, 0);
        var tr = d.Pt(L, L);
        var tl = d.Pt(0, L);
        var p = d.Pt(2 * r2, L);
        var top = d.Pt(r2, L + r2);
        var q = d.Pt(r2, L - r2);
        var left = d.Pt(r2 / 2 - m, m + r2 / 2);
        var right = d.Pt(r2 / 2 + m, m - r2 / 2);
        d.Polygon(d.Fill("Blue", "#FF03ACFD", null), p, tr, br);
        string pink = "#FFD0127F";
        string pinkLine = d.Outline("Pink", pink, 3.5);
        d.Polygon(pinkLine, tl, top, p, q);
        d.Polygon(pinkLine, q, left, bl, right);
        d.Inked(d.Outline("Frame", "#FF0410D8", 4.5), bl, br, tr, tl);
        string pinkMark = d.Text("PinkMark", pink, 34);
        d.Label("4", V.At(r2, L + 0.6), pinkMark, 34);
        d.Label("4", V.At(r2 / 2 - 0.05, m), pinkMark, 34);
        d.Label("?", V.At(L - 0.42, L - 0.85), d.ShapeText("NavyMark", "#FF1E28B4", 32), 32);
        d.Scene(r2 / 2 - m, L + r2, L, 0);
    }

    // A big square with a semicircle on its top side; a middle square on its bottom side, centered, its
    // top corners on the arc; the smallest square (area 7) between the middle one's top and the big
    // one's top. The shaded area (the big square less the two)? The arc makes the middle square 3/5 of
    // the big one and the small one 2/5. Dragging the small square slides it along the middle one's
    // top: the shaded area stays.
    [Puzzle("three-squares-and-a-semi-circle")]
    public static void ThreeSquaresAndASemiCircle(Lgf d)
    {
        var c = d.Pt(5, 10, "C");
        var east = d.Pt(10, 10);
        var west = d.Pt(0, 10);
        var a = d.Slider(V.At(2, 6), V.At(4, 6), 0.86, (3.72, 6), "A");
        string big = d.Box(null, V.At(0, 0), V.At(10, 10));
        string middle = d.Box(null, V.At(2, 0), V.At(8, 6));
        var b = d.Pt(new V("A.X + 4", "6"));
        var e = d.Pt(new V("A.X + 4", "10"));
        var f = d.Pt(new V("A.X", "10"));
        string small = d.Polygon(null, a, b, e, f);
        d.Region(d.Fill("Blue", "#FF8CF4FB", null), "Difference", big, middle, small);
        string navy = d.Outline("Navy", "#FF0A14B4", 3.5);
        d.Box(navy, V.At(2, 0), V.At(8, 6));
        d.Polygon(navy, a, b, e, f);
        d.Arc("CircleArc", d.Line("Arc", Lgf.Ink, 3), c, west, east);
        d.Inked(d.Outline("Frame", Lgf.Ink, 4.5), d.Pt(0, 0), d.Pt(10, 0), east, west);
        d.Label("7", new V("A.X + 2", "8"), d.Text("CyanMark", "#FF0DC4F5", 34), 34);
        d.Scene(0, 10, 10, 0);
    }

    /// <summary>A double arrow along a fixed segment from a to b, a gap at its middle (half of it either side) holding a text of the given size: Dimension with the text centered at any size</summary>
    private static void GappedDoubleArrow(Lgf d, string lineStyle, string textStyle, double size, V a, V b, string text, double gap)
    {
        var middle = Mid(d, a, b);
        string length = d.Num($"distance({a.X.Replace(".X", "")}, {b.X.Replace(".X", "")})");
        var along = new V($"({b.X} - {a.X}) / {length}", $"({b.Y} - {a.Y}) / {length}");
        d.Vector(lineStyle, d.Pt(middle - gap * along), a);
        d.Vector(lineStyle, d.Pt(middle + gap * along), b);
        d.Label(text, middle, textStyle, size);
    }

    // Three congruent rectangles: a blue and a red one upright side by side, 4 wide together, and a
    // yellow one tilted across both with a corner on each outer side, on the bottom and on the top.
    // Its diagonal? Nothing to drag: the corners on the four sides tilt it by 30° and make the
    // rectangles √3 times as tall as wide.
    [Puzzle("three-congruent-rectangles")]
    public static void ThreeCongruentRectangles(Lgf d)
    {
        d.Texture = "FeltTip";
        double h = 2 * Math.Sqrt(3);
        var a = d.Pt(0, h / 2);
        var b = d.Pt(1, 0);
        var c = d.Pt(4, h / 2);
        var e = d.Pt(3, h);
        string blueBox = d.Box(d.Fill("Blue", "#FF02D8FB", null), V.At(0, 0), V.At(2, h));
        string redBox = d.Box(d.Fill("Red", "#FFF42C10", null), V.At(2, 0), V.At(4, h));
        string yellow = d.Polygon(null, a, b, c, e);
        d.Region(d.Fill("Green", "#FFA5FC01", null), "Intersection", yellow, blueBox);
        d.Region(d.Fill("Orange", "#FFFECB01", null), "Intersection", yellow, redBox);
        d.Polygon(d.Outline("YellowLine", "#FFF5E600", 3, dark: "#FFF5E600"), a, b, c, e);
        d.Box(d.Outline("BlueLine", "#FF03D6EC", 3.5), V.At(0, 0), V.At(2, h));
        d.Box(d.Outline("RedLine", "#FFE60103", 3.5), V.At(2, 0), V.At(4, h));
        GappedDoubleArrow(d, d.Line("Arrow", Lgf.Ink, 4.5, dark: Lgf.Ink), d.ShapeText("ShapeMark", Lgf.Ink, 60), 60, b, e, "?", 0.62);
        GappedDoubleArrow(d, d.Line("Ink", Lgf.Ink, 4.5), d.Text("Mark", Lgf.Ink, 44), 44, d.Pt(0, -0.45), d.Pt(4, -0.45), "4", 0.3);
        d.Scene(0, h, 4, -0.7);
    }

    // A circle in a rectangle touching its sides and bottom, and a 3-4-5 triangle with its acute corner
    // where the circle touches the bottom and its right angle on the circle where the rectangle's top
    // crosses it; the hypotenuse crosses the circle again on the top. What fraction of the rectangle is
    // shaded (the triangle below the top)? Nothing to drag: the circle fixes it (the top is 0.8 of the
    // radius above the center).
    [Puzzle("rectangle-circle-and-triangle")]
    public static void RectangleCircleAndTriangle(Lgf d)
    {
        d.Texture = "FeltTip";
        var o = d.Pt(0, 0);
        var b = d.Pt(0, -1);
        var p = d.Pt(-0.6, 0.8);
        var q = d.Pt(0.6, 0.8);
        var a = d.Pt(0.75, 1.25);
        d.Polygon(d.Fill("Cyan", "#FF02FEFF", null), b, p, q);
        d.Box(d.Outline("Green", "#FF8EEC30", 3.5), V.At(-1, -1), V.At(1, 0.8));
        d.Circle(d.Line("Circle", Lgf.Ink, 3), o, b);
        d.Inked(d.Outline("Blue", "#FF0306D8", 4.5), b, p, a);
        d.Scene(-1, 1.25, 1, -1);
    }

    // Two regular hexagons: the red one of side 4, the blue one of side 3 with a corner at the red one's
    // right corner and the next but one on the red one's top side. How far is the blue corner between
    // them from the red one's top right corner? Nothing to drag: the sides fix it.
    [Puzzle("two-regular-hexagons")]
    public static void TwoRegularHexagons(Lgf d)
    {
        double r3 = Math.Sqrt(3), tx = 4 - Math.Sqrt(15);
        var red = Enumerable.Range(0, 6).Select(i => d.Pt(4 * Math.Cos(i * Math.PI / 3), 4 * Math.Sin(i * Math.PI / 3))).ToArray();
        // the blue corners: at the red one's right corner, the one between (the apex of a 30° isosceles
        // triangle on the short diagonal), on the red top; the center is the first plus the third less the second
        double mx = (4 + tx) / 2, my = r3, ux = -2 * r3 / Math.Sqrt(27), uy = -Math.Sqrt(15) / Math.Sqrt(27);
        double vx = mx + 1.5 * ux, vy = my + 1.5 * uy;
        double cx = 4 + tx - vx, cy = 2 * r3 - vy;
        double start = Math.Atan2(-cy, 4 - cx);
        var blue = Enumerable.Range(0, 6).Select(i => d.Pt(cx + 3 * Math.Cos(start - i * Math.PI / 3), cy + 3 * Math.Sin(start - i * Math.PI / 3))).ToArray();
        string redShape = d.Polygon(d.Fill("Pink", "#FFFDC5F5", null), red);
        string blueShape = d.Polygon(d.Fill("Cyan", "#FF40E8FC", null), blue);
        d.Region(d.Fill("Purple", "#FFC6B4F9", null), "Intersection", redShape, blueShape);
        d.Polygon(d.Outline("Red", "#FFC8061A", 3.5), red);
        string navy = "#FF0606D2";
        d.Polygon(d.Outline("Blue", navy, 3.5), blue);
        GappedDoubleArrow(d, d.Line("Arrow", navy, 3, dark: "#FF101070"), d.ShapeText("NavyMark", "#FF14147A", 32), 32, blue[1], red[1], "?", 0.48);
        d.Label("4", V.At(-3.75, 2.3), d.Text("RedMark", "#FFC8061A", 36), 36);
        d.Label("3", V.At(cx + 3.4, cy + 1.25), d.Text("BlueMark", navy, 36), 36);
        d.Scene(-4.2, cy + 3, cx + 3.75, -2 * r3);
    }

    // A semicircle and a circle centered on its arc, crossing it at P (left) and Q (right); lines from
    // the diameter's left end through P and from its right end through Q meet the circle again at X and
    // Y; the angle between the chords XQ and PY? Dragging the center along the arc, or the top of the
    // circle to change its size: the angle stays (90°, so a plain arc marks it).
    [Puzzle("angle-in-a-circle-and-semi-circle")]
    public static void AngleInACircleAndSemiCircle(Lgf d)
    {
        const double R = 5;
        var o = d.Pt(0, 0, "O");
        var a = d.Pt(-R, 0, "A");
        var b = d.Pt(R, 0, "B");
        var c = d.ArcSlider(o, R, 55, 80, 67.4, (0, 0), "C");
        var top = d.Slider(new V("C.X", $"C.Y + {0.35 * R}"), new V("C.X", $"C.Y + {0.75 * R}"), (0.61 - 0.35) / 0.4, (1.92, 4.62 + 0.61 * R), "H");
        string r = d.Num("H.Y - C.Y", "r");
        // P and Q: the circles' crossings, a along the line to the center and h either side of it
        string along = d.Num($"({2 * R * R} - {r} ^ 2) / {2 * R * R}", "k");
        string side = d.Num($"sqrt(1 - {along} ^ 2)", "w");
        var p = d.Pt(new V($"{along} * C.X - {side} * C.Y", $"{along} * C.Y + {side} * C.X"), "P");
        var q = d.Pt(new V($"{along} * C.X + {side} * C.Y", $"{along} * C.Y - {side} * C.X"), "Q");
        var x = SecondCrossing(d, a, p, c, "X");
        var y = SecondCrossing(d, b, q, c, "Y");
        string semi = "#FF0DA7D6";
        d.Arc("CircleArc", d.Line("Semi", semi, 3), o, b, a);
        d.Segment(d.Line("Diameter", semi, 3.5), a, b);
        d.Circle(d.Line("Circle", "#FF1948DF", 3.5), c, top);
        string teal = d.Line("Teal", "#FF10B4A0", 3);
        d.Segment(teal, a, x);
        d.Segment(teal, x, q);
        string green = d.Line("Green", "#FF9CD43C", 2.5);
        d.Segment(green, p, y);
        d.Segment(green, y, b);
        var z = d.Crossing(x, q, p, y, "Z");
        string redColor = "#FFE05A5C";
        d.ArcMark(d.Line("Red", redColor, 3), z, p, q, 0.3);
        string toP = d.Num("distance(Z, P)"), toQ = d.Num("distance(Z, Q)");
        var under = new V($"(P.X - Z.X) / {toP} + (Q.X - Z.X) / {toQ}", $"(P.Y - Z.Y) / {toP} + (Q.Y - Z.Y) / {toQ}");
        d.Label("?", z + (0.75 / Math.Sqrt(2)) * under, d.Text("RedMark", redColor, 28), 28);
        d.Scene(-R, 7.8, R, 0);
    }

    // A circle, an orange square whose bottom and right sides touch it, and a pink square with a corner
    // where the circle touches the bottom, the next corner on the circle (so its far side runs through
    // the circle's top); its opposite corner is 2 from the orange square's corner. The circle's area?
    // Dragging the pink square's corner along the circle resizes and turns it: its far corner stays the
    // radius from the orange corner.
    [Puzzle("circle-and-two-squares")]
    public static void CircleAndTwoSquares(Lgf d)
    {
        const double r = 2, s = 2.92;
        var o = d.Pt(0, 0, "O");
        var bottom = d.Pt(0, -r, "Bt");
        var k = d.Pt(r, -r, "K");
        var m = d.ArcSlider(o, r, 108, 172, 139.7, (0, 0), "M");
        var side = m - bottom;
        var far = d.Pt(bottom + new V(side.Y, V.Neg(side.X)), "F");
        var fourth = d.Pt(m + new V(side.Y, V.Neg(side.X)));
        d.Box(d.Outline("Orange", "#FFE06A1E", 3), V.At(r - s, -r), V.At(r, s - r));
        d.Circle(d.Line("Circle", Lgf.Ink, 3.5), o, bottom);
        d.Inked(d.Outline("Magenta", "#FFE0208C", 3), bottom, m, fourth, far);
        string blue = "#FF3A22E0";
        GappedDoubleArrow(d, d.Line("Arrow", blue, 3.5), d.Text("BlueMark", blue, 30), 30, k, far, "2", 0.3);
        d.Scene(-r, 2.75, 3.6, -2.35);
    }

    // A regular hexagon cut by a line from its top left corner to the middle of its bottom side, and
    // two lines from a point on it to the top right and the right corners. Each yellow part is a
    // third of the hexagon; what fraction is green? Nothing to drag: the thirds put the point 2/5 of
    // the way down the line (the line itself is fixed by the left part being a third).
    [Puzzle("divided-hexagon")]
    public static void DividedHexagon(Lgf d)
    {
        d.Texture = "FeltTip";
        double h = S3 * 4;
        var right = d.Pt(4, 0);
        var topRight = d.Pt(2, h);
        var topLeft = d.Pt(-2, h);
        var left = d.Pt(-4, 0);
        var bottomLeft = d.Pt(-2, -h);
        var bottomRight = d.Pt(2, -h);
        var q = d.Pt(0, -h);
        var p = d.Pt(-1.2, 0.2 * h);
        string yellow = d.Fill("Yellow", "#FFFFFE44", null);
        d.Polygon(yellow, topLeft, left, bottomLeft, q, p);
        d.Polygon(yellow, p, q, bottomRight, right);
        d.Polygon(d.Fill("Green", "#FF9AFC5C", null), topLeft, topRight, p);
        string ink = d.Line("Ink", "#FF2A2A1A", 6, dark: Lgf.DarkInk);
        d.Segment(ink, topLeft, q);
        d.Segment(ink, p, topRight);
        d.Segment(ink, p, right);
        d.Inked(d.Outline("Frame", "#FF2A2A1A", 6, dark: Lgf.DarkInk), right, topRight, topLeft, left, bottomLeft, bottomRight);
        d.Scene(-4, h, 4, -h);
    }

    // Two squares: a cyan one, and a navy one with a side from the cyan one's top left corner to a point
    // on its bottom side 5 from the corner; a dashed line through the cyan one's top right corner
    // parallel to that side cuts off the green strip. Its area? Dragging the cyan square's corner
    // resizes it (the 5 stays): the strip gets longer and thinner, its area stays 25.
    [Puzzle("two-overlapping-squares")]
    public static void TwoOverlappingSquares(Lgf d)
    {
        d.Texture = "FeltTip";
        var tr = d.Slider(V.At(7, 7), V.At(12, 12), (10.8 - 7) / 5, (10.8, 10.8), "TR");
        string a = "TR.X";
        var tl = d.Pt(new V("0", a), "TL");
        var b = d.Pt(5, 0, "B");
        // the navy square's other side from each end: the side TL to B turned a quarter counterclockwise
        var turn = new V(a, "5");
        var far = d.Pt(b + turn, "F");
        var top = d.Pt(tl + turn, "N");
        // where the dashed line through the cyan corner is, as a share of that side
        string share = d.Num($"{a} ^ 2 / (25 + {a} ^ 2)", "y");
        var dashTop = d.Pt(tl + share * turn);
        var dashBottom = d.Pt(b + share * turn);
        d.Polygon(d.Fill("Green", "#FF01ECBB", null), dashTop, top, far, dashBottom);
        d.Box(d.Outline("Cyan", "#FF04DAF6", 5), V.At(0, 0), new V(a, a));
        string navy = d.Outline("Navy", "#FF1414C8", 5);
        d.Polygon(navy, tl, b, far, top);
        d.Dashed(d.Line("Dash", Lgf.Ink, 4.5), dashTop, dashBottom, 6);
        GappedDoubleArrow(d, d.Line("Arrow", Lgf.Ink, 4), d.Text("Mark", Lgf.Ink, 34), 34, d.Pt(0, -1.3), d.Pt(5, -1.3), "5", 0.8);
        d.Scene(0, 16, 16, -2.1);
    }

    // A right triangle with a square in its right angle and another on its long side, standing under
    // the first one's inner corner; each square's perimeter is a third of the triangle's. What
    // fraction is shaded? The squares make it a 3-4-5 triangle with squares of side 1. Dragging the
    // lower square slides it along the bottom (its top still under the corner): the fraction stays.
    [Puzzle("two-squares-in-a-triangle")]
    public static void TwoSquaresInATriangle(Lgf d)
    {
        d.Texture = "FeltTip";
        var a = d.Pt(0, 0);
        var c = d.Pt(5, 0);
        var b = d.Pt(1.8, 2.4);
        string orange = d.Fill("Orange", "#FFE87C02", Lgf.Ink, 5);
        d.Polygon(orange, b, d.Pt(1.2, 1.6), d.Pt(2, 1), d.Pt(2.6, 1.8));
        var l = d.Slider(V.At(1.03, 0), V.At(1.97, 0), (1.33 - 1.03) / 0.94, (1.33, 0), "L");
        d.Polygon(orange, l, d.Pt(new V("L.X + 1", "0")), d.Pt(new V("L.X + 1", "1")), d.Pt(new V("L.X", "1")));
        d.Inked(d.Outline("Frame", Lgf.Ink, 5), a, c, b);
        d.Scene(0, 2.4, 5, 0);
    }

    // A square, a quarter circle of area π in its top left corner, and a semicircle of area π whose
    // diameter touches the quarter circle at its own middle and whose arc touches the right and bottom
    // sides. The square's area? Nothing to drag: the two areas fix it (the semicircle's center is on the
    // diagonal, 2 from the corner).
    [Puzzle("a-semi-circle-and-quarter-circle-in-a-square")]
    public static void ASemiCircleAndQuarterCircleInASquare(Lgf d)
    {
        double r2 = Math.Sqrt(2), s = 2 * r2;
        var tl = d.Pt(0, s);
        var c = d.Pt(r2, r2);
        string green = "#FF22B87A", navy = "#FF0614E1";
        d.Arc("CircleSector", d.Fill("Green", "#FFA2EFCD", green, 3), tl, d.Pt(0, s - 2), d.Pt(2, s));
        d.Arc("CircleSegment", d.Fill("Blue", "#FF08DEFD", navy, 3.5), c, d.Pt(r2 - 1, r2 - 1), d.Pt(r2 + 1, r2 + 1));
        d.Inked(d.Outline("Frame", Lgf.Ink, 5), d.Pt(0, 0), d.Pt(s, 0), d.Pt(s, s), tl);
        d.Label("π", V.At(0.65, s - 0.65), d.ShapeText("GreenMark", green, 42), 42);
        d.Label("π", V.At(r2 + 0.35, r2 - 0.35), d.ShapeText("NavyMark", navy, 42), 42);
        // a little paper under it: the page's buttons sit at the bottom right corner
        d.Scene(0, s, s, -0.3);
    }

    // A regular hexagon standing on a corner and a triangle (equilateral, as she drew it) with its top
    // at the hexagon's top corner and the middle of its base at the bottom one, so its sides run through
    // the hexagon's lower corners. What fraction of the total area is shaded (the parts in only one of
    // them)? Nothing to drag: the hexagon fixes the triangle.
    [Puzzle("overlapping-triangle-and-hexagon")]
    public static void OverlappingTriangleAndHexagon(Lgf d)
    {
        const double R = 4;
        var hexagon = Enumerable.Range(0, 6).Select(i => d.Pt(R * Math.Cos((90 + 60 * i) * Math.PI / 180), R * Math.Sin((90 + 60 * i) * Math.PI / 180))).ToArray();
        var (top, upperLeft, lowerLeft, bottom, lowerRight, upperRight) = (hexagon[0], hexagon[1], hexagon[2], hexagon[3], hexagon[4], hexagon[5]);
        double w = 2 * R / Math.Sqrt(3);
        var a = d.Pt(-w, -R);
        var c = d.Pt(w, -R);
        string pink = d.Fill("Pink", "#FFFDA8F8", null);
        d.Polygon(pink, top, upperLeft, lowerLeft);
        d.Polygon(pink, top, lowerRight, upperRight);
        d.Polygon(pink, a, bottom, lowerLeft);
        d.Polygon(pink, c, lowerRight, bottom);
        d.Polygon(d.Outline("Navy", "#FF0A0E50", 4, dark: "#FFB4B8F0"), hexagon);
        d.Inked(d.Outline("Frame", Lgf.Ink, 5), a, c, top);
        d.Scene(-w, R, w, -R);
    }

    // A rectangle and a triangle of the same area on its bottom side, so the triangle is twice as tall
    // and the rectangle's top crosses its sides at their middles. What fraction of the total is shaded
    // (the parts in only one of them)? Dragging the triangle's top anywhere changes its shape and the
    // rectangle's height with it: the fraction stays.
    [Puzzle("overlapping-triangle-and-rectangle")]
    public static void OverlappingTriangleAndRectangle(Lgf d)
    {
        const double W = 5.6;
        var a = d.Pt(0, 0);
        var c = d.Pt(W, 0);
        string range = d.Polygon(null, d.Pt(0.8, 2.6), d.Pt(5, 2.6), d.Pt(5, 5.6), d.Pt(0.8, 5.6));
        var t = d.HandleIn(range, 3.64, 4, "T");
        var left = Mid(d, a, t);
        var right = Mid(d, c, t);
        var topLeft = d.Pt(new V("0", "T.Y / 2"));
        var topRight = d.Pt(new V(Lgf.N(W), "T.Y / 2"));
        string purple = d.Fill("Purple", "#FFCDB2F2", null);
        d.Polygon(purple, left, right, t);
        d.Polygon(purple, a, left, topLeft);
        d.Polygon(purple, c, topRight, right);
        d.Polygon(d.Outline("Orange", "#FFEC840A", 3.5), a, c, topRight, topLeft);
        d.Inked(d.Outline("Magenta", "#FFD8017B", 4), a, c, t);
        d.Scene(0, 4, W, 0);
    }

    /// <summary>Where the line from a point through a point on a circle around the center meets the circle again</summary>
    private static V SecondCrossing(Lgf d, V from, V through, V center, string name)
    {
        string f = from.X.Replace(".X", ""), t = through.X.Replace(".X", ""), c = center.X.Replace(".X", "");
        string length = d.Num($"distance({f}, {t})");
        var direction = new V($"({t}.X - {f}.X) / {length}", $"({t}.Y - {f}.Y) / {length}");
        string step = d.Num($"-2 * (({t}.X - {c}.X) * {direction.X} + ({t}.Y - {c}.Y) * {direction.Y})");
        return d.Pt(through + step * direction, name);
    }
}
