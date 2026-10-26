#nullable disable

static partial class Puzzles
{
    /// <summary>
    /// A filled angle mark of the given radius in drawing units (a sector, counterclockwise from the
    /// side toward a to the side toward b): it keeps her proportions at every size, where Angle's
    /// radius is in pixels
    /// </summary>
    static string WedgeInUnits(Lgf d, string style, V vertex, V a, V b, double radius)
    {
        string toA = d.Num($"sqrt(({a.X} - {vertex.X}) ^ 2 + ({a.Y} - {vertex.Y}) ^ 2)");
        string toB = d.Num($"sqrt(({b.X} - {vertex.X}) ^ 2 + ({b.Y} - {vertex.Y}) ^ 2)");
        var start = d.Pt(vertex + Lgf.N(radius) * new V($"({a.X} - {vertex.X}) / {toA}", $"({a.Y} - {vertex.Y}) / {toA}"));
        var end = d.Pt(vertex + Lgf.N(radius) * new V($"({b.X} - {vertex.X}) / {toB}", $"({b.Y} - {vertex.Y}) / {toB}"));
        return d.Arc("CircleSector", style, vertex, start, end);
    }

    // A rectangle of width 3 whose top corners are on a circle, its bottom side touching it; lines from
    // a point on the circle to the top corners, to where the sides cross the circle and to where the
    // bottom touches it make four equal angles. The circle's area? Equal angles there are equal arcs,
    // which puts the crossings 60° apart: the rectangle is fixed. Dragging the point along the top arc
    // moves the lines, the four angles stay equal.
    [Puzzle("rectangle-overlapping-a-circle")]
    public static void RectangleOverlappingACircle(Lgf d)
    {
        d.Texture = "FeltTip";
        double r = Math.Sqrt(3), w = 1.5, h = r / 2;
        string navy = "#FF0A0A44";
        var o = d.Pt(0, 0, "O");
        var rim = d.Pt(r, 0);
        var a = d.Pt(-w, h, "A");
        var b = d.Pt(w, h, "B");
        var l = d.Pt(-w, -h, "L");
        var q = d.Pt(w, -h, "R");
        var t = d.Pt(0, -r, "T");
        var bl = d.Pt(-w, -r);
        var br = d.Pt(w, -r);
        d.Polygon(d.Outline("Rectangle", navy, 3.5), a, b, br, bl);
        d.Circle(d.Line("Circle", "#FFCA0B5B", 3), o, rim);
        var p = d.ArcSlider(o, r, 40, 140, 105, (0, 0), "P");
        string yellow = d.Line("Yellow", "#FFFCEA04", 4);
        foreach (var x in new[] { a, l, t, q, b })
        {
            d.Segment(yellow, p, x);
        }

        string wedge = d.Wedge("OrangeWedge", "#FFE8700A");
        d.Angle(wedge, p, a, l, 36);
        d.Angle(wedge, p, l, t, 36);
        d.Angle(wedge, p, t, q, 36);
        d.Angle(wedge, p, q, b, 36);
        d.Dimension(d.Line("Arrow", navy, 3), d.Text("NavyMark", navy, 30), br, bl, 0.3, "3", 0.2);
        d.Scene(-r, r, r, -r - 0.45);
    }

    // Four semicircles on a vertical line: the largest to the left, the smallest to the right around
    // the same center, and two equal ones between their ends, the left one cut out of the largest, the
    // right one around the smallest. The chord of the largest at the height of the smallest's top end
    // is 2; the shaded area? It is half the ring between the two concentric circles, half of π times
    // the chord squared over four. Dragging the chord's left end up and down resizes every semicircle
    // (the chord stays 2): the shaded area stays.
    [Puzzle("swirling-semi-circles")]
    public static void SwirlingSemiCircles(Lgf d)
    {
        var c = d.Pt(0, 0, "C");
        var h = d.Slider(V.At(-2, 0.3), V.At(-2, 1.5), (0.52 - 0.3) / 1.2, (-2, 0.52), "H");
        string s = "H.Y";
        string big = d.Num($"sqrt(4 + {s} ^ 2)", "R");
        var top = d.Pt(new V("0", big), "Top");
        var bottom = d.Pt(new V("0", $"-{big}"), "Bottom");
        var up = d.Pt(new V("0", s), "Up");
        var down = d.Pt(new V("0", $"-{s}"), "Down");
        var upper = d.Pt(new V("0", $"({big} - {s}) / 2"), "Cu");
        var lower = d.Pt(new V("0", $"({s} - {big}) / 2"), "Cl");
        string largest = d.Arc("CircleSegment", null, c, top, bottom, visible: false);
        string leftMiddle = d.Arc("CircleSegment", null, upper, top, down, visible: false);
        string rightMiddle = d.Arc("CircleSegment", null, lower, bottom, up, visible: false);
        string smallest = d.Arc("CircleSegment", null, c, down, up, visible: false);
        d.Region(d.Fill("Blue", "#FF05A6FC", null), "Difference", d.Region(null, "Union", largest, rightMiddle), leftMiddle, smallest);
        string blue = d.Line("BlueLine", "#FF0448F0", 3);
        d.Arc("CircleArc", blue, c, top, bottom);
        d.Arc("CircleArc", blue, upper, top, down);
        d.Arc("CircleArc", blue, lower, bottom, up);
        d.Arc("CircleArc", blue, c, down, up);
        d.Segment(blue, top, bottom);
        string pink = "#FFF0206D";
        // her arrow lies on the chord itself, the 2 over it
        string arrow = d.Line("PinkArrow", pink, 3);
        var middle = d.Pt(new V("-1", s));
        d.Vector(arrow, middle, h);
        d.Vector(arrow, middle, up);
        d.Label("2", new V("-0.85", $"{s} + 0.25"), d.Text("PinkMark", pink, 30));
        d.Scene(-2.2, 2.2, 1.35, -2.2);
    }

    // A regular hexagon of side 2 standing on the quarter circle's bottom radius, its right corner and
    // its top right corner on the arc. The quarter circle's area? Nothing to drag: the two corners on
    // the arc put the hexagon one side from the center, and the radius is √28.
    [Puzzle("hexagon-in-a-quarter-circle")]
    public static void HexagonInAQuarterCircle(Lgf d)
    {
        const double a = 2;
        double r = Math.Sqrt(28);
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(r, 0);
        var north = d.Pt(0, r);
        d.Arc("CircleSector", d.Fill("Purple", "#FF8A48DB", "#FF6A20C8"), o, east, north);
        var hexagon = HexagonUp(d, V.At(a, 0), Lgf.N(a));
        // over the sector, which polygons are otherwise under
        d.Inked(d.Fill("Lavender", "#FFD8BDF4", Lgf.Ink, 4.5), hexagon);
        // on the purple: ink in both themes
        d.Dimension(d.Line("Arrow", Lgf.Ink, 3, dark: Lgf.Ink), d.ShapeText("ShapeMark", Lgf.Ink, 30), hexagon[5], hexagon[4], 0.4, "2", 0.3);
        d.Scene(0, r, r, 0);
    }

    // A regular hexagon in a circle; six equal semicircles on its sides, each touching the next, and a
    // circle in the middle touching them all. What fraction of the outer circle is shaded? Nothing to
    // drag: touching makes the semicircles' diameter the hexagon's height over two (√3/2 of a side), and
    // the middle circle as big as they are.
    [Puzzle("six-semi-circles-inside-a-hexagon")]
    public static void SixSemiCirclesInsideAHexagon(Lgf d)
    {
        d.Texture = "FeltTip";
        const double s = 4;
        double rho = s * Math.Sqrt(3) / 4;
        var o = d.Pt(0, 0, "O");
        var corners = Enumerable.Range(0, 6).Select(i => d.Pt(s * Math.Cos(i * Math.PI / 3), s * Math.Sin(i * Math.PI / 3))).ToArray();
        d.Circle(d.Line("Circle", "#FFF26010", 2.5), o, corners[0]);
        string orange = d.Fill("Orange", "#FFFDA503", "#FFD8141A", 3);
        for (int i = 0; i < 6; i++)
        {
            double x0 = s * Math.Cos(i * Math.PI / 3), y0 = s * Math.Sin(i * Math.PI / 3);
            double x1 = s * Math.Cos((i + 1) * Math.PI / 3), y1 = s * Math.Sin((i + 1) * Math.PI / 3);
            double mx = (x0 + x1) / 2, my = (y0 + y1) / 2, ux = (x1 - x0) / s, uy = (y1 - y0) / s;
            d.Arc("CircleSegment", orange, d.Pt(mx, my), d.Pt(mx + rho * ux, my + rho * uy), d.Pt(mx - rho * ux, my - rho * uy));
        }

        d.Circle(orange, o, d.Pt(rho, 0));
        // over the semicircles, whose diameters lie on its sides
        d.Inked(d.Outline("Hexagon", "#FFD8141A", 3), corners);
        d.Scene(-s, s, s, -s);
    }

    // Four quarter circles: in the big one a green one on the bottom radius from the corner, a blue one
    // hanging from a point on the arc straight above the green one's center down to its top, and a
    // yellow one from the left radius at that height out to the blue one, touching the green one. What
    // fraction is shaded? Nothing to drag: the touching makes the corner on the arc a 3-4-5 point.
    [Puzzle("four-quarter-circles")]
    public static void FourQuarterCircles(Lgf d)
    {
        const double R = 5, g = 3, h = 4, b = h - g, y = R - g;
        var o = d.Pt(0, 0, "O");
        var green = d.Pt(g, 0, "G");
        var blue = d.Pt(g, h, "B");
        var yellow = d.Pt(0, h, "Y");
        d.Arc("CircleSector", d.Fill("Green", "#FF36DFBC", "#FF05A8AC"), green, d.Pt(g, g), o);
        d.Arc("CircleSector", d.Fill("Blue", "#FF0564D6", "#FF1040C0"), blue, d.Pt(g - b, h), d.Pt(g, h - b));
        d.Arc("CircleSector", d.Fill("Yellow", "#FFFFFE9F", null), yellow, d.Pt(0, h - y), d.Pt(y, h));
        d.Arc("CircleSector", d.Outline("Red", "#FFD8607E", 3), o, d.Pt(R, 0), d.Pt(0, R));
        d.Scene(0, R, R, 0);
    }

    // Two triangles on one line and a circle: the circle is the blue triangle's incircle, and the
    // teal triangle's sides run from a point on the circle (its top) through the circle's touching
    // points on the blue triangle's slanted sides. The sum of the three marked angles (at the teal
    // top and between the teal and the blue sides at those touching points)? Dragging the blue
    // triangle's top reshapes it around the circle, dragging the teal top moves it along the circle:
    // the angles change, their sum stays 180°.
    [Puzzle("two-overlapping-triangles-and-a-circle")]
    public static void TwoOverlappingTrianglesAndACircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double r = 2;
        var o = d.Pt(0, 0, "O");
        var bottom = d.Pt(0, -r, "T0");
        var baseLeft = d.Pt(-10, -r);
        var baseRight = d.Pt(10, -r);
        var c = d.HandleIn(d.Circle(null, d.Pt(0.91, 5.28), d.Pt(0.91 + 0.5, 5.28)), 0.91, 5.28, "C");
        string d2 = d.Num("C.X ^ 2 + C.Y ^ 2", "dd");
        string k1 = d.Num($"{Lgf.N(r * r)} / {d2}", "k1");
        string k2 = d.Num($"{Lgf.N(r)} * sqrt({d2} - {Lgf.N(r * r)}) / {d2}", "k2");
        var t1 = d.Pt(new V($"{k1} * C.X - {k2} * C.Y", $"{k1} * C.Y + {k2} * C.X"), "T1");
        var t2 = d.Pt(new V($"{k1} * C.X + {k2} * C.Y", $"{k1} * C.Y - {k2} * C.X"), "T2");
        var p = d.ArcSlider(o, r, 92, 126, 111, (0, 0), "P");
        var a = d.Crossing(c, t1, baseLeft, baseRight, "A");
        var b = d.Crossing(c, t2, baseLeft, baseRight, "B");
        var e = d.Crossing(p, t1, baseLeft, baseRight, "E");
        var f = d.Crossing(p, t2, baseLeft, baseRight, "F");
        string wedge = d.Wedge("RedWedge", "#FFFC3569");
        WedgeInUnits(d, wedge, t1, e, a, 0.55);
        WedgeInUnits(d, wedge, p, t1, t2, 0.55);
        WedgeInUnits(d, wedge, t2, b, f, 0.55);
        d.Circle(d.Line("Green", "#FF3EDD72", 3.5), o, bottom);
        string teal = d.Line("Teal", "#FF048CBA", 4);
        d.Segment(teal, e, p);
        d.Segment(teal, p, f);
        d.Segment(teal, e, f);
        string blue = d.Line("Blue", "#FF0695F6", 4);
        d.Segment(blue, a, c);
        d.Segment(blue, c, b);
        d.Segment(blue, a, b);
        d.Scene(-5.6, 5.3, 5.9, -r);
    }

    // A rectangle, its diagonal, and two congruent squares standing on the diagonal on either side of
    // it, each with its two far corners on two sides of the rectangle; the three red parts of the
    // diagonal are equal. What fraction is shaded? Nothing to drag: the conditions make the rectangle
    // two by one, the red parts a seventh of the diagonal and the squares' sides two sevenths.
    [Puzzle("two-squares-diagonally-inside-a-rectangle")]
    public static void TwoSquaresDiagonallyInsideARectangle(Lgf d)
    {
        const double W = 4, H = 2;
        V[] left = [d.Pt(4.0 / 7, 2.0 / 7), d.Pt(12.0 / 7, 6.0 / 7), d.Pt(8.0 / 7, H), d.Pt(0, 10.0 / 7)];
        V[] right = [d.Pt(W - 4.0 / 7, H - 2.0 / 7), d.Pt(W - 12.0 / 7, H - 6.0 / 7), d.Pt(W - 8.0 / 7, 0), d.Pt(W, H - 10.0 / 7)];
        string blue = d.Fill("Blue", "#FF0597F2", "#FF0204A0", 4);
        d.Polygon(blue, left);
        d.Polygon(blue, right);
        string red = d.Line("Red", "#FFE8051E", 3);
        d.Segment(red, d.Pt(0, 0), left[0], "TwoTicks");
        d.Segment(red, left[1], right[1], "TwoTicks");
        d.Segment(red, right[0], d.Pt(W, H), "TwoTicks");
        foreach (var corner in new[] { left[0], left[1], right[0], right[1] })
        {
            d.DotAt(corner);
        }

        d.Box(d.Outline("Frame", Lgf.Ink, 4.5), V.At(0, 0), V.At(W, H));
        d.Scene(0, H, W, 0);
    }

    // A smiling face: a big semicircle hanging from its diameter, cut by a chord parallel to it; a
    // second semicircle on the chord, and two circles between the diameter and the chord. The yellow
    // (the second semicircle outside the first) is 16; the black (the first semicircle above the
    // chord, outside the circles)? Dragging the chord's end along the big arc moves the chord: the
    // circles are as wide as the band, and the black stays the yellow.
    [Puzzle("two-circles-and-two-semi-circles")]
    public static void TwoCirclesAndTwoSemiCircles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 5;
        var o = d.Pt(0, 0, "O");
        var west = d.Pt(-R, 0);
        var east = d.Pt(R, 0);
        var e = d.ArcSlider(o, R, 322, 348, 333.1, (0, 0), "E");
        string h = "(-E.Y)";
        var w = d.Pt(new V("-E.X", "E.Y"), "W");
        var c = d.Pt(new V("0", "E.Y"), "C");
        string offset = d.Num($"0.4 * sqrt({Lgf.N(R * R)} - {Lgf.N(R)} * {h}) + 0.3 * {h}", "e");
        var leftEye = d.Pt(new V($"-{offset}", "E.Y / 2"), "L");
        var rightEye = d.Pt(new V(offset, "E.Y / 2"), "Rt");
        var leftRim = d.Pt(new V($"-{offset}", "0"));
        var rightRim = d.Pt(new V(offset, "0"));
        string first = d.Arc("CircleSegment", null, o, west, east, visible: false);
        string second = d.Arc("CircleSegment", null, c, w, e, visible: false);
        string below = d.Arc("CircleSegment", null, o, w, e, visible: false);
        string eye1 = d.Circle(null, leftEye, leftRim);
        string eye2 = d.Circle(null, rightEye, rightRim);
        d.Region(d.Fill("Black", "#FF020114", null), "Difference", first, below, eye1, eye2);
        d.Region(d.Fill("Yellow", "#FFFFFF01", null), "Difference", second, below);
        string red = d.PaperFill("Eye", "#FFCE0411", 3.5);
        d.Circle(red, leftEye, leftRim);
        d.Circle(red, rightEye, rightRim);
        // around the yellow, which stays bright in the dark theme: a blue between light and deep
        string blue = d.Outline("Blue", "#FF0510C8", 3.5, dark: "#FF6B7BFF");
        d.Arc("CircleSegment", blue, o, west, east);
        d.Arc("CircleSegment", blue, c, w, e);
        d.Scene(-R, 0, R, -6.9);
    }

    // Two squares on a line: a tilted one standing on its corner, and an upright one as high as the
    // tilted one's right corner, which is its top left corner. The pink is both squares above the
    // line from the tilted one's left corner to the upright one's bottom right corner; the distance
    // from the tilted one's top corner to the upright one's top right corner is 6. The pink area? It
    // is half that distance squared. Dragging the upright square's foot along the line tilts the
    // other one (the 6 stays): the pink area stays.
    [Puzzle("two-tilted-squares")]
    public static void TwoTiltedSquares(Lgf d)
    {
        var a = d.Pt(0, 0, "A");
        var foot = d.Slider(V.At(0.4, 0), V.At(2.6, 0), (1.4 - 0.4) / 2.2, (1.4, 0), "K");
        string b = d.Num("sqrt(36 - K.X ^ 2) / 2", "b");
        var corner = d.Pt(new V("K.X", b), "B");
        var top = d.Pt(new V($"K.X - {b}", $"K.X + {b}"), "C");
        var left = d.Pt(new V($"-{b}", "K.X"), "D");
        var far = d.Pt(new V($"K.X + {b}", b), "F");
        var end = d.Pt(new V($"K.X + {b}", "0"), "E");
        d.Polygon(d.Fill("Pink", "#FFFBB2FC", null), left, top, corner, far, end);
        string frame = d.Outline("Frame", Lgf.Ink, 4);
        d.Polygon(frame, a, corner, top, left);
        d.Polygon(frame, foot, end, far, corner);
        string ink = d.Line("Ink", Lgf.Ink, 4);
        d.Segment(ink, d.Pt(-1.7, 0), d.Pt(new V($"K.X + {b} + 0.4", "0")));
        string navy = "#FF1E2470";
        d.Segment(d.Line("Navy", navy, 3.5), left, end);
        d.Dimension(d.Line("Arrow", navy, 3), d.Text("NavyMark", navy, 32), top, far, 0, "6", 0.45);
        d.Scene(-2.95, 4.35, 4.55, 0);
    }

    // Two congruent regular hexagons overlapping side by side; in the overlap a circle touching its
    // four slanted sides, and above and below it a semicircle on the circle, touching the overlap's
    // top (bottom) and slanted sides. What fraction of the total area is shaded? Nothing to drag: the
    // semicircles fix the overlap, its middle width four thirds of a side.
    [Puzzle("a-circle-and-two-semi-circles-in-two-hexagons")]
    public static void ACircleAndTwoSemiCirclesInTwoHexagons(Lgf d)
    {
        d.Texture = "FeltTip";
        const double s = 3, m = 2 * s / 3;
        double h = S3 * s, r = m * S3, rho = S3 * m / 2;
        V[] Hexagon(double cx) => [d.Pt(cx - s, 0), d.Pt(cx - s / 2, -h), d.Pt(cx + s / 2, -h), d.Pt(cx + s, 0), d.Pt(cx + s / 2, h), d.Pt(cx - s / 2, h)];
        var left = Hexagon(m - s);
        var right = Hexagon(s - m);
        string leftShape = d.Polygon(null, left);
        string rightShape = d.Polygon(null, right);
        string yellow = d.Fill("Yellow", "#FFFFFF01", null);
        d.Region(yellow, "Difference", leftShape, rightShape);
        d.Region(yellow, "Difference", rightShape, leftShape);
        string frame = d.Outline("Frame", Lgf.Ink, 4);
        d.Polygon(frame, left);
        d.Polygon(frame, right);
        string red = d.Outline("Red", "#FFA70221", 3);
        var o = d.Pt(0, 0, "O");
        d.Circle(red, o, d.Pt(r, 0));
        d.Arc("CircleSegment", red, d.Pt(0, r), d.Pt(rho, r), d.Pt(-rho, r));
        d.Arc("CircleSegment", red, d.Pt(0, -r), d.Pt(-rho, -r), d.Pt(rho, -r));
        d.Scene(m - 2 * s, h, 2 * s - m, -h);
    }

    // Four squares of area 1 in a tilted zig-zag over a rectangle: the purple one's corners on the
    // rectangle's left and bottom sides, the blue one's top corner on its top side, the green one's
    // corner at its top right corner. The rectangle's area? Nothing to drag: the corners on the sides
    // tilt the squares by the angle whose tangent is 1/2.
    [Puzzle("four-squares-overlapping-a-rectangle")]
    public static void FourSquaresOverlappingARectangle(Lgf d)
    {
        const double k = 2;
        double c = 2 / Math.Sqrt(5), s = 1 / Math.Sqrt(5);
        (double X, double Y) u = (k * c, -k * s), v = (k * s, k * c);
        V At(int i, int j) => d.Pt(i * u.X + j * v.X, i * u.Y + j * v.Y);
        V[] Square(int i, int j) => [At(i, j), At(i + 1, j), At(i + 1, j + 1), At(i, j + 1)];
        string ink = Lgf.Ink;
        d.Polygon(d.Fill("Purple", "#FF8858D6", ink, 4), Square(0, 0));
        d.Polygon(d.Fill("Blue", "#FF6AB0F0", ink, 4), Square(0, 1));
        d.Polygon(d.Fill("Cyan", "#FF4AD8F7", ink, 4), Square(1, 1));
        d.Polygon(d.Fill("Green", "#FF3CEA84", ink, 4), Square(1, 2));
        double right = 3 * v.X + 2 * u.X, top = 3 * v.Y + 2 * u.Y, bottom = u.Y;
        d.Box(d.Outline("Rectangle", "#FF0A1AA8", 4), V.At(0, bottom), V.At(right, top));
        d.Scene(0, 3 * v.Y + u.Y, right, bottom);
    }

    // A right triangle with a vertical leg of 6 and two equal semicircles on its hypotenuse, each
    // ending at the other's center (the dots): the lower one inside, touching the bottom leg, the upper
    // one outside, ending at the top corner. The total shaded area? Nothing to drag: as she drew it the
    // hypotenuse is at 30° and the four steps from the far corner (to the lower semicircle's end, the
    // two centers and the top corner) are equal, each the radius, 3. The puzzle only says the
    // semicircles touch the leg and end at the corner, which leaves the angle free and the area with
    // it, so the drawing keeps hers.
    [Puzzle("two-semi-circles-and-a-triangle")]
    public static void TwoSemiCirclesAndATriangle(Lgf d)
    {
        const double r = 3;
        double cx = 4 * r * S3;
        var a = d.Pt(0, 0, "A");
        var q = d.Pt(r * S3, r / 2, "Q");
        var p1 = d.Pt(2 * r * S3, r, "P1");
        var p2 = d.Pt(3 * r * S3, 1.5 * r, "P2");
        var b = d.Pt(cx, 2 * r, "B");
        var c = d.Pt(cx, 0, "C");
        string lower = d.Arc("CircleSegment", null, p1, q, p2, visible: false);
        string upper = d.Arc("CircleSegment", null, p2, b, p1, visible: false);
        d.Region(d.Fill("Teal", "#FF74E8CA", null), "Union", lower, upper);
        string teal = "#FF0B9796";
        string arc = d.Line("TealLine", teal, 3);
        d.Arc("CircleArc", arc, p1, q, p2);
        d.Arc("CircleArc", arc, p2, b, p1);
        // her dots are big: discs in drawing units, under the hypotenuse inked over them
        string dot = d.Fill("Dot", Lgf.Ink, Lgf.Ink, 1);
        d.Circle(dot, p1, d.Pt(2 * r * S3 + 0.2, r));
        d.Circle(dot, p2, d.Pt(3 * r * S3 + 0.2, 1.5 * r));
        d.Inked(d.Outline("Triangle", Lgf.Ink, 4.5), a, c, b);
        d.Dimension(d.Line("TealArrow", teal, 3), d.Text("TealMark", teal, 32), b, c, 1.0, "6", 0.75);
        d.Scene(0, 1.5 * r + r, cx + 1.25, 0);
    }

    // A regular hexagon and a square with a corner at the hexagon's bottom left corner and its top side
    // through the hexagon's top right corner. The angle at the square's top left corner between its
    // top side and the line to the hexagon's bottom right corner? That corner is on the circle over
    // the diagonal from the bottom left to the top right corner, as is the hexagon's bottom right
    // corner: the angle is the hexagon's, 60°. Dragging the square's top left corner along that circle
    // turns and resizes the square: the angle stays.
    [Puzzle("a-square-and-a-hexagon")]
    public static void ASquareAndAHexagon(Lgf d)
    {
        d.Texture = "FeltTip";
        const double k = 3;
        var hexagon = HexagonUp(d, V.At(0, 0), Lgf.N(k));
        var center = d.Pt(k / 2, k * S3, "M");
        var t = d.ArcSlider(center, k, 116, 162, 140.9, (k / 2, k * S3), "T");
        string q = d.Num("sqrt(T.X ^ 2 + T.Y ^ 2)", "q");
        var u = new V("T.Y", "-T.X");
        var right = d.Pt(u, "U");
        var far = d.Pt(V.Of("T") + u, "W");
        d.Polygon(d.Outline("Hexagon", "#FF0A06A8", 4), hexagon);
        d.Polygon(d.Outline("Square", Lgf.Ink, 4.5), hexagon[0], right, far, t);
        string purple = d.Line("Purple", "#FFA884E6", 3);
        d.Segment(purple, t, hexagon[1]);
        d.ArcMark(d.Line("PurpleMark", "#FF8A50D8", 3.5), t, hexagon[1], far, 0.65);
        string mark = d.Text("PurpleText", "#FF7A30C8", 34);
        string ux = d.Num($"T.Y / {q}", "ux");
        string uy = d.Num($"-T.X / {q}", "uy");
        // on the bisector of the 60° angle, which is the side turned 30° toward the purple line
        d.Label("?", V.Of("T") + new V($"1.15 * ({s3} * {ux} + 0.5 * {uy})", $"1.15 * ({s3} * {uy} - 0.5 * {ux})"), mark);
        d.Scene(-1.6, 5.35, 5.5, -0.1);
    }
}
