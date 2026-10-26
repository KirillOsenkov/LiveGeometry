#nullable disable

static partial class Puzzles
{
    // A big semicircle, a tilted rectangle whose top side runs from a point 8 left of the top of the
    // semicircle (the orange line) through a point on the arc (the dot) to a corner on the arc, and a
    // small semicircle on that side, centered on the dot, from the rectangle's corner, touching its
    // bottom side. The rectangle's area? Dragging the far corner along the arc tilts the rectangle:
    // its length times the small radius is the power of the corner, 8 squared, and stays 64.
    [Puzzle("two-semi-circles-and-a-rectangle")]
    public static void TwoSemiCirclesAndARectangle(Lgf d)
    {
        const double R = 8.6;
        var o = d.Pt(0, 0, "O");
        var e = d.Pt(R, 0, "E");
        var w = d.Pt(-R, 0, "W");
        var t = d.Pt(0, R, "T");
        var p = d.Pt(-8, R, "P");
        var q = d.ArcSlider(o, R, 24, 60, 33, (0, 0), "Q");
        string k = d.Num($"64 / ((Q.X + 8) ^ 2 + (Q.Y - {Lgf.N(R)}) ^ 2)", "k");
        var c = d.Pt(p + k * (q - p), "C");
        string r = d.Num("distance(P, C)", "r");
        string length = d.Num("distance(P, Q)", "len");
        // the unit normal of the top side, toward the inside of the rectangle (down)
        var n = new V($"(Q.Y - P.Y) / {length}", $"(P.X - Q.X) / {length}");
        var bottomRight = d.Pt(q + r * n, "Br");
        var bottomLeft = d.Pt(p + r * n, "Bl");
        var end = d.Pt(c + (c - p), "D");
        d.Polygon(d.Fill("Pink", "#FFFCE8F7", "#FFD0108A", 3), p, q, bottomRight, bottomLeft);
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        d.Arc("CircleSegment", frame, c, p, end);
        d.Arc("CircleSegment", frame, o, e, w);
        d.DotAt(o, "Od", 13);
        d.DotAt(c, "Cd", 13);
        string orange = "#FFEF5A10";
        d.Dimension(d.Line("Arrow", orange, 3), d.Text("OrangeMark", orange, 30), p, t, 0, "8", 0.9);
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(-9.4, R + 0.5, R, -0.5);
    }

    // A regular hexagon of side 4 cut by a horizontal chord so that the part above (green) is a third
    // of it, and a semicircle on the chord. Its area? Nothing to drag: the third puts the chord
    // 6 - 2√3 below the top side, and it is 4√3 long.
    [Puzzle("semi-circle-on-a-hexagon")]
    public static void SemiCircleOnAHexagon(Lgf d)
    {
        d.Texture = "FeltTip";
        double sqrt3 = Math.Sqrt(3);
        var hexagon = HexagonFromLeft(d, V.At(-4, 0), "4");
        double h = 4 * sqrt3 - 6, half = 2 * sqrt3;
        var left = d.Pt(-half, h);
        var right = d.Pt(half, h, "Cr");
        var center = d.Pt(0, h, "C");
        d.Polygon(d.Fill("Green", "#FF02E8C0", null), left, right, hexagon[2], hexagon[1]);
        d.Polygon(d.Outline("Frame", Lgf.Ink, 3), hexagon);
        d.Arc("CircleSegment", d.Outline("Blue", "#FF1020D0", 3), center, right, left);
        d.Dimension(d.Line("Arrow"), d.Text(), hexagon[4], hexagon[5], 0.5, "4", 0.55);
        d.Scene(-4, h + half, 4, -2 * sqrt3 - 0.9);
    }

    // A square whose diagonal is cut into three lengths of 4; a red line from the bottom left corner
    // through the second cut to the right side, and one back from there through the first cut to the
    // left side. The total shaded area? Nothing to drag: the diagonal fixes the square, and the lines
    // meet the sides halfway up and three quarters up.
    [Puzzle("divided-square")]
    public static void DividedSquare(Lgf d)
    {
        d.Texture = "FeltTip";
        double s = 6 * Math.Sqrt(2);
        var bl = d.Pt(0, 0);
        var br = d.Pt(s, 0);
        var tl = d.Pt(0, s);
        var p1 = d.Pt(s / 3, 2 * s / 3);
        var p2 = d.Pt(2 * s / 3, s / 3);
        var e = d.Pt(s, s / 2);
        var f = d.Pt(0, 3 * s / 4);
        string yellow = d.Fill("Yellow", "#FFFFF602", null);
        d.Polygon(yellow, tl, f, p1);
        d.Polygon(yellow, p1, e, p2);
        d.Polygon(yellow, bl, p2, br);
        string orange = "#FFF06A05";
        // across the yellow, which stays bright in the dark theme: deep there
        d.Segment(d.Line("Diagonal", orange, 3.5, dark: "#FFB04A00"), tl, br);
        string red = d.Line("Red", "#FFD02020", 3.5, dark: "#FFE8483A");
        d.Segment(red, f, e);
        d.Segment(red, bl, e);
        d.Box(d.Outline("Frame", Lgf.Ink, 5), V.At(0, 0), V.At(s, s));
        string mark = d.Text("OrangeMark", orange, 32);
        double k = 0.6 / Math.Sqrt(2);
        d.Label("4", V.At(s / 6 + k, 5 * s / 6 + k), mark);
        d.Label("4", V.At(s / 2 - k, s / 2 - k), mark);
        d.Label("4", V.At(5 * s / 6 + k, s / 6 + k), mark);
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(0, s, s, -0.5);
    }

    // Three equilateral triangles of side 30 side by side (up, down, up), the first cut into three
    // equal areas (blue), the second into four (orange), the third into five (pink), each cut
    // starting where the last one ended. The missing length, from the first cut's foot to the third's
    // first foot? Nothing to drag: the areas fix every cut (thirds, quarters and fifths of the sides).
    [Puzzle("three-adjacent-triangles")]
    public static void ThreeAdjacentTriangles(Lgf d)
    {
        double h = 15 * Math.Sqrt(3);
        var a = d.Pt(0, 0);
        var b = d.Pt(30, 0, "B");
        var r = d.Pt(60, 0);
        var t = d.Pt(15, h, "T");
        var t2 = d.Pt(45, h, "T2");
        var x = d.Pt(t + (1.0 / 3) * (b - t), "X");
        var y = d.Pt(15, 0, "Y");
        var z = d.Pt(t2 + 0.25 * (b - t2), "Z");
        var w = d.Pt(t2 + 0.625 * (b - t2), "W");
        var v1 = d.Pt(46, 0, "V1");
        var v2 = d.Pt(54, 0, "V2");
        string ink = d.Line("Ink", Lgf.Ink, 4);
        d.Polygon(d.Outline("Frame", Lgf.Ink, 4), a, r, t2, t);
        d.Segment(ink, t, b);
        d.Segment(ink, b, t2);
        string blue = d.Line("Blue", "#FF1E0CA0", 3.5);
        d.Segment(blue, a, x);
        d.Segment(blue, x, y);
        string orange = d.Line("Orange", "#FFF56A08", 3.5);
        d.Segment(orange, t, z);
        d.Segment(orange, z, x);
        d.Segment(orange, x, w);
        string pink = d.Line("Pink", "#FFFC48B8", 3.5);
        d.Segment(pink, w, v1);
        d.Segment(pink, v1, z);
        d.Segment(pink, z, v2);
        d.Segment(pink, v2, t2);
        string red = "#FFE83228";
        string arrow = d.Line("Arrow", red, 3);
        string mark = d.Text("RedMark", red, 34);
        d.Dimension(arrow, mark, t, t2, 4.4, "30", 3.8);
        d.Dimension(arrow, mark, v1, y, 3.6, "?", 2.6);
        d.Scene(0, h + 6, 60, -5.4);
    }

    // A square, a circle in its top left corner touching two sides, and an equilateral triangle of side
    // 15 in the bottom right corner whose left side touches the circle at the triangle's top. The
    // diameter? Nothing to drag: the tangency makes the radius 15 (the square's side 22.5 + 7.5√3).
    [Puzzle("triangle-and-circle-inside-square")]
    public static void TriangleAndCircleInsideSquare(Lgf d)
    {
        const double r = 15;
        double L = 22.5 + 7.5 * Math.Sqrt(3);
        var c = d.Pt(r, L - r, "C");
        var left = d.Pt(0, L - r);
        var right = d.Pt(2 * r, L - r);
        var a = d.Pt(L - 15, 0);
        var b = d.Pt(L, 0);
        var apex = d.Pt(L - 7.5, 7.5 * Math.Sqrt(3));
        d.Box(d.Outline("Square", "#FFA8EE18", 4), V.At(0, 0), V.At(L, L));
        d.Circle(d.Outline("Circle", "#FF12A078", 3.5), c, left);
        d.Polygon(d.Outline("Triangle", "#FF08E8FC", 4), a, b, apex);
        string blue = "#FF0812C8";
        string arrow = d.Line("Arrow", blue, 3.5);
        string mark = d.Text("BlueMark", blue, 30);
        d.Vector(arrow, c, left);
        d.Vector(arrow, c, right);
        d.Label("?", V.At(r, L - r + 2.6), mark);
        var under = d.Pt(L - 7.5, -2);
        d.Vector(arrow, under, d.Pt(L - 15, -2));
        d.Vector(arrow, under, d.Pt(L, -2));
        d.Label("15", V.At(L - 7.5, -4.6), mark);
        // paper below: the page's buttons sit over the bottom right corner, by the 15
        d.Scene(0, L, L, -9);
    }

    // A right triangle and two equal squares side by side on its hypotenuse, covering it; a line from
    // the right angle to the squares' shared top corner. The angle it makes with the base? Dragging
    // the top corner up and down changes the triangle: that shared corner is always on the circle
    // through the triangle's corners, halfway round the far arc, so the angle stays 45°.
    [Puzzle("two-squares-on-a-triangle")]
    public static void TwoSquaresOnATriangle(Lgf d)
    {
        var c = d.Pt(0, 0, "C");
        var a = d.Pt(-6, 0, "A");
        var b = d.Slider(V.At(0, 1.6), V.At(0, 4.6), (2.85 - 1.6) / 3, (0, 2.85), "B");
        var up = new V("-B.Y / 2", "3");
        var m = Mid(d, a, b);
        var a2 = d.Pt(a + up);
        var b2 = d.Pt(b + up);
        var p = d.Pt(m + up, "P");
        string purple = d.Outline("Purple", "#FF5A10C4", 3.5);
        d.Polygon(purple, a, b, b2, a2);
        d.Segment(d.Line("PurpleLine", "#FF5A10C4", 3.5), m, p);
        string cyan = "#FF10EEFC";
        string line = d.Line("Cyan", cyan, 3.5);
        d.Segment(line, c, p);
        d.Inked(d.Outline("Frame", Lgf.Ink, 4.5), a, c, b);
        d.ArcMark(line, c, p, a, 0.9);
        d.Label("?", V.At(-1.5, 0.75), d.Text("CyanMark", cyan, 30));
        d.Scene(-7.6, 6.1, 0.2, -0.3);
    }

    // A red rectangle and a blue square corner to corner on a line; a semicircle above the line around
    // their shared corner, as high as the rectangle, and one below from the first one's left end to the
    // rectangle's far corner, through the square's bottom right corner. Which shaded area is larger:
    // the orange (the rectangle outside the upper semicircle) or the pink (the square outside the
    // lower one)? Dragging the rectangle's top right corner changes both its sides (the square follows,
    // its side squared being their product): the two areas stay equal.
    [Puzzle("square-rectangle-and-two-semi-circles")]
    public static void SquareRectangleAndTwoSemiCircles(Lgf d)
    {
        var o = d.Pt(0, 0, "O");
        var h = d.HandleIn(d.Polygon(null, d.Pt(3.2, 2.5), d.Pt(4.6, 2.5), d.Pt(4.6, 3.3), d.Pt(3.9, 3.3)), 4.14, 3, "H");
        string s = d.Num("sqrt(H.X * H.Y)", "s");
        var east = d.Pt(new V("H.Y", "0"), "E");
        var north = d.Pt(new V("0", "H.Y"), "N");
        var west = d.Pt(new V("-H.Y", "0"), "W");
        var far = d.Pt(new V("H.X", "0"), "F");
        var c2 = d.Pt(new V("(H.X - H.Y) / 2", "0"), "C2");
        V[] rectangle = [o, far, h, north];
        V[] square = [d.Pt(new V($"-{s}", $"-{s}")), d.Pt(new V("0", $"-{s}")), o, d.Pt(new V($"-{s}", "0"))];
        string quarter = d.Arc("CircleSector", null, o, east, north, visible: false);
        d.Region(d.Fill("Orange", "#FFFEDB02", null), "Difference", d.Polygon(null, rectangle), quarter);
        d.Region(d.Fill("Pink", "#FFFDBFFE", null), "Difference", d.Polygon(null, square), d.Circle(null, c2, far));
        d.Polygon(d.Outline("Red", "#FFEE0606", 3.5), rectangle);
        d.Polygon(d.Outline("Blue", "#FF1A10B0", 3.5), square);
        d.Arc("CircleArc", d.Line("Lime", "#FFB4F01E", 3), o, east, west);
        d.Arc("CircleArc", d.Line("Green", "#FF07A131", 3), c2, west, far);
        d.Segment(d.Line("Ink", Lgf.Ink, 3.5), west, far);
        d.Scene(-3.6, 3.1, 4.3, -3.6);
    }

    // A circle, a big equilateral triangle and two upside-down ones on its bottom corners, their outer
    // top corners on the circle, and a rectangle on the left one with its top left corner on the
    // circle; pink lines from that corner and from the right triangle's outer corner to a point on the
    // circle. The angle there? The rectangle's corners on the circle make the center level with its
    // middle, so the two pink ends are opposite ends of a diameter. Dragging the point on the circle
    // moves the angle round, and dragging the rectangle's corner along the circle restacks the
    // triangles (the small ones a fixed share of the width): it stays a right angle.
    [Puzzle("three-triangles-and-a-rectangle-inside-a-circle")]
    public static void ThreeTrianglesAndARectangleInsideACircle(Lgf d)
    {
        var o = d.Pt(0, 0, "O");
        var rim = d.Pt(4, 0);
        var r = d.ArcSlider(o, 4, 164, 176, 165, (0, 0), "R");
        var x = d.ArcSlider(o, 4, 35, 140, 67.5, (0, 0), "X");
        string a = d.Num("-0.516 * R.X", "a");
        string b = d.Num("-1.484 * R.X", "b");
        string baseY = d.Num($"-R.Y - {s3} * {a}", "base");
        var p = d.Pt(new V("R.X", "-R.Y"), "P");
        var q = d.Pt(new V("-R.X", "-R.Y"), "Q");
        var pr = d.Pt(new V($"R.X + {a}", "-R.Y"), "Pr");
        var ql = d.Pt(new V($"-R.X - {a}", "-R.Y"), "Ql");
        var b1 = d.Pt(new V($"-{b} / 2", baseY), "B1");
        var b2 = d.Pt(new V($"{b} / 2", baseY), "B2");
        var apex = d.Pt(new V("0", $"{baseY} + {s3} * {b}"), "Ap");
        var top = d.Pt(new V($"R.X + {a}", "R.Y"));
        V[] big = [b1, b2, apex];
        V[] left = [p, pr, b1];
        V[] right = [ql, q, b2];
        d.Region(d.Fill("Lavender", "#FFEAD5FB", null), "Union", d.Polygon(null, big), d.Polygon(null, left), d.Polygon(null, right));
        d.Polygon(d.Fill("Cyan", "#FF92E7FA", "#FF02A4D0", 3.5), p, pr, top, r);
        string purple = d.Outline("Purple", "#FF48189A", 3.5);
        d.Polygon(purple, big);
        d.Polygon(purple, left);
        d.Polygon(purple, right);
        d.Circle(d.Outline("Circle", "#FF7030C0", 3.5), o, rim);
        string pink = "#FFF573C1";
        string line = d.Line("Pink", pink, 3.5);
        d.Segment(line, r, x);
        d.Segment(line, x, q);
        d.ArcMark(line, x, r, q, 0.42);
        string toR = d.Num("distance(X, R)", "xr");
        string toQ = d.Num("distance(X, Q)", "xq");
        var bisector = new V($"(R.X - X.X) / {toR} + (Q.X - X.X) / {toQ}", $"(R.Y - X.Y) / {toR} + (Q.Y - X.Y) / {toQ}");
        d.Label("?", x + Lgf.N(1.0 / Math.Sqrt(2)) * bisector, d.Text("PinkMark", pink, 32), 32);
        d.Scene(-4, 4, 4, -4);
    }

    // A square, a right triangle in its top right corner from the top side to the bottom right corner
    // covering a quarter of it, and an isosceles triangle on the left side whose third corner is on the
    // right triangle's long side (the equal angles marked). What fraction does it cover? Nothing to
    // drag: the quarter puts the right triangle's top corner at the middle of the top side, and the
    // equal sides put the third corner at (0.8, 0.4) of the side.
    [Puzzle("triangle-inside-a-square")]
    public static void TriangleInsideASquare(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 8;
        var bl = d.Pt(0, 0, "BL");
        var br = d.Pt(L, 0, "BR");
        var tr = d.Pt(L, L, "TR");
        var tl = d.Pt(0, L, "TL");
        var m = d.Pt(L / 2, L, "M");
        var p = d.Pt(0.8 * L, 0.4 * L, "P");
        d.Polygon(d.Fill("Yellow", "#FFFFFC02", null), tl, bl, p);
        d.Polygon(d.Fill("Magenta", "#FFBE0068", null), m, br, tr);
        string ink = d.Line("Ink", Lgf.Ink, 4);
        d.Segment(ink, tl, p);
        d.Segment(ink, bl, p);
        d.Segment(ink, m, br);
        d.Box(d.Outline("Frame", Lgf.Ink, 4.5), V.At(0, 0), V.At(L, L));
        // on the yellow, which stays bright in the dark theme: deep there too
        string mark = d.Line("Mark", "#FFC8B000", 2.5, dark: "#FF7A6A00");
        foreach (double radius in new[] { 0.75, 1.05 })
        {
            d.ArcMark(mark, bl, p, tl, radius);
            d.ArcMark(mark, p, tl, bl, radius);
        }

        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(0, L, L, -0.5);
    }

    // A semicircle and a big black triangle on its diameter, its sides crossing the arc at two corners
    // of an equilateral triangle whose third corner is the center. The blue triangle cut off at the top
    // is 3; the black triangle's area? Dragging the red triangle's left corner round the arc (the other
    // follows 60° on) moves the top: the blue triangle is always the big one halved (the chord is the
    // radius, the base the diameter), so the black one stays 12.
    [Puzzle("triangle-inside-circle-and-triangle")]
    public static void TriangleInsideCircleAndTriangle(Lgf d)
    {
        const double R = 4;
        var o = d.Pt(0, 0, "O");
        var a = d.Pt(-R, 0, "A");
        var b = d.Pt(R, 0, "B");
        var p = d.ArcSlider(o, R, 122, 170, 143.3, (0, 0), "P");
        var q = d.Pt(V.Of("P").Turn(-60), "Q");
        var t = d.Crossing(a, p, b, q, "T");
        d.Polygon(d.Fill("Cyan", "#FF80EEF5", null), t, p, q);
        d.Arc("CircleArc", d.Line("Blue", "#FF1008A5", 3.5), o, b, a);
        string red = d.Line("Red", "#FFD01818", 3.5);
        d.Segment(red, o, p);
        d.Segment(red, o, q);
        d.Segment(red, p, q);
        d.Inked(d.Outline("Frame", Lgf.Ink, 4.5), a, b, t);
        d.DotAt(o, "Od", 13);
        d.Label("3", Mid(d, t, p, q), d.ShapeText("BlueMark", "#FF0A85E5", 40), 40);
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(-R, 6.8, R, -0.5);
    }

    // A square, an equilateral triangle with a corner at its top right corner and the others on the
    // left and bottom sides, and a unit circle in the triangle. The total shaded area? Nothing to
    // drag: the unit circle makes the triangle's side 2√3, and the triangle fixes the square.
    [Puzzle("circle-in-triangle-in-square")]
    public static void CircleInTriangleInSquare(Lgf d)
    {
        d.Texture = "FeltTip";
        double s = 2 * Math.Sqrt(3), c15 = Math.Cos(Math.PI / 12), s15 = Math.Sin(Math.PI / 12);
        double L = s * c15;
        var tr = d.Pt(L, L);
        var p = d.Pt(0, L - s * s15);
        var q = d.Pt(L - s * s15, 0);
        var center = Mid(d, tr, p, q);
        string square = d.Polygon(null, d.Pt(0, 0), d.Pt(L, 0), tr, d.Pt(0, L));
        string triangle = d.Polygon(null, tr, p, q);
        d.Region(d.Fill("Orange", "#FFFEAA01", null), "Difference", square, triangle);
        d.Polygon(d.Outline("Red", "#FFE80605", 3.5, dark: "#FFFF4A3A"), tr, p, q);
        d.Circle(d.Outline("Purple", "#FF7A10C0", 3.5), center, d.Pt(center + V.At(1, 0)));
        d.Box(d.Outline("Frame", Lgf.Ink, 4.5), V.At(0, 0), V.At(L, L));
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(0, L, L, -0.25);
    }

    // A square in a circle and two quarter circles on opposite corners of the square; their overlap
    // is 16. The total shaded area (the circle outside the square)? Nothing to drag: the overlap fixes
    // the square, and both are (π/2 - 1) times its area.
    [Puzzle("quarter-circles-in-a-square-in-a-circle")]
    public static void QuarterCirclesInASquareInACircle(Lgf d)
    {
        d.Texture = "FeltTip";
        double h = Math.Sqrt(16 / (Math.PI / 2 - 1)) / 2;
        var o = d.Pt(0, 0, "O");
        var bl = d.Pt(-h, -h);
        var br = d.Pt(h, -h);
        var tr = d.Pt(h, h);
        var tl = d.Pt(-h, h);
        string circle = d.Circle(null, o, tr);
        d.Region(d.Fill("Magenta", "#FFF40074", null), "Difference", circle, d.Polygon(null, bl, br, tr, tl));
        d.Circle(d.Outline("Frame", Lgf.Ink, 3.5), o, tr);
        string teal = d.Line("Teal", "#FF20907A", 3.5);
        string lime = d.Line("Lime", "#FFB8F030", 3.5);
        d.Segment(teal, tl, tr);
        d.Segment(teal, tl, bl);
        d.Segment(lime, br, tr);
        d.Segment(lime, br, bl);
        d.Arc("CircleArc", teal, tl, bl, tr);
        d.Arc("CircleArc", lime, br, tr, bl);
        d.Label("16", V.At(0, 0), d.Text("Mark", Lgf.Ink, 34), 34);
        d.Scene(-h * Math.Sqrt(2), h * Math.Sqrt(2), h * Math.Sqrt(2), -h * Math.Sqrt(2));
    }

    // A regular hexagon of area 24 and an equilateral triangle from its top right corner to the middle
    // of its bottom side (the ticks), the third corner out on the left. The triangle's area? Nothing to
    // drag: the hexagon fixes it (13/4 of a side squared over the hexagon's 6, so 13).
    [Puzzle("triangle-overlapping-a-hexagon")]
    public static void TriangleOverlappingAHexagon(Lgf d)
    {
        double a = Math.Sqrt(16 / Math.Sqrt(3));
        var hexagon = HexagonFromLeft(d, V.At(-a, 0), Lgf.N(a));
        var v = hexagon[2];
        var m = Mid(d, hexagon[4], hexagon[5]);
        var w = d.Pt(v + (m - v).Turn(-60));
        d.Polygon(d.Fill("Lime", "#FFE6FDB8", null), v, m, w);
        string cyan = "#FF18C8DA";
        d.Polygon(d.Outline("Cyan", cyan, 4), hexagon);
        string ticks = d.Line("Ticks", cyan, 3);
        d.Segment(ticks, hexagon[5], m, "TwoTicks");
        d.Segment(ticks, m, hexagon[4], "TwoTicks");
        d.Inked(d.Outline("Green", "#FF8EE010", 3.5), v, m, w);
        d.Scene(-a * 1.25, a * S3, a, -a * S3 - 0.2);
    }
}
