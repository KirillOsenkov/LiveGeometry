#nullable disable

static partial class Puzzles
{
    // A semicircle and a quadrilateral on its diameter: two of its sides cross the arc, cutting it
    // into a red arc in the middle and blue arcs that add up to the red one (a quarter of the circle
    // each). Its angle at the top (a) and at the right end of the diameter (b): a + b? The point where
    // the left side crosses the arc is dragged along it (the red arc goes with it, staying a quarter),
    // and so is the corner on the arc at the right: the angles change, their sum stays 135°.
    [Puzzle("quadrilateral-splitting-a-semi-circle")]
    public static void QuadrilateralSplittingASemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4;
        var o = d.Pt(0, 0, "O");
        var w = d.Pt(-R, 0, "W");
        var e = d.Pt(R, 0, "E");
        var x = d.ArcSlider(o, R, 140, 168, 152, (0, 0), "X");
        var y = d.Pt(V.Of("X").Turn(-90), "Y");
        var q = d.ArcSlider(o, R, 12, 46, 40, (0, 0), "Q");
        var p = d.Crossing(w, x, q, y, "P");
        string blue = d.Line("Blue", "#FF0A4CF8", 3);
        d.Arc("CircleArc", blue, o, x, w);
        d.Arc("CircleArc", blue, o, q, y);
        d.Arc("CircleArc", blue, o, e, q);
        d.Arc("CircleArc", d.Line("Red", "#FFF90606", 3), o, y, x);
        d.Inked(d.Outline("Navy", "#FF2E3F9D", 3), w, e, q, p);
        string orange = "#FFF2A12C";
        string mark = d.Line("Orange", orange, 2.5);
        d.ArcMark(mark, p, w, q, 0.42);
        d.ArcMark(mark, e, q, w, 0.42);
        string text = d.Text("OrangeMark", orange, 30);
        d.Label("a", p + 0.75 * UnitBisector(d, p, w, q), text);
        d.Label("b", e + 0.8 * UnitBisector(d, e, q, w), text);
        d.Label("a+b = ?", V.At(2.3, 5.1), text);
        d.Scene(-R, 6.3, R, 0);
    }

    /// <summary>The unit vector halving the angle at the vertex between the directions to a and b (an angle under 180°)</summary>
    private static V UnitBisector(Lgf d, V vertex, V a, V b)
    {
        string toA = d.Num($"distance({vertex.X[..^2]}, {a.X[..^2]})");
        string toB = d.Num($"distance({vertex.X[..^2]}, {b.X[..^2]})");
        var sum = d.Pt(new V($"({a.X} - {vertex.X}) / {toA} + ({b.X} - {vertex.X}) / {toB}", $"({a.Y} - {vertex.Y}) / {toA} + ({b.Y} - {vertex.Y}) / {toB}"));
        string length = d.Num($"sqrt({sum.X} ^ 2 + {sum.Y} ^ 2)");
        return new V($"{sum.X} / {length}", $"{sum.Y} / {length}");
    }

    /// <summary>A regular hexagon with a vertex at the top (upright sides), given its center and side: top, upper left, lower left, bottom, lower right, upper right</summary>
    private static V[] PointyHexagon(Lgf d, V center, string side) =>
    [
        d.Pt(center + new V("0", side)),
        d.Pt(center + new V($"-{s3} * {side}", $"0.5 * {side}")),
        d.Pt(center + new V($"-{s3} * {side}", $"-0.5 * {side}")),
        d.Pt(center + new V("0", $"-1 * {side}")),
        d.Pt(center + new V($"{s3} * {side}", $"-0.5 * {side}")),
        d.Pt(center + new V($"{s3} * {side}", $"0.5 * {side}")),
    ];

    /// <summary>A double arrow from a to b (two vectors from the middle) with a text beside its middle, the given distance to its left (units)</summary>
    private static void DoubleArrow(Lgf d, string lineStyle, string textStyle, V a, V b, string text, double textOffset)
    {
        var middle = Mid(d, a, b);
        d.Vector(lineStyle, middle, a);
        d.Vector(lineStyle, middle, b);
        string length = d.Num($"distance({a.X[..^2]}, {b.X[..^2]})");
        d.Label(text, middle + Lgf.N(textOffset) * new V($"({a.Y} - {b.Y}) / {length}", $"({b.X} - {a.X}) / {length}"), textStyle);
    }

    // Two regular hexagons, upright, the small one's top right corner on the big one's top left
    // corner; their tops are 6 apart. How long is the pink line, from the small one's lower left
    // corner to the big one's lower right? The small one's top is dragged round the big one's top,
    // 6 away: the sizes trade, the pink line stays 12.
    [Puzzle("two-regular-hexagons-ii")]
    public static void TwoRegularHexagonsII(Lgf d)
    {
        var top = d.Pt(0, 0, "T");
        var t1 = d.ArcSlider(top, 6, 187, 197, 191, (0, 0), "S");
        const string sqrt3 = "1.7320508075688772";
        string a = d.Num($"-S.X / {sqrt3} + S.Y", "a");
        string b = d.Num($"-S.X / {sqrt3} - S.Y", "b");
        var v = d.Pt(top - new V($"{s3} * {b}", $"0.5 * {b}"), "V");
        var small = PointyHexagon(d, d.Pt(v - new V($"{s3} * {a}", $"0.5 * {a}")), a);
        var big = PointyHexagon(d, d.Pt(v + new V($"{s3} * {b}", $"-0.5 * {b}")), b);
        string navy = "#FF0201BD";
        d.Polygon(d.Fill("Blue", "#FF7E83E1", navy, 3.5), small);
        d.Polygon(d.Fill("Lavender", "#FFD2DAF4", navy, 3.5), big);
        DoubleArrow(d, d.Line("Arrow", Lgf.Ink, 3), d.Text("Mark", Lgf.Ink, 32), t1, top, "6", 0.6);
        string magenta = "#FFE602C8";
        DoubleArrow(d, d.Line("Pink", magenta, 3.5), d.ShapeText("PinkMark", magenta, 36), small[2], big[4], "?", 1.2);
        d.Scene(-8.1, 1.1, 4.1, -9.6);
    }

    // Three squares stacked in a circle: the small one (side 6) on the left of the middle one, the
    // middle one on the right of the big one, 6 in from its left side; the big one's bottom corners and
    // the small one's top corners on the circle. The circle's area? Nothing to drag: the 6s make the
    // squares 6, 12 and 18, and the circle is the one through those four corners.
    [Puzzle("three-squares-in-a-circle-ii")]
    public static void ThreeSquaresInACircleII(Lgf d)
    {
        var o = d.Pt(9, 17, "O");
        var rim = d.Pt(0, 0);
        d.Circle(d.Outline("Circle", "#FF0401D0", 3.5), o, rim);
        string blue = d.Fill("Blue", "#FF0999ED", Lgf.Ink, 4.5);
        d.Box(blue, V.At(0, 0), V.At(18, 18));
        d.Box(blue, V.At(6, 18), V.At(18, 30));
        d.Box(blue, V.At(6, 30), V.At(12, 36));
        string cyan = "#FF2EB8EE";
        string arrow = d.Line("Arrow", cyan, 3);
        string text = d.Text("CyanMark", cyan, 26);
        DoubleArrow(d, arrow, text, d.Pt(4.6, 30.4), d.Pt(4.6, 35.6), "6", 2.1);
        DoubleArrow(d, arrow, text, d.Pt(12.4, 31.3), d.Pt(17.6, 31.3), "6", 2.1);
        DoubleArrow(d, arrow, text, d.Pt(0.4, 19.3), d.Pt(5.6, 19.3), "6", 2.1);
        double r = Math.Sqrt(370);
        d.Scene(9 - r, 17 + r, 9 + r, 17 - r);
    }

    // Two rectangles: ABCD and AECF across it, sharing the diagonal AC. AB = 5, AE = 7, AF = 1; the
    // angle AEB? Nothing to drag: the numbers fix it (AC² = 50 = 2 AB², so ABCD is a square, and E is at
    // (5.6, 4.2)).
    [Puzzle("rectangle-across-a-rectangle")]
    public static void RectangleAcrossARectangle(Lgf d)
    {
        d.Texture = "FeltTip";
        var a = d.Pt(0, 0, "A");
        var b = d.Pt(5, 0, "B");
        var c = d.Pt(5, 5, "C");
        var topLeft = d.Pt(0, 5, "D");
        var e = d.Pt(5.6, 4.2, "E");
        var f = d.Pt(-0.6, 0.8, "F");
        d.Polygon(d.Outline("Red", "#FFD0140A", 3.5), a, b, c, topLeft);
        d.Polygon(d.Outline("Blue", "#FF6A4CDE", 3.5), a, e, c, f);
        string orange = d.Line("Orange", "#FFFA9208", 3.5);
        d.Segment(orange, e, b);
        d.ArcMark(orange, e, a, b, 0.45);
        string text = d.Text("Letter", Lgf.Ink, 28);
        d.Label("A", V.At(-0.05, -0.55), text);
        d.Label("B", V.At(5.15, -0.55), text);
        d.Label("C", V.At(5.35, 5.3), text);
        d.Label("D", V.At(-0.45, 5.35), text);
        d.Label("E", V.At(6.0, 4.15), text);
        d.Label("F", V.At(-1.0, 0.95), text);
        d.Scene(-1.3, 5.7, 6.4, -0.9);
    }

    // Two squares side by side and an equilateral triangle on their bottom line, as wide as both, its
    // right side through the big square's top right corner. The angle at the small square's top right
    // corner, between its top side and the dashed line to the triangle's top? Nothing to drag: the
    // corner on the triangle's side makes the big square √3 times the small one.
    [Puzzle("two-squares-and-an-equilateral-triangle")]
    public static void TwoSquaresAndAnEquilateralTriangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double b = 2;
        double a = b * Math.Sqrt(3);
        string frame = d.Outline("Frame", Lgf.Ink, 6);
        d.Box(frame, V.At(0, 0), V.At(a, a));
        d.Box(frame, V.At(a, 0), V.At(a + b, b));
        var left = d.Pt(0, 0);
        var right = d.Pt(a + b, 0);
        var apex = d.Pt((a + b) / 2, (a + b) * S3, "T");
        d.Polygon(d.Outline("Green", "#FF088E58", 4), left, right, apex);
        var corner = d.Pt(a + b, b, "K");
        var inner = d.Pt(a, b);
        string lime = d.Line("Lime", "#FF80B502", 4);
        d.Dashed(lime, apex, corner, 7, 0.55);
        d.ArcMark(lime, corner, apex, inner, 0.33);
        d.Scene(0, (a + b) * S3, a + b, 0);
    }

    // A rectangle in a circle and two quarter circles from its opposite corners, their arcs meeting
    // at its middle and crossing the diagonal at its thirds (the dots; the ticks say the three parts
    // are equal). What fraction is shaded? Nothing to drag: the thirds fix the rectangle's shape
    // (width² = 1.4 height²), and the radius is the circle's.
    [Puzzle("quarter-circles-in-a-rectangle-in-a-circle")]
    public static void QuarterCirclesInARectangleInACircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double h = 5;
        double w = h * Math.Sqrt(1.4), r = Math.Sqrt(w * w + h * h) / 2;
        var bl = d.Pt(0, 0);
        var br = d.Pt(w, 0);
        var tr = d.Pt(w, h);
        var tl = d.Pt(0, h);
        var m = d.Pt(w / 2, h / 2, "M");
        string teal = "#FF05BFA9";
        string cyan = d.Fill("Cyan", "#FF10EDE6", teal, 3);
        d.Arc("CircleSector", cyan, tl, d.Pt(0, h - r), d.Pt(r, h));
        d.Arc("CircleSector", cyan, br, d.Pt(w, r), d.Pt(w - r, 0));
        d.Inked(d.Outline("Teal", teal, 3.5), bl, br, tr, tl);
        d.Circle(d.Outline("Circle", "#FF0C10E0", 4), m, bl);
        var p1 = d.Pt(w / 3, h / 3);
        var p2 = d.Pt(2 * w / 3, 2 * h / 3);
        string blue = d.Line("Blue", "#FF0A8FE0", 3.5);
        d.Segment(blue, bl, p1, "TwoTicks");
        d.Segment(blue, p1, p2, "TwoTicks");
        d.Segment(blue, p2, tr, "TwoTicks");
        d.DotAt(p1, "P1");
        d.DotAt(p2, "P2");
        d.Scene(w / 2 - r, h / 2 + r, w / 2 + r, h / 2 - r);
    }

    // An isosceles triangle cut into three equal areas: the pink triangle (isosceles, its equal sides
    // ticked) from a point on the left side to the bottom right corner and the middle of the base.
    // The angle at that point, between the side and the line to the corner? Nothing to drag: the equal
    // areas put the point a third of the way down and the pink corner at the middle of the base, and
    // the pink isosceles triangle then makes the outer one √2 times as high as half its base.
    [Puzzle("nested-isosceles-triangles")]
    public static void NestedIsoscelesTriangles(Lgf d)
    {
        const double half = 3;
        double height = half * Math.Sqrt(2);
        var a = d.Pt(-half, 0, "A");
        var c = d.Pt(half, 0, "C");
        var t = d.Pt(0, height, "T");
        var p = d.Pt(-half / 3, 2 * height / 3, "P");
        var q = d.Pt(0, 0, "Q");
        d.Polygon(d.Fill("Yellow", "#FFFFFF68", null), a, q, p);
        d.Polygon(d.Fill("Pink", "#FFFDABDB", null), p, q, c);
        d.Polygon(d.Fill("Orange", "#FFFDCD7A", null), p, c, t);
        // between the fills, which are mid-tones in the dark theme: deep there
        string pink = d.Line("PinkLine", "#FFE85A5A", 3, dark: "#FF8A0A3A");
        d.Segment(pink, p, q, "OneTick");
        d.Segment(pink, p, c);
        d.Segment(pink, q, c, "OneTick");
        string ink = d.Line("Ink", Lgf.Ink, 4.5);
        d.Segment(ink, a, c);
        d.Segment(ink, a, t, "TwoTicks");
        d.Segment(ink, t, c, "TwoTicks");
        string orange = "#FFF7901A";
        d.ArcMark(d.Line("OrangeLine", orange, 3, dark: "#FF8A3200"), p, c, t, 0.35);
        d.Label("?", p + V.At(0.66, 0.2), d.ShapeText("OrangeMark", orange, 36));
        d.Scene(-half, height, half, 0);
    }

    // Five equal circles in a row in a rectangle that just holds them, its top corners on a
    // semicircle over its bottom side. Each circle has area 1; the semicircle's? Nothing to drag: the
    // radius is √29 circles' radii.
    [Puzzle("five-circles-in-a-rectangle-in-a-semi-circle")]
    public static void FiveCirclesInARectangleInASemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        double big = Math.Sqrt(29);
        var o = d.Pt(0, 0, "O");
        d.Arc("CircleSegment", d.Outline("Blue", "#FF0D5AD7", 4), o, d.Pt(big, 0), d.Pt(-big, 0));
        // over the semicircle's diameter
        d.Inked(d.Outline("Cyan", "#FF1AA8D0", 3), d.Pt(-5, 0), d.Pt(5, 0), d.Pt(5, 2), d.Pt(-5, 2));
        (string Name, string Color, string Outline)[] circles =
        [
            ("Red", "#FFFE584B", "#FFE3424C"),
            ("Orange", "#FFFE8D02", "#FFF26A0E"),
            ("Yellow", "#FFFFFF01", "#FFF5EE00"),
            ("Lime", "#FF9EF902", "#FF62C51A"),
            ("Green", "#FF02C76C", "#FF0A8476"),
        ];
        for (int i = 0; i < 5; i++)
        {
            var center = d.Pt(-4 + 2 * i, 1);
            d.Circle(d.Fill(circles[i].Name, circles[i].Color, circles[i].Outline, 3), center, d.Pt(-3 + 2 * i, 1));
        }

        d.Scene(-big, big, big, 0);
    }

    // A rectangle, a semicircle on its bottom side and a smaller one hanging from its top side, both
    // ending at its right corners and touching each other; the red line from the top left corner is
    // their common tangent there. What fraction is shaded (under the red line)? Dragging the top left
    // corner along the top changes the rectangle's width (its height stays) and the two semicircles
    // with it: the red line always ends halfway up the right side, the shaded part stays 3/4.
    [Puzzle("two-semi-circles-inside-a-rectangle")]
    public static void TwoSemiCirclesInsideARectangle(Lgf d)
    {
        const double H = 4;
        var corner = d.Slider(V.At(-5.85, H), V.At(-7.8, H), (6.72 - 5.85) / 1.95, (-6.72, H), "L");
        // the semicircles touch: (R + r)² = (R - r)² + H², so R r = H² / 4; the tangent from the corner
        // is as long to both: R - r = H² / (2 W)
        string gap = d.Num($"{Lgf.N(H * H / 2)} / (-L.X)", "D");
        string big = d.Num($"({gap} + sqrt({gap} ^ 2 + {Lgf.N(H * H)})) / 2", "R");
        string small = d.Num($"{big} - {gap}", "r");
        var bottomLeft = d.Pt(new V("L.X", "0"), "BL");
        var bottomRight = d.Pt(0, 0, "BR");
        var topRight = d.Pt(0, H, "TR");
        var middle = d.Pt(0, H / 2, "MR");
        d.Polygon(d.Fill("Pink", "#FFFCA6A8", null), bottomLeft, bottomRight, middle, corner);
        var bigCenter = d.Pt(new V($"-{big}", "0"), "CB");
        var smallCenter = d.Pt(new V($"-{small}", Lgf.N(H)), "CS");
        string ink = d.Line("Ink", Lgf.Ink, 4);
        d.Arc("CircleArc", ink, bigCenter, bottomRight, d.Pt(new V($"-2 * {big}", "0")));
        d.Arc("CircleArc", ink, smallCenter, d.Pt(new V($"-2 * {small}", Lgf.N(H))), topRight);
        d.Segment(d.Line("Red", "#FFE30A22", 3), corner, middle);
        d.Inked(d.Outline("Frame", Lgf.Ink, 4.5), bottomLeft, bottomRight, topRight, corner);
        d.Scene(-7.2, H, 0, 0);
    }

    // An equilateral triangle and two squares inside it, sharing their left side's line: one standing
    // on the base, the other on top of it with its top corners on the triangle's sides. The angle at
    // their shared corner between the line to the top of the triangle and the bottom square's
    // diagonal? That corner is dragged along the line from the top (always 75° to the level): the
    // squares trade sizes, the angle stays 120°.
    [Puzzle("two-squares-inside-an-equilateral-triangle")]
    public static void TwoSquaresInsideAnEquilateralTriangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double s = 6;
        double height = s * S3;
        double c75 = Math.Cos(75 * Math.PI / 180), s75 = Math.Sin(75 * Math.PI / 180);
        V At(double b) => V.At(-c75 * (height - b) / s75, b);
        var apex = d.Pt(0, height, "T");
        var p = d.Slider(At(2.3), At(0.9), (2.3 - 2.07) / 1.4, (-c75 * (height - 2.07) / s75, 2.07), "P");
        string b = "P.Y", a = "(-2 * P.X)";
        var left = d.Pt(-s / 2, 0);
        var right = d.Pt(s / 2, 0);
        d.Polygon(d.Fill("Cyan", "#FF03DFFC", null), left, right, apex);
        string navy = "#FF062E6D";
        string blue = d.Fill("Blue", "#FF03A4F2", navy, 4);
        var foot = d.Pt(new V($"P.X + {b}", "0"), "F");
        d.Box(blue, new V("P.X", "0"), new V(foot.X, b));
        d.Box(blue, V.Of("P"), new V($"-1 * P.X", $"{b} + {a}"));
        d.Inked(d.Outline("Frame", navy, 4.5), left, right, apex);
        d.Segment(d.Line("Royal", "#FF0737CD", 3), p, apex);
        string line = d.Line("Navy", navy, 3);
        d.Segment(line, p, foot);
        d.ArcMark(line, p, foot, apex, 0.4);
        d.Label("?", p + V.At(0.72, 0.6), d.ShapeText("NavyMark", "#FF0A2A8A", 36));
        d.Scene(-s / 2, height, s / 2, 0);
    }

    // A square, its top and bottom sides cut in thirds (dashed lines), and a rectangle with its
    // corners on the four sides, two of them at the thirds of the top and bottom. What fraction of the
    // square does it cover? Nothing to drag: the right angles put the other two corners at the thirds
    // of the sides too.
    [Puzzle("tilted-rectangle-in-a-square")]
    public static void TiltedRectangleInASquare(Lgf d)
    {
        d.Texture = "FeltTip";
        const double S = 6;
        d.Polygon(d.Fill("Orange", "#FFFDDC66", "#FFF28A04", 3.5), d.Pt(2, 0), d.Pt(6, 4), d.Pt(4, 6), d.Pt(0, 2));
        string dash = d.Line("Dash", Lgf.Ink, 3.5);
        d.Dashed(dash, d.Pt(2, S), d.Pt(2, 0), 9, 0.6);
        d.Dashed(dash, d.Pt(4, S), d.Pt(4, 0), 9, 0.6);
        string ink = d.Line("Ink", Lgf.Ink, 4.5);
        V[] top = [d.Pt(0, S), d.Pt(2, S), d.Pt(4, S), d.Pt(6, S)];
        V[] bottom = [d.Pt(0, 0), d.Pt(2, 0), d.Pt(4, 0), d.Pt(6, 0)];
        for (int i = 0; i < 3; i++)
        {
            d.Segment(ink, top[i], top[i + 1], "TwoTicks");
            d.Segment(ink, bottom[i], bottom[i + 1], "TwoTicks");
        }

        d.Segment(ink, top[0], bottom[0]);
        d.Segment(ink, top[3], bottom[3]);
        d.Scene(0, S, S, 0);
    }

    // A quarter circle with a circle cut out of it, the circle on its bottom side; the chord across
    // the top of the circle, level, is 12. The total shaded area? The circle's center is dragged (it
    // grows and moves, the quarter circle grows with it so that the chord stays 12): the shaded area
    // stays 36π.
    [Puzzle("circle-in-a-quarter-circle")]
    public static void CircleInAQuarterCircle(Lgf d)
    {
        var o = d.Pt(0, 0, "O");
        string room = d.Polygon(null, d.Pt(3.8, 3), d.Pt(9.2, 3), d.Pt(8.4, 5), d.Pt(6, 5));
        var c = d.HandleIn(room, 8.26, 4.64, "C");
        string radius = d.Num("sqrt(144 + 4 * C.Y ^ 2)", "R");
        var east = d.Pt(new V(radius, "0"), "E");
        var north = d.Pt(new V("0", radius), "N");
        var foot = d.Pt(new V("C.X", "0"), "F");
        string quarter = d.Arc("CircleSector", null, o, east, north, visible: false);
        string hole = d.Circle(null, c, foot);
        d.Region(d.Fill("Purple", "#FFBC9CEC", null), "Difference", quarter, hole);
        d.Arc("CircleSector", d.Outline("Frame", Lgf.Ink, 3.5), o, east, north);
        d.Circle(d.Outline("Circle", Lgf.Ink, 3.5), c, foot);
        DoubleArrow(d, d.Line("Arrow", "#FFFFFFFF", 3, dark: "#FFFFFFFF"), d.Text("WhiteMark", "#FFFFFFFF", 30, dark: "#FFFFFFFF"),
            d.Pt(new V("0.15", "2 * C.Y")), d.Pt(new V("11.85", "2 * C.Y")), "12", 0.9);
        d.Scene(0, 15.4, 15.4, 0);
    }

    // Three semicircles on upright diameters: the smallest bulging left, its diameter a chord of the
    // middle one; the middle one bulging left, its diameter a chord of the largest, which bulges right.
    // The middle one is cut into three stripes of equal width by the other two diameters. The smallest
    // has area 5; the largest? Nothing to drag: the equal stripes fix the radii (√5, 3 and √10 stripes).
    [Puzzle("three-semi-circles")]
    public static void ThreeSemiCircles(Lgf d)
    {
        double ra = Math.Sqrt(5), rb = 3, rc = Math.Sqrt(10);
        var ca = d.Pt(0, 0, "A");
        var cb = d.Pt(2, 0, "B");
        var cc = d.Pt(1, 0, "C");
        var aTop = d.Pt(0, ra);
        var aBottom = d.Pt(0, -ra);
        var bTop = d.Pt(2, rb);
        var bBottom = d.Pt(2, -rb);
        var cTop = d.Pt(1, rc);
        var cBottom = d.Pt(1, -rc);
        string middle = d.Arc("CircleSegment", null, cb, bTop, bBottom, visible: false);
        string Stripe(double from, double to) => d.Polygon(null, d.Pt(from, -4), d.Pt(to, -4), d.Pt(to, 4), d.Pt(from, 4));
        d.Region(d.Fill("Yellow", "#FFFEFE5E", null), "Intersection", middle, Stripe(-1.5, 0));
        d.Region(d.Fill("Green", "#FFB6F173", null), "Intersection", middle, Stripe(0, 1));
        d.Region(d.Fill("Blue", "#FF0EC4F3", null), "Intersection", middle, Stripe(1, 2.5));
        string frame = d.Outline("Frame", Lgf.Ink, 4);
        d.Arc("CircleSegment", frame, ca, aTop, aBottom);
        d.Arc("CircleSegment", frame, cb, bTop, bBottom);
        d.Arc("CircleSegment", frame, cc, cBottom, cTop);
        d.Scene(-ra, rc, 1 + rc, -rc);
    }
}
