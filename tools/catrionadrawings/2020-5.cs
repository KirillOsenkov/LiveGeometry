#nullable disable

static partial class Puzzles
{
    // An equilateral triangle of side 8 cut by a level line into two equal areas, and a semicircle on
    // that line through the top corner. The semicircle's area? Nothing to drag: the halving puts the
    // line where the top triangle's side is 8 / √2, and the semicircle's radius is that triangle's height.
    [Puzzle("semi-circle-splitting-a-triangle")]
    public static void SemiCircleSplittingATriangle(Lgf d)
    {
        d.Texture = "FeltTip";
        double h = 4 * Math.Sqrt(3), r = 2 * Math.Sqrt(6), y = h - r, half = 2 * Math.Sqrt(2);
        var a = d.Pt(-4, 0);
        var b = d.Pt(4, 0);
        var t = d.Pt(0, h);
        var c = d.Pt(0, y);
        var east = d.Pt(r, y);
        var west = d.Pt(-r, y);
        var left = d.Pt(-half, y);
        var right = d.Pt(half, y);
        d.Polygon(d.Fill("Yellow", "#FFFFFF01", null), left, right, t);
        d.Polygon(d.Fill("Blue", "#FF0482D6", null), a, b, right, left);
        string yellow = d.Line("YellowLine", "#FFFEF000", 4);
        d.Arc("CircleArc", yellow, c, east, west);
        d.Segment(yellow, west, east);
        d.Inked(d.Outline("Frame", "#FF0565C8", 3.5), a, b, t);
        string blue = "#FF0390DC";
        d.Dimension(d.Line("Arrow", blue, 3), d.Text("BlueMark", blue, 34), b, a, 0.85, "8", 0.4);
        d.Scene(-r, h, r, -1.1);
    }

    // A square, its diagonal, and two semicircles on it: one from the bottom left corner to a point on
    // the diagonal, on the upper side, the other from there to the top right corner, on the lower side.
    // The point is 4 from the bottom right corner; the shaded area? Dragging the point (it stays 4 from
    // that corner) slides it along the diagonal and resizes the square: the semicircles trade sizes,
    // their squared diameters always add up to twice 4², and the shaded area stays.
    [Puzzle("two-semi-circles-diagonally-across-a-square")]
    public static void TwoSemiCirclesDiagonallyAcrossASquare(Lgf d)
    {
        var k = d.Pt(0, 0, "K");
        var p = d.ArcSlider(k, 4, 106, 164, 121, (0, 0), "P");
        string s = d.Num("P.Y - P.X", "s");
        var bl = d.Pt(new V($"-{s}", "0"), "BL");
        var tr = d.Pt(new V("0", s), "TR");
        var tl = d.Pt(new V($"-{s}", s), "TL");
        string green = d.Fill("Green", "#FFB8FB18", "#FF22D84A", 3);
        d.Arc("CircleSegment", green, Mid(d, bl, p), p, bl);
        d.Arc("CircleSegment", green, Mid(d, p, tr), p, tr);
        d.Inked(d.Outline("Frame", "#FF0A0A30", 3.5, dark: Lgf.DarkInk), bl, k, tr, tl);
        string teal = "#FF08E0C8";
        d.Dimension(d.Line("Arrow", teal, 3, dark: "#FF3FF0DA"), d.Text("TealMark", teal, 34, dark: "#FF3FF0DA"), p, k, 0, "4", 0.42);
        d.Scene(-6.3, 5.7, 0.6, -0.1);
    }

    // Two rectangles side by side on a line, lines from the big one's top corners through the small
    // one's top corners meeting on the line; the purple and the red triangle both have area 10. The
    // rectangles' total area? The lines meeting makes the rectangles similar, the equal areas make the
    // small one half the size. Dragging the point where the lines meet changes the big rectangle's
    // shape (its area stays 40): the triangles stay 10, the total stays.
    [Puzzle("two-neighbouring-rectangles")]
    public static void TwoNeighbouringRectangles(Lgf d)
    {
        var o = d.Pt(0, 0, "O");
        var f = d.Slider(V.At(12, 0), V.At(18, 0), (14.8 - 12) / 6, (14.8, 0), "F");
        string w = d.Num("F.X / 2", "w");
        string h = d.Num("80 / F.X", "h");
        var tl = d.Pt(new V("0", h), "TL");
        var tr = d.Pt(new V(w, h), "TR");
        var q = d.Pt(new V(w, $"{h} / 2"), "Q");
        var br = d.Pt(new V(w, "0"), "BR");
        var smallTop = d.Pt(new V("0.75 * F.X", $"{h} / 2"), "ST");
        var smallFoot = d.Pt(new V("0.75 * F.X", "0"), "SF");
        d.Polygon(d.Fill("Purple", "#FFAF82EE", null), tl, tr, q);
        d.Polygon(d.Fill("Red", "#FFFB7AA0", null), tr, f, q);
        string purple = d.Line("PurpleLine", "#FF7A10C8", 3);
        d.Segment(purple, tl, f);
        d.Segment(purple, tr, f);
        string ink = d.Line("Ink", Lgf.Ink, 3);
        d.Segment(ink, d.Pt(-0.6, 0), d.Pt(new V("F.X + 1.2", "0")));
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        d.Inked(frame, o, br, tr, tl);
        d.Inked(frame, br, smallFoot, smallTop, q);
        d.Scene(-0.6, 5.6, 16.2, -0.2);
    }

    // A semicircle and a smaller one on a chord of it, bulging toward the diameter; the big one has
    // twice the small one's area. The red line from the chord's end to where the small arc meets the
    // diameter: the angle between it and the chord? Half the area makes the chord a right angle at the
    // center, so the small arc runs through the center. Dragging the chord's end along the arc turns
    // the chord: the angle stays.
    [Puzzle("two-overlapping-semi-circles")]
    public static void TwoOverlappingSemiCircles(Lgf d)
    {
        const double R = 5;
        var o = d.DotAt(V.At(0, 0), "O");
        var east = d.Pt(R, 0);
        var west = d.Pt(-R, 0);
        var b = d.ArcSlider(o, R, 32, 80, 64.5, (0, 0), "B");
        var a = d.DotAt(b.Turn(90), "A");
        var m = Mid(d, a, b);
        string teal = d.Line("Teal", "#FF06CCB8", 3);
        d.Arc("CircleArc", teal, o, east, west);
        d.Segment(teal, west, east);
        string navy = d.Line("Navy", "#FF04037E", 3.5);
        d.Arc("CircleArc", navy, m, a, b);
        d.Segment(navy, a, b);
        string redColor = "#FFF52A5C";
        string red = d.Line("Red", redColor, 3.5);
        d.Segment(red, b, o);
        // a mark in units, so that the question mark beside it keeps clear of it on a phone too
        d.ArcMark(red, b, a, o, 0.75);
        // on the bisector of the 45° angle: the two unit vectors add up to 2 cos 22.5° long
        var bisector = new V($"(A.X - B.X) / {Lgf.N(R * Math.Sqrt(2))} - B.X / {Lgf.N(R)}", $"(A.Y - B.Y) / {Lgf.N(R * Math.Sqrt(2))} - B.Y / {Lgf.N(R)}");
        d.Label("?", b + (1.4 / 1.8477590650225735) * bisector, d.Text("RedMark", redColor, 30));
        d.Scene(-R, R, R, -0.7);
    }

    // A semicircle, a level chord of it with a semicircle hanging from it, and a yellow semicircle on
    // the diameter touching the chord, its ends where the hanging arc crosses the diameter. What
    // fraction of the big semicircle is yellow? Nothing to drag: the yellow ends on the hanging arc put
    // the chord at the big radius over √3.
    [Puzzle("nested-semi-circles")]
    public static void NestedSemiCircles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 6;
        double h = R / Math.Sqrt(3), rho = R * Math.Sqrt(2.0 / 3);
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0);
        var west = d.Pt(-R, 0);
        var yellowEast = d.Pt(h, 0);
        var yellowWest = d.Pt(-h, 0);
        var c = d.Pt(0, h);
        var chordLeft = d.Pt(-rho, h);
        var chordRight = d.Pt(rho, h);
        d.Arc("CircleSegment", d.Fill("Yellow", "#FFFFF601", "#FFFCE41A", 3), o, yellowEast, yellowWest);
        string orange = d.Line("Orange", "#FFF29A1A", 3);
        d.Arc("CircleArc", orange, c, chordLeft, chordRight);
        d.Segment(orange, chordLeft, chordRight);
        string crimson = d.Line("Crimson", "#FFD2145C", 3.5);
        d.Arc("CircleArc", crimson, o, east, west);
        d.Segment(crimson, west, yellowWest);
        d.Segment(crimson, yellowEast, east);
        d.Scene(-R, R, R, h - rho);
    }

    // An equilateral triangle point down, two equal circles side by side along its top, each touching
    // the other, the top and a slanted side; below the line under the circles it is shaded. What
    // fraction? Nothing to drag: the circles fix the line.
    [Puzzle("two-circles-in-a-triangle-ii")]
    public static void TwoCirclesInATriangleII(Lgf d)
    {
        const double A = 4;
        double r = A / (1 + Math.Sqrt(3)), y = -2 * r, x = A - 2 * r / Math.Sqrt(3);
        var topLeft = d.Pt(-A, 0);
        var topRight = d.Pt(A, 0);
        var apex = d.Pt(0, -A * Math.Sqrt(3));
        var left = d.Pt(-x, y);
        var right = d.Pt(x, y);
        d.Polygon(d.Fill("Teal", "#FF12EFF1", null), left, right, apex);
        d.Segment(d.Line("Blue", "#FF048CF8", 3.5), left, right);
        string green = d.Line("Green", "#FF10BC90", 3);
        d.Circle(green, d.Pt(-r, -r), d.Pt(-r, 0));
        d.Circle(green, d.Pt(r, -r), d.Pt(r, 0));
        d.Inked(d.Outline("Frame", Lgf.Ink, 3.5), topLeft, topRight, apex);
        d.Scene(-A, 0, A, -A * Math.Sqrt(3));
    }

    /// <summary>A double arrow from a to b drawn whole (two arrows from its middle), its text beside it at the given place, as she writes a length over its arrow</summary>
    private static void DoubleArrowWithTextBeside(Lgf d, string lineStyle, string textStyle, V a, V b, string text, V textAt)
    {
        var middle = Mid(d, a, b);
        d.Vector(lineStyle, middle, a);
        d.Vector(lineStyle, middle, b);
        d.Label(text, textAt, textStyle);
    }

    // A regular hexagon on a line, a line from its bottom left corner through its right corner, and a
    // unit circle touching both lines and the hexagon's lower right side. The hexagon's side? Nothing
    // to drag: the lines meet at 30°, so the circle's center is (2 + √3, 1), and being 1 from the
    // lower right side (at 60° through (s, 0)) makes the side 2. It touches the slanted line just
    // past the hexagon's corner, as she drew it.
    [Puzzle("tangential-hexagon-and-circle")]
    public static void TangentialHexagonAndCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        double s = 2;
        var hexagon = HexagonUp(d, V.At(0, 0), Lgf.N(s));
        var center = d.Pt(2 + Math.Sqrt(3), 1);
        var rim = d.Pt(3 + Math.Sqrt(3), 1);
        string ink = d.Line("Ink", Lgf.Ink, 3.5);
        d.Segment(ink, hexagon[0], d.Pt(5.9, 0));
        d.Segment(ink, hexagon[0], d.Pt(6.5 * Math.Sqrt(3) / 2, 3.25));
        string orange = "#FFEE4A10";
        d.Inked(d.Outline("Hexagon", orange, 3), hexagon);
        string magenta = "#FFE402A0";
        d.Circle(d.Line("Circle", magenta, 4), center, rim);
        DoubleArrowWithTextBeside(d, d.Line("MagentaArrow", magenta, 2.5), d.Text("MagentaMark", magenta, 28), center, rim, "1", center + V.At(0.5, 0.33));
        double top = s * Math.Sqrt(3) + 0.22;
        DoubleArrowWithTextBeside(d, d.Line("OrangeArrow", orange, 2.5), d.Text("OrangeMark", orange, 30), d.Pt(0, top), d.Pt(s, top), "?", V.At(s / 2, top + 0.33));
        d.Scene(-1.1, top + 0.5, 5.9, 0);
    }

    // Three rectangles: a tilted navy one; a pink one with its top right corner at the navy one's top
    // corner, the navy one's left corner on its left side and its bottom right corner on the navy one's
    // bottom side; a cyan one with its top left corner at that left corner, its top right corner on the
    // navy one's right side and its bottom on the navy one's bottom corner. Which is the largest? Each
    // is twice a triangle on a side of the navy one, half of it: all three are equal. Dragging the navy
    // one's top corner turns it, its bottom corner changes its width: the three stay equal.
    [Puzzle("three-overlapping-rectangles")]
    public static void ThreeOverlappingRectangles(Lgf d)
    {
        d.Texture = "FeltTip";
        var q = d.Pt(0, 0, "Q");
        var p = d.ArcSlider(q, 5, 18, 30, 26.6, (0, 0), "P");
        var k = d.Slider(new V("0.6 * P.Y", "-0.6 * P.X"), new V("0.92 * P.Y", "-0.92 * P.X"), (0.81 - 0.6) / 0.32, (1.813, -3.622), "K");
        var far = d.Pt(p + k, "Fr");
        string bottom = d.Num("K.Y + (P.X - K.X) / P.X * P.Y", "pb");
        string right = d.Num("P.X - P.Y / K.Y * K.X", "cr");
        var pinkFoot = d.Pt(new V("P.X", bottom), "PF");
        var cyanTop = d.Pt(new V(right, "0"), "CT");
        d.Polygon(d.Outline("Pink", "#FFEE1BA7", 3.5), d.Pt(new V("0", bottom)), pinkFoot, p, d.Pt(new V("0", "P.Y")));
        d.Polygon(d.Outline("Cyan", "#FF20E0F5", 3.5), d.Pt(new V("0", "K.Y")), d.Pt(new V(right, "K.Y")), cyanTop, q);
        d.Polygon(d.Outline("Navy", "#FF0604BA", 3.5), q, k, far, p);
        foreach (var corner in new[] { q, pinkFoot, cyanTop })
        {
            d.DotAt(corner);
        }

        d.Scene(-0.3, 2.6, 6.4, -4.0);
    }

    // Circles of diameters 2 and 3 side by side on a line, touching, and one of diameter 1 resting on
    // both. Lines from where the two big ones touch the line through where each touches the small one
    // meet at its top. The angle there? Nothing to drag: the diameters fix it.
    [Puzzle("three-circles-forming-a-triangle")]
    public static void ThreeCirclesFormingATriangle(Lgf d)
    {
        d.Texture = "FeltTip";
        double gap = Math.Sqrt(6);
        // the small circle's center: 1.5 from the middle one's center (0, 1), 2 from the large one's (√6, 1.5)
        double dx = gap, dy = 0.5, distance = Math.Sqrt(dx * dx + dy * dy);
        double along = (1.5 * 1.5 - 2 * 2 + distance * distance) / (2 * distance), across = Math.Sqrt(1.5 * 1.5 - along * along);
        double sx = along * dx / distance - across * dy / distance, sy = 1 + along * dy / distance + across * dx / distance;
        var footLeft = d.Pt(0, 0);
        var footRight = d.Pt(gap, 0);
        var top = d.Pt(sx, sy + 0.5, "T");
        d.Circle(d.Line("Medium", "#FF10C0FE", 3), d.Pt(0, 1), d.Pt(0, 2));
        d.Circle(d.Line("Large", "#FF0418F5", 3), d.Pt(gap, 1.5), d.Pt(gap, 3));
        d.Circle(d.Line("Small", "#FF2AE5FC", 3), d.Pt(sx, sy), top);
        string navyColor = "#FF0008BC";
        string navy = d.Line("Navy", navyColor, 4);
        d.Segment(navy, footLeft, top);
        d.Segment(navy, footRight, top);
        d.Angle(navy, top, footLeft, footRight, 30);
        d.Label("?", top + V.At(0.13, -0.55), d.Text("NavyMark", navyColor, 30));
        d.Segment(d.Line("Ink", Lgf.Ink, 4), d.Pt(-0.9, 0), d.Pt(gap + 0.9, 0));
        d.Scene(-1, 3, gap + 1.5, 0);
    }

    // Four squares of area 16 in a staircase, corner to corner, the last one stepping down beside the
    // third; the circle through the first one's bottom left corner and the last one's right corners.
    // Its area? Nothing to drag: the squares fix it.
    [Puzzle("four-squares-in-a-circle")]
    public static void FourSquaresInACircle(Lgf d)
    {
        double radius = Math.Sqrt(85);
        string green = d.Fill("Green", "#FFBEE85E", "#FF0A0843", 3, darkOutline: Lgf.DarkInk);
        string text = d.ShapeText("GreenMark", "#FF07A818", 30);
        foreach (var (x, y) in new[] { (0, 0), (4, 4), (8, 8), (12, 4) })
        {
            d.Box(green, V.At(x, y), V.At(x + 4, y + 4));
            d.Label("16", V.At(x + 2, y + 2), text);
        }

        d.Circle(d.Line("Circle", "#FF036D50", 3), d.Pt(7, 6), d.Pt(0, 0));
        d.Scene(7 - radius, 6 + radius, 7 + radius, 6 - radius);
    }

    // Three rectangles of the same area: a blue one 10 high lying on the bottom, a yellow one 10 wide
    // standing at its right end, and a red one 7 wide along the line from the blue one's top left
    // corner to the yellow one's top right corner. The area? Nothing to drag: equal areas make the
    // blue and the yellow equally long, and the red one's length the diagonal (here in tenths: 70 long).
    [Puzzle("three-equal-area-rectangles")]
    public static void ThreeEqualAreaRectangles(Lgf d)
    {
        d.Texture = "FeltTip";
        var o = d.Pt(0, 0);
        var blueTop = d.Pt(0, 1);
        var blueEnd = d.Pt(7, 1);
        var blueFoot = d.Pt(7, 0);
        var yellowFoot = d.Pt(8, 0);
        var yellowTop = d.Pt(8, 7);
        var yellowLeft = d.Pt(7, 7);
        var redTop = d.Pt(8 - 0.42, 7 + 0.56);
        var redLeft = d.Pt(-0.42, 1 + 0.56);
        d.Polygon(d.Fill("Blue", "#FF010D86", null), o, blueFoot, blueEnd, blueTop);
        d.Polygon(d.Fill("Yellow", "#FFFCDA01", null), blueFoot, yellowFoot, yellowTop, yellowLeft);
        d.Polygon(d.Fill("Red", "#FFE00002", null), blueTop, yellowTop, redTop, redLeft);
        // inked after coloring: the yellow one's corner shows through the red
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        d.Inked(frame, o, blueFoot, blueEnd, blueTop);
        d.Inked(frame, blueFoot, yellowFoot, yellowTop, yellowLeft);
        d.Inked(frame, blueTop, yellowTop, redTop, redLeft);
        string navy = "#FF0A1A9A";
        DoubleArrowWithTextBeside(d, d.Line("NavyArrow", navy, 2.5), d.Text("NavyMark", navy, 22), d.Pt(-0.32, 0), d.Pt(-0.32, 1), "10", V.At(-0.95, 0.5));
        string gold = "#FFE6B800";
        DoubleArrowWithTextBeside(d, d.Line("GoldArrow", gold, 2.5), d.Text("GoldMark", gold, 18), d.Pt(7, -0.28), d.Pt(8, -0.28), "10", V.At(7.5, -0.82));
        string red = "#FFD0105A";
        DoubleArrowWithTextBeside(d, d.Line("RedArrow", red, 2.5), d.Text("RedMark", red, 24), d.Pt(8 + 0.24, 7 + 0.18), d.Pt(8 + 0.24 - 0.42, 7 + 0.18 + 0.56), "7", V.At(8 + 0.62, 7 + 0.78));
        d.Scene(-1.3, 7.9, 8.8, -1.05);
    }

    // A square cut into four equal strips, its inscribed circle, and two more rectangles of the same
    // area at 30°, turned half round about the center from each other: each one's long sides run
    // through the circle's top and the left end of the top strip line, and through the circle's left
    // end. The shaded parts of the circle inside them? Nothing to drag: the strips fix it (each shaded
    // part lies between chords of 120° and 60°, a sixth of the circle).
    [Puzzle("six-equal-area-rectangles")]
    public static void SixEqualAreaRectangles(Lgf d)
    {
        d.Texture = "FeltTip";
        double r = Math.Sqrt(3), s = 2 * r, width = (S3 - 0.5) * r, half = s * s / 4 / width / 2;
        var o = d.Pt(0, 0, "O");
        var rim = d.Pt(0, r);
        string circle = d.Circle(null, o, rim);
        (double X, double Y) along = (S3, 0.5), normal = (-0.5, S3);
        V Corner(double sign, double t, double n) => d.Pt(sign * (t * along.X + n * normal.X), sign * (t * along.Y + n * normal.Y));
        var rectangles = new[] { 1.0, -1.0 }.Select(sign => new[] { Corner(sign, -half, 0.5 * r), Corner(sign, half, 0.5 * r), Corner(sign, half, S3 * r), Corner(sign, -half, S3 * r) }).ToArray();
        string yellow = d.Fill("Yellow", "#FFFFFF01", null);
        foreach (var corners in rectangles)
        {
            d.Region(yellow, "Intersection", circle, d.Polygon(null, corners));
        }

        string green = d.Line("Green", "#FF22B85A", 2);
        foreach (var y in new[] { -r / 2, 0, r / 2 })
        {
            d.Segment(green, d.Pt(-r, y), d.Pt(r, y));
        }

        d.Box(d.Outline("Square", "#FF22B85A", 2), V.At(-r, -r), V.At(r, r));
        d.Circle(d.Line("Circle", "#FF2EB82E", 3), o, rim);
        string gray = d.Outline("Gray", "#FF5E7066", 3, dark: "#FFB8C4BC");
        foreach (var corners in rectangles)
        {
            d.Inked(gray, corners);
        }

        d.Scene(-2.6, 2.75, 2.6, -2.75);
    }

    // Three equal squares: two side by side, tilted, a corner on the line, leaning on the top left
    // corner of the third, which stands on the line. A line from the tilted pair's top corner to the
    // third square's bottom right corner; the angle where it crosses that square's top? Nothing to
    // drag: leaning on the corner tilts the pair by 30°.
    [Puzzle("angle-formed-by-three-squares")]
    public static void AngleFormedByThreeSquares(Lgf d)
    {
        double u = Math.Sqrt(3);
        var b = d.Pt(0, 0);
        var t = d.Pt(u, 1, "T");
        var r = d.Pt(u - 0.5, 1 + S3, "R");
        var l = d.Pt(-0.5, S3);
        var foot = d.Pt(u + 1, 0, "F");
        var top = d.Pt(u + 1, 1);
        var corner = d.Pt(u, 0);
        d.Segment(d.Line("Base", "#FFA6DA10", 4, dark: "#FFC8F060"), d.Pt(-0.45, 0), d.Pt(u + 1.15, 0));
        // the green stays bright in the dark theme: a deep green outline there parts the squares
        string green = d.Fill("Green", "#FF85EBA4", "#FF02D86A", 3, darkOutline: "#FF0B5A30");
        d.Polygon(green, b, t, r, l);
        d.Polygon(green, corner, foot, top, t);
        d.Segment(d.Line("GreenLine", "#FF02D86A", 3, dark: "#FF0B5A30"), Mid(d, b, t), Mid(d, l, r));
        var x = d.Crossing(r, foot, t, top, "X");
        string purpleColor = "#FF5A20C4";
        string purple = d.Line("Purple", purpleColor, 3.5);
        d.Segment(purple, r, x);
        // inside the green square her purple over the green came out a deep teal
        d.Segment(d.Line("Teal", "#FF1F6F70", 3.5, dark: "#FF0F3F40"), x, foot);
        d.Angle(purple, x, top, r, 30);
        d.Label("?", x + V.At(0.14, 0.36), d.Text("PurpleMark", purpleColor, 30));
        d.Scene(-0.5, 1 + S3, u + 1.15, 0);
    }
}
