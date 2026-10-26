#nullable disable

static partial class Puzzles
{
    // Two isosceles triangles on the same base line: the yellow one right-angled on its base, the
    // green one with legs as long as that base, the blue line its third side. Same area: the angle x?
    // Nothing to drag: equal areas fix the green one's apex angle (150 degrees).
    [Puzzle("two-isosceles-triangles")]
    public static void TwoIsoscelesTriangles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double b = 4;
        var a = d.Pt(0, 0, "A");
        var bb = d.Pt(b, 0, "B");
        var y = d.Pt(b / 2, b / 2, "Y");
        var c = d.Pt(b + b * S3, b / 2, "C");
        string yellow = d.Line("Yellow", "#FFE8C01E", 4);
        d.Segment(yellow, a, y, "OneTick");
        d.Segment(yellow, y, bb, "OneTick");
        d.Angle(d.Line("YellowMark", "#FFE8C01E", 2.5), y, bb, a, 18);
        string green = d.Line("Green", "#FF7DB060", 4);
        d.Segment(green, a, bb, "TwoTicks");
        d.Segment(green, bb, c, "TwoTicks");
        string blue = "#FF3D7AA6";
        d.Segment(d.Line("Blue", blue, 4), a, c);
        d.ArcMark(d.Line("BlueMark", blue, 3), c, a, bb, 1.1);
        d.Label("x", V.At(5.85, 1.3), d.Text("BlueText", "#FF2E6A80", 24), 24);
        d.Scene(-0.2, 2.2, b + b * S3 + 0.2, -0.2);
    }

    // A circle, four points on it and one on the bottom chord: the two pink triangles, the big one on
    // the bottom chord and the one under the top chord are all isosceles. The angle at the top right?
    // Nothing to drag: that fixes the angles (the four points are four corners of a regular pentagon).
    [Puzzle("four-isosceles-triangles-in-a-circle")]
    public static void FourIsoscelesTrianglesInACircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 5;
        V OnCircle(double degrees, string name) => d.Pt(R * Math.Cos(degrees * Math.PI / 180), R * Math.Sin(degrees * Math.PI / 180), name);
        var o = d.Pt(0, 0, "O");
        var dd = OnCircle(198, "D");
        var a = OnCircle(126, "A");
        var b = OnCircle(54, "B");
        var c = OnCircle(342, "C");
        double ax = R * Math.Cos(126 * Math.PI / 180), dx = R * Math.Cos(198 * Math.PI / 180);
        var e = d.Pt(2 * ax - dx, R * Math.Sin(198 * Math.PI / 180), "E");
        d.Circle(d.Outline("Circle", "#FF5F7878", 4), o, c);
        string ink = d.Line("Ink", Lgf.Ink, 4);
        d.Segment(ink, dd, e);
        d.Segment(ink, a, b);
        d.Segment(ink, b, c);
        d.Segment(ink, a, c);
        string pink = d.Line("Pink", "#FFC42A8E", 4);
        d.Segment(pink, a, dd, "TwoTicks");
        d.Segment(pink, a, e, "TwoTicks");
        d.Segment(pink, e, c, "TwoTicks");
        string orange = "#FFF07A45";
        d.ArcMark(d.Line("OrangeMark", orange, 3), b, a, c, 0.55);
        d.Label("?", V.At(2.25, 3.45),d.Text("OrangeText", orange, 26));
        d.Scene(-R, R, R, -R);
    }

    // Two concentric regular octagons, the inner one upright. The blue square has two opposite corners
    // on corners of the inner octagon and a third on the outer one; the red square is drawn around the
    // outer octagon. How do the squares compare? Nothing to drag: the blue square's corner fixes the
    // outer octagon.
    [Puzzle("two-octagons-and-two-squares")]
    public static void TwoOctagonsAndTwoSquares(Lgf d)
    {
        const double r = 4;
        double rad = Math.PI / 180;
        V[] Ring(double radius, double from) => Enumerable.Range(0, 8)
            .Select(i => d.Pt(radius * Math.Cos((from + 45 * i) * rad), radius * Math.Sin((from + 45 * i) * rad))).ToArray();
        var inner = Ring(r, 22.5);
        double k = r * Math.Cos(22.5 * rad);
        double outer = k * Math.Sqrt(2);
        var big = Ring(outer, 0);
        double apothem = outer * Math.Cos(22.5 * rad);
        var red = Enumerable.Range(0, 4)
            .Select(i => d.Pt(apothem * Math.Sqrt(2) * Math.Cos((67.5 + 90 * i) * rad), apothem * Math.Sqrt(2) * Math.Sin((67.5 + 90 * i) * rad))).ToArray();
        string ink = d.Outline("Octagon", Lgf.Ink, 3.5);
        d.Polygon(ink, inner);
        d.Polygon(ink, big);
        d.Polygon(d.Outline("Red", "#FFE5502C", 3), red);
        d.Box(d.Outline("Blue", "#FF2A4A9A", 3.5), V.At(-k, -k), V.At(r * Math.Sin(22.5 * rad), r * Math.Sin(22.5 * rad)));
        double edge = apothem * Math.Sqrt(2) * Math.Cos(22.5 * rad);
        d.Scene(-edge, edge, edge, -edge);
    }

    // A regular hexagon, the star of its two big triangles, and a small regular hexagon parallel to the
    // big one with its corners on the star's sides (the middles of the sides of the star's own inner
    // hexagon). The star outside the small hexagon is shaded; the small one has area 18: the shaded area?
    // Nothing to drag: the parallel sides fix the small hexagon (half the big one's size).
    [Puzzle("hexagon-in-a-star-in-a-hexagon")]
    public static void HexagonInAStarInAHexagon(Lgf d)
    {
        const double R = 4;
        V[] Corners(double radius, params int[] at) => at
            .Select(i => d.Pt(radius * Math.Cos((90 + 60 * i) * Math.PI / 180), radius * Math.Sin((90 + 60 * i) * Math.PI / 180))).ToArray();
        var big = Corners(R, 0, 1, 2, 3, 4, 5);
        var small = Corners(R / 2, 0, 1, 2, 3, 4, 5);
        string up = d.Polygon(null, big[0], big[2], big[4]);
        string down = d.Polygon(null, big[1], big[3], big[5]);
        string hole = d.Polygon(null, small);
        d.Region(d.Fill("Yellow", "#FFFEFDCF", null), "Difference", d.Region(null, "Union", up, down), hole);
        // the star's sides cross the yellow, which stays bright in the dark theme: deep there
        string yellow = d.Outline("YellowLine", "#FFFEF530", 3, dark: "#FF8C7A00");
        d.Polygon(yellow, big[0], big[2], big[4]);
        d.Polygon(yellow, big[1], big[3], big[5]);
        d.Polygon(d.Outline("Orange", "#FFEC8A48", 2.5, dark: "#FFE07030"), small);
        d.Polygon(d.Outline("Blue", "#FF0505A8", 3.5), big);
        d.Scene(-R * S3, R, R * S3, -R);
    }

    // Three regular hexagons on a shelf: the small one (area 10) and a big one standing side by side,
    // touching, the other big one lying on the small one, its left corner on the small one's and its
    // lower right side along the standing one's upper left side. The red triangle's area? Nothing to
    // drag: that fixes the big ones at one and a half times the small one's side.
    [Puzzle("stacked-hexagons")]
    public static void StackedHexagons(Lgf d)
    {
        const double s = 2, S = 3;
        var small = HexagonFromLeft(d, V.At(-s, s * S3), Lgf.N(s));
        var right = HexagonFromLeft(d, V.At(1.5 * s - S / 2, S * S3), Lgf.N(S));
        var top = HexagonFromLeft(d, V.At(-s / 2 - S / 2, 2 * s * S3 + S * S3), Lgf.N(S));
        d.Polygon(d.Fill("Navy", "#FFA8BCDC", "#FF023190", 3), small);
        d.Polygon(d.Fill("Red", "#FFE0868F", "#FFE0484A", 2), top[1], top[0], right[3]);
        string cyan = d.Outline("Cyan", "#FF0B82BA", 3);
        d.Polygon(cyan, top);
        d.Polygon(cyan, right);
        string shelf = d.Line("Shelf", "#FF023851", 3);
        var a = d.Pt(-2.85, 0);
        var b = d.Pt(7.95, 0);
        d.Segment(shelf, a, b);
        d.Segment(shelf, a, d.Pt(-3.0, -0.55));
        d.Segment(shelf, b, d.Pt(8.15, -0.55));
        d.Label("10", V.At(0, s * S3), d.ShapeText("NavyMark", "#FF023190", 30), 30);
        d.Scene(-3.0, 2 * s * S3 + 2 * S * S3, 8.15, -0.55);
    }

    // A square of area 100 with its corners on the sides of a big square, the corner triangle under it
    // of area 24, and a circle in the opposite corner touching the square. The circle's area? Nothing
    // to drag: the areas fix the triangle (legs 6 and 8), so the big square (14) and the circle (radius 2).
    [Puzzle("a-square-triangle-and-circle-in-a-square")]
    public static void ASquareTriangleAndCircleInASquare(Lgf d)
    {
        const double L = 14, a = 6;
        var bottom = d.Pt(a, 0);
        var right = d.Pt(L, a);
        var top = d.Pt(L - a, L);
        var left = d.Pt(0, L - a);
        d.Box(d.PaperFill(), V.At(0, 0), V.At(L, L));
        d.Polygon(d.Fill("Peach", "#FFF6CDB0"), d.Pt(0, 0), bottom, left);
        d.Polygon(d.Fill("Yellow", "#FFF9EFA6"), bottom, right, top, left);
        var c = d.Pt(L - 2, 2);
        d.Circle(d.Fill("Pink", "#FFEE98B4"), c, d.Pt(L, 2));
        d.Label("100", V.At(L / 2, L / 2), d.ShapeText("GoldMark", "#FFD4B820", 30), 30);
        d.Label("24", V.At(1.6, 1.9), d.ShapeText("OrangeMark", "#FFE07A50", 28), 28);
        d.Label("?", c, d.ShapeText("RedMark", "#FFB0182E", 28), 28);
        // room under the circle for the page's buttons in the bottom right corner
        d.Scene(0, L, L, -1.4);
    }

    // A regular octagon, a line from a top corner straight down to a bottom corner and a diagonal from
    // the other top corner; the two triangles between them are shaded. Their area equals the octagon's
    // perimeter: what is it? Nothing to drag: the octagon is fixed up to its size, which is the answer.
    [Puzzle("region-of-an-octagon")]
    public static void RegionOfAnOctagon(Lgf d)
    {
        d.Texture = "FeltTip";
        const double a = 4;
        double w = a / 2 + a / Math.Sqrt(2);
        var octagon = new[] { (-a / 2, w), (-w, a / 2), (-w, -a / 2), (-a / 2, -w), (a / 2, -w), (w, -a / 2), (w, a / 2), (a / 2, w) }
            .Select(p => d.Pt(p.Item1, p.Item2)).ToArray();
        var topLeft = octagon[0];
        var topRight = octagon[7];
        var right = octagon[5];
        var bottom = octagon[4];
        var x = d.Pt(a / 2, w - a);
        string magenta = d.Fill("Magenta", "#FFD801B8", null);
        d.Polygon(magenta, topLeft, topRight, x);
        d.Polygon(magenta, x, right, bottom);
        string line = "#FFB0068C";
        d.Polygon(d.Outline("Frame", line, 3.5), octagon);
        string ink = d.Line("Line", line, 3.5);
        d.Segment(ink, topRight, bottom);
        d.Segment(ink, topLeft, right);
        d.Scene(-w, w, w, -w);
    }

    // A square and a semicircle with its diameter's ends on the bottom and right sides, its arc touching
    // the top and the left sides. More pink or blue? Nothing to drag: touching both sides fixes it (the
    // diameter runs parallel to the diagonal; her bottom end is a little further right).
    [Puzzle("semi-circle-in-a-square")]
    public static void SemiCircleInASquare(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 6;
        double r = L * (2 - Math.Sqrt(2)), h = r / Math.Sqrt(2);
        var c = d.Pt(r, L - r);
        var low = d.Pt(r - h, 0);
        var high = d.Pt(L, L - r + h);
        d.Box(d.Fill("Pink", "#FFFD8E92", Lgf.Ink, 4), V.At(0, 0), V.At(L, L));
        d.Arc("CircleSegment", d.Fill("Blue", "#FF6AA0D6", Lgf.Ink, 4), c, high, low);
        d.Scene(0, L, L, -0.5);
    }

    // An equilateral triangle and a regular hexagon of the same perimeter (side 6 and 3), each with its
    // incircle. The ratio of the circles' areas? Nothing to drag: only the size is free.
    [Puzzle("polygons-with-equal-perimeter")]
    public static void PolygonsWithEqualPerimeter(Lgf d)
    {
        const double t = 6, h = 3, left = 7.45;
        var a = d.Pt(0, 0);
        var b = d.Pt(t, 0);
        var apex = d.Pt(t / 2, t * S3);
        var hexagon = HexagonFromLeft(d, V.At(left, h * S3), Lgf.N(h));
        var inTriangle = d.Pt(t / 2, t * S3 / 3);
        var inHexagon = d.Pt(left + h, h * S3);
        d.Circle(d.Fill("Green", "#FFCFE6DC", "#FF0E9060", 3), inTriangle, d.Pt(t / 2, 0));
        d.Circle(d.Fill("Lime", "#FFEEF6C4", "#FF8CDC5A", 3), inHexagon, d.Pt(left + h, 0));
        string ink = d.Outline("Frame", "#FF253B50", 2.5);
        d.Inked(ink, a, b, apex);
        d.Inked(ink, hexagon);
        d.Scene(0, t * S3, left + 2 * h, 0);
    }

    // Shapes stacked on a shelf: two congruent 2:1 rectangles, one lying, one leaning from its top
    // corner down to the shelf against a congruent equilateral triangle (so tilted 30 degrees); a
    // semicircle (area 4) under the leaning one, the pink circle on the lying one between the leaning
    // one and the other triangle. The circle's area? Nothing to drag: the shapes fix each other.
    [Puzzle("stacked-shapes")]
    public static void StackedShapes(Lgf d)
    {
        const double w = 2;
        double s3 = Math.Sqrt(3), rho = w / s3, r = w * S3;
        var a = d.Pt(2 * w, w);
        var b = d.Pt(2 * w + w / 2, w + w * S3);
        var dd = d.Pt(2 * w + w * s3, 0);
        var c = d.Pt(2 * w + w * s3 + w / 2, w * S3);
        string blue = d.Fill("Blue", "#FFE6ECFB", "#FF102190", 3);
        d.Box(blue, V.At(0, 0), V.At(2 * w, w));
        d.Polygon(blue, a, b, c, dd);
        string green = d.Fill("Green", "#FFF0F8D2", "#FF86C060", 3);
        d.Polygon(green, d.Pt(0, w), d.Pt(w, w), d.Pt(w / 2, w + w * S3));
        d.Polygon(green, dd, d.Pt(2 * w + w * s3 + w, 0), c);
        var m = d.Pt(2 * w + rho, 0);
        d.Arc("CircleSegment", d.Fill("Purple", "#FFE4E0F8", "#FF7A6AC4", 3), m, d.Pt(2 * w + 2 * rho, 0), d.Pt(2 * w, 0));
        var o = d.Pt(1.5 * w, w + r);
        d.Circle(d.Fill("Pink", "#FFF7E6F3", "#FFA1008A", 3), o, d.Pt(1.5 * w, w));
        d.Segment(d.Line("Shelf", "#FF057CB3", 3), d.Pt(-0.75, 0), d.Pt(10.6, 0));
        d.Label("?", o, d.ShapeText("MagentaMark", "#FFA1008A", 30), 30);
        d.Label("4", V.At(2 * w + rho, rho * 0.45), d.ShapeText("PurpleMark", "#FF6A58B8", 26));
        d.Scene(-0.75, w + 2 * r, 10.6, -0.3);
    }

    // Four squares around a corner (areas 100, 25, 16 and the red one) and a circle through four of
    // their corners. The circle's area? Nothing to drag: the crossing chords fix the red square (side 8)
    // and so the circle. Drawn on a computer.
    [Puzzle("four-squares-overlapping-a-circle")]
    public static void FourSquaresOverlappingACircle(Lgf d)
    {
        d.Texture = null;
        d.Box(d.Fill("Red", "#FFFBEBEB", "#FFC83A36", 2.5), V.At(-8, 0), V.At(0, 8));
        d.Box(d.Fill("Blue", "#FFE8F1FA", "#FF2A5DA8", 2.5), V.At(-10, -10), V.At(0, 0));
        d.Box(d.Fill("Orange", "#FFFAF0E7", "#FFD9622B", 2.5), V.At(0, 0), V.At(4, 4));
        d.Box(d.Fill("Purple", "#FFF1EFFC", "#FF7467BF", 2.5), V.At(0, -5), V.At(5, 0));
        var o = d.Pt(-1.5, -3, "O");
        var east = d.Pt(5, 0);
        d.Circle(d.Outline("Circle", "#FF808080", 2.5), o, east);
        foreach (var (x, y) in new[] { (0.0, 4.0), (5.0, 0.0), (0.0, -10.0), (-8.0, 0.0) })
        {
            d.DotAt(V.At(x, y));
        }

        d.Label("100", V.At(-5.2, -4.8), d.ShapeText("BlueMark", "#FF1A1AB0", 24), 24);
        d.Label("16", V.At(2.1, 1.85), d.ShapeText("OrangeMark", "#FFE07B00", 24), 24);
        d.Label("25", V.At(2.5, -2.7), d.ShapeText("PurpleMark", "#FF8000D0", 24), 24);
        double r = Math.Sqrt(51.25);
        d.Scene(-10, 8, -1.5 + r, -3 - r);
    }

    // A triangle in a semicircle (on its diameter, the apex on the arc), its incircle, and the triangle
    // of the points where the circle touches its sides. The angle at the point on the diameter?
    // Dragging the apex along the arc changes everything else: the angle stays 45 degrees.
    [Puzzle("triangle-in-circle-in-triangle-in-semi-circle")]
    public static void TriangleInCircleInTriangleInSemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 5;
        var o = d.Pt(0, 0, "O");
        var a = d.Pt(-R, 0, "A");
        var b = d.Pt(R, 0, "B");
        d.Arc("CircleArc", d.Line("Semicircle", "#FF3AA6C8", 4), o, b, a);
        var t = d.ArcSlider(o, R, 32, 148, 72, (0, 0), "T");
        string toB = d.Num("distance(T, B)", "ta");
        string toA = d.Num("distance(T, A)", "tb");
        string sum = d.Num($"{toB} + {toA} + {2 * R}", "p");
        var i = d.Pt(new V($"({R} * ({toA} - {toB}) + {2 * R} * T.X) / {sum}", $"{2 * R} * T.Y / {sum}"), "I");
        var p = d.Pt(new V("I.X", "0"), "P");
        // the touching points: as far from each corner as the half perimeter less the opposite side
        string alongA = d.Num($"({toA} + {2 * R} - {toB}) / (2 * {toA})", "ka");
        string alongB = d.Num($"({toB} + {2 * R} - {toA}) / (2 * {toB})", "kb");
        var onA = d.Pt(a + alongA * (t - a), "Q");
        var onB = d.Pt(b + alongB * (t - b), "S");
        string ink = d.Line("Ink", Lgf.Ink, 3.5);
        d.Segment(ink, a, b);
        d.Segment(ink, a, t);
        d.Segment(ink, b, t);
        d.Circle(d.Outline("Circle", "#FF0E7BAD", 4), i, p);
        string red = d.Line("Red", "#FFF0703F", 3);
        d.Segment(red, p, onA);
        d.Segment(red, p, onB);
        d.Segment(red, onA, onB);
        string orange = "#FFE89A60";
        d.ArcMark(d.Line("OrangeMark", orange, 2.5), p, onB, onA, 0.42);
        d.Label("?", new V("I.X", "0.55 * I.Y"), d.Text("OrangeText", orange, 24), 24);
        d.Scene(-R, R, R, 0);
    }
}
