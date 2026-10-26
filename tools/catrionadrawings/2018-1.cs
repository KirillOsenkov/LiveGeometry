#nullable disable

static partial class Puzzles
{
    // Two congruent rectangles in a semicircle, one lying and one standing side by side on the
    // diameter, the lying one's top left corner and the standing one's top right corner on the arc;
    // red lines from those corners to a point of the arc between them. The angle there? The two
    // corners are always a quarter turn apart (the rectangles meet at the center), so the angle is
    // 135 degrees. Dragging the lying one's corner along the arc changes the rectangles' shape, and
    // the point on the arc can be dragged between the corners: the angle stays.
    [Puzzle("two-rectangles-in-a-semi-circle")]
    public static void TwoRectanglesInASemiCircle(Lgf d)
    {
        const double R = 5;
        string red = "#FFD0101C";
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0, "E");
        var west = d.Pt(-R, 0, "W");
        var p = d.ArcSlider(o, R, 122, 160, 145, (0, 0), "P");
        var q = d.Pt(new V("P.Y", "-P.X"), "Q");
        var v = d.ArcSlider(o, R, 76, 118, 105, (0, 0), "Vx");
        string green = d.Fill("Green", "#FFA8EE48", "#FF6CC818");
        d.Box(green, new V("P.X", "0"), new V("0", "P.Y"));
        d.Box(green, V.At(0, 0), new V("P.Y", "-P.X"));
        d.Arc("CircleSegment", d.Outline("Frame", "#FF1018E8", 3.5), o, east, west);
        string line = d.Line("Red", red, 3.5);
        d.Segment(line, p, v);
        d.Segment(line, v, q);
        d.Angle(line, v, p, q, 20);
        d.Label("?", v + V.At(0.1, -1), d.Text("RedMark", red, 30), 30);
        d.Scene(-R, R, R, 0);
    }

    // A regular hexagon cut by two purple lines crossing on its long diagonal, each from a corner to
    // the far side, making two kites: the blue one at the bottom left corner (9) and the pink one at
    // the top right corner (1). The yellow area? Nothing to drag: the two areas fix where the lines
    // cross (pink / blue = (1 - s)^2 / (sqrt(3) (1 + s)^2) with s the crossing's place along the
    // half-diagonal, s = 0.39), so it is a little lower than in her picture.
    [Puzzle("divided-hexagon-iii")]
    public static void DividedHexagonIII(Lgf d)
    {
        const double A = 4;
        double r = Math.Pow(3, 0.25) / 3, s = (1 - r) / (1 + r), u = (1 - s) / (1 + s);
        double h = A * 2 * S3;
        var right = d.Pt(A, 0, "R");
        var b = d.Pt(A / 2, h / 2, "B");
        var a = d.Pt(-A / 2, h / 2, "Ta");
        var left = d.Pt(-A, 0, "L");
        var bottomLeft = d.Pt(-A / 2, -h / 2, "D");
        var c = d.Pt(A / 2, -h / 2, "C");
        var x = d.Pt(s * A / 2, s * h / 2, "X");
        var side = d.Pt(A / 2 + u * A / 2, h / 2 - u * h / 2, "Sr");
        var top = d.Pt(A / 2 - u * A, h / 2, "Tt");
        d.Polygon(d.Fill("Yellow", "#FFFDE58F", null), left, a, top, x);
        d.Polygon(d.Fill("Pink", "#FFF08AB4", null), x, top, b, side);
        d.Polygon(d.Fill("Blue", "#FF9CEDFA", null), left, x, c, bottomLeft);
        d.Polygon(d.Outline("Frame", Lgf.Ink, 3), right, b, a, left, bottomLeft, c);
        string purple = d.Line("Purple", "#FF1A10B0", 3);
        d.Segment(purple, left, side);
        d.Segment(purple, top, c);
        string navy = d.ShapeText("NavyMark", "#FF0A10A8", 32);
        d.Label("?", Mid(d, left, a, top, x), d.ShapeText("PurpleMark", "#FF6E2860", 32), 32);
        d.Label("1", Mid(d, x, top, b, side), navy, 32);
        d.Label("9", Mid(d, left, x, c, bottomLeft), navy, 32);
        d.Scene(-A, h / 2, A, -h / 2);
    }

    // A 3-4-5 right triangle and the circles on its three sides as diameters; shaded is what lies in
    // an odd number of the four shapes (the three disks and the triangle). The total? Nothing to
    // drag: the sides are given. Every disk part is a flat Difference region: the big disk without
    // the two small ones and the triangle, the two lunes outside it, and the triangle without each
    // small disk (the triangle is covered by the two small disks, so the parts in just one of them
    // are these).
    [Puzzle("circles-around-a-triangle")]
    public static void CirclesAroundATriangle(Lgf d)
    {
        var a = d.Pt(0, 0, "A");
        var b = d.Pt(4, 0, "B");
        var c = d.Pt(4, 3, "C");
        var big = d.Pt(2, 1.5, "Mc");
        var bottom = d.Pt(2, 0, "Mb");
        var side = d.Pt(4, 1.5, "Ms");
        string bigDisk = d.Circle(null, big, a);
        string bottomDisk = d.Circle(null, bottom, a);
        string sideDisk = d.Circle(null, side, b);
        string triangle = d.Polygon(null, a, b, c);
        string blue = d.Fill("Blue", "#FF98D4F0", null);
        d.Region(blue, "Difference", bigDisk, bottomDisk, sideDisk, triangle);
        d.Region(blue, "Difference", bottomDisk, bigDisk);
        d.Region(blue, "Difference", sideDisk, bigDisk);
        d.Region(blue, "Difference", triangle, sideDisk);
        d.Region(blue, "Difference", triangle, bottomDisk);
        d.Circle(d.Outline("YellowGreen", "#FF9CC23A", 3), big, a);
        d.Circle(d.Outline("Green", "#FF3EB42C", 3), bottom, a);
        d.Circle(d.Outline("BlueRim", "#FF1C8AE0", 3), side, b);
        d.Inked(d.Outline("Frame", Lgf.Ink, 4), a, b, c);
        d.Scene(-0.5, 4, 5.5, -2);
    }

    // Three identical rectangles stacked in a semicircle of diameter 10: two side by side on the
    // diameter with their outer top corners on the arc, the third on top of them, centered, with its
    // top corners on the arc. The total area? Nothing to drag: the corners on the arc fix the
    // rectangles (twice as wide as high, sqrt(5) high).
    [Puzzle("rectangles-stacked-in-a-semi-circle")]
    public static void RectanglesStackedInASemiCircle(Lgf d)
    {
        const double R = 5;
        double h = Math.Sqrt(5), w = 2 * h;
        string green = "#FF9CC21E";
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0, "E");
        var west = d.Pt(-R, 0, "W");
        string blue = d.Fill("Blue", "#FFD4F2FC", "#FF2EBFC0", 3.5);
        d.Box(blue, V.At(-w, 0), V.At(0, h));
        d.Box(blue, V.At(0, 0), V.At(w, h));
        d.Box(blue, V.At(-w / 2, h), V.At(w / 2, 2 * h));
        d.Arc("CircleSegment", d.Outline("Frame", green, 3.5), o, east, west);
        d.Dimension(d.Line("GreenArrow", green, 2.5), d.Text("GreenMark", green, 28), east, west, 0.7, "10", 0.5);
        d.Scene(-R, R, R, -1);
    }

    // A square with two circles of radius 4 in opposite corners, touching each other, and a pink
    // square in a third corner whose sides touch the circles (its bottom the lower one's top, its
    // right side the upper one's left point). Its area? Nothing to drag: the touching circles fix
    // the big square (8 + 4 sqrt(2)) and with it the pink one (4 sqrt(2)).
    [Puzzle("two-circles-two-squares")]
    public static void TwoCirclesTwoSquares(Lgf d)
    {
        double s = 4 * Math.Sqrt(2), L = 8 + s;
        string orange = "#FFF2A41E";
        d.Box(d.Fill("Pink", "#FFF4C8D2", "#FFE8A8B4"), V.At(0, L - s), V.At(s, L));
        d.Box(d.Outline("Frame", Lgf.Ink, 3), V.At(0, 0), V.At(L, L));
        var low = d.Pt(4, 4, "C1");
        var high = d.Pt(L - 4, L - 4, "C2");
        var lowRim = d.Pt(8, 4);
        var highRim = d.Pt(L, L - 4);
        string rim = d.Outline("Orange", orange, 3);
        d.Circle(rim, low, lowRim);
        d.Circle(rim, high, highRim);
        string arrow = d.Line("Arrow", Lgf.Ink, 3);
        d.Vector(arrow, low, lowRim);
        d.Vector(arrow, high, highRim);
        d.Label("4", V.At(6, 4.75), d.Text("Mark", Lgf.Ink, 30), 30);
        d.Label("4", V.At(L - 2, L - 3.25), d.Text("Mark", Lgf.Ink, 30), 30);
        d.Label("?", V.At(s / 2, L - s / 2), d.ShapeText("ShapeMark", Lgf.Ink, 34), 34);
        d.Scene(0, L, L, 0);
    }

    // A square cut by its diagonals and four lines from the left corners: the top one to the bottom
    // midpoint and the right midpoint, the bottom one to the top midpoint and the right midpoint. What
    // fraction is yellow? Nothing to drag: corners and midpoints fix it. The yellow parts are the
    // seven faces she colored, worked out as the crossings of those lines (side 4).
    [Puzzle("subdivided-square")]
    public static void SubdividedSquare(Lgf d)
    {
        d.Texture = "FeltTip";
        var bl = d.Pt(0, 0, "BL");
        var br = d.Pt(4, 0, "BR");
        var tr = d.Pt(4, 4, "TR");
        var tl = d.Pt(0, 4, "TL");
        var top = d.Pt(2, 4, "Tm");
        var right = d.Pt(4, 2, "Rm");
        var bottom = d.Pt(2, 0, "Bm");
        var center = d.Pt(2, 2, "Ct");
        var left = d.Pt(1, 2);
        var upper = d.Pt(1.6, 3.2);
        var lower = d.Pt(1.6, 0.8);
        var upperDiagonal = d.Pt(4.0 / 3, 8.0 / 3);
        var lowerDiagonal = d.Pt(4.0 / 3, 4.0 / 3);
        var upperRight = d.Pt(8.0 / 3, 8.0 / 3);
        var lowerRight = d.Pt(8.0 / 3, 4.0 / 3);
        string yellow = d.Fill("Yellow", "#FFFEF202", null);
        d.Polygon(yellow, tl, bl, left);
        d.Polygon(yellow, top, tr, upperRight, upper);
        d.Polygon(yellow, bottom, br, lowerRight, lower);
        d.Polygon(yellow, tl, upper, upperDiagonal);
        d.Polygon(yellow, bl, lower, lowerDiagonal);
        d.Polygon(yellow, upperDiagonal, center, lowerDiagonal, left);
        d.Polygon(yellow, right, upperRight, center, lowerRight);
        // across the yellow (bright in the dark theme) and the paper: a vivid red there
        string red = d.Line("Red", "#FFCC0835", 7, dark: "#FFF0405F");
        d.Segment(red, tl, br);
        d.Segment(red, bl, tr);
        d.Segment(red, tl, right);
        d.Segment(red, bl, right);
        d.Segment(red, tl, bottom);
        d.Segment(red, bl, top);
        d.Inked(d.Outline("Frame", "#FFCC0835", 7, dark: "#FFF0405F"), bl, br, tr, tl);
        d.Scene(0, 4, 4, 0);
    }

    // A triangle and a circle through two of its corners, crossing its other two sides; the chords
    // from those corners to the crossings, and between the crossings, cut it into five parts: 15
    // (blue), 3 (yellow), 7 (green) and two pink ones. The total? The yellow triangle is the blue one
    // shrunk by the ratio the green one is the whole triangle shrunk by (both 1 / sqrt(5)), so the
    // drawing keeps that ratio: the crossings are where the circle through the corners meets the
    // sides when the green triangle is the whole one scaled by 1 / sqrt(5). Dragging the far corner
    // changes the triangle and the circle, the ratios stay. Her numbers can't all hold at once (the
    // pink parts would multiply to 15 x 3 = 45 and add up to 10), so the parts' areas are only in
    // the ratios of 15, 3 and 7 roughly, as in her picture.
    [Puzzle("chords-dividing-a-triangle-into-five")]
    public static void ChordsDividingATriangleIntoFive(Lgf d)
    {
        const double r = 0.4472135955;
        var a = d.Pt(0, 0, "A");
        var b = d.Pt(0.84, -5.04, "B");
        string room = d.Circle(null, d.Pt(7.3, -2.26), d.Pt(8.3, -2.26));
        var z = d.HandleIn(room, 7.3, -2.26, "Z");
        string za = d.Num("distance(Z, A)", "za");
        string zb = d.Num("distance(Z, B)", "zb");
        var p = d.Pt(z + d.Num($"{Lgf.N(r)} * {zb} / {za}") * (a - z), "P");
        var q = d.Pt(z + d.Num($"{Lgf.N(r)} * {za} / {zb}") * (b - z), "Q");
        var x = d.Crossing(a, q, b, p, "X");
        // the circle through A (the origin), B and P
        string pp = d.Num("P.X ^ 2 + P.Y ^ 2", "pp");
        string bb = d.Num("B.X ^ 2 + B.Y ^ 2", "bb");
        string den = d.Num("2 * (B.X * P.Y - P.X * B.Y)", "den");
        var o = d.Pt(new V($"(P.Y * {bb} - B.Y * {pp}) / {den}", $"(B.X * {pp} - P.X * {bb}) / {den}"), "O");
        d.Polygon(d.Fill("Blue", "#FFA5F2FA", null), a, b, x);
        d.Polygon(d.Fill("Violet", "#FFF8B8EA", null), a, x, p);
        d.Polygon(d.Fill("Pink", "#FFFCAAD4", null), b, x, q);
        d.Polygon(d.Fill("Yellow", "#FFFCFC65", null), x, p, q);
        d.Polygon(d.Fill("Green", "#FFE4FE82", null), p, q, z);
        d.Circle(d.Outline("Circle", "#FF04A004", 3), o, a);
        string blue = d.Line("Chord", "#FF0890E4", 3);
        d.Segment(blue, a, q);
        d.Segment(blue, b, p);
        d.Segment(blue, p, q);
        d.Inked(d.Outline("Frame", Lgf.Ink, 3.5), a, b, z);
        string mark = d.ShapeText("ShapeMark", Lgf.Ink, 28);
        d.Label("15", Mid(d, a, b, x), mark, 28);
        d.Label("3", Mid(d, x, p, q), mark, 28);
        d.Label("7", Mid(d, p, q, z), mark, 28);
        d.Scene(-1.3, 0.6, 8.4, -5.3);
    }

    // Three congruent rectangles, each with a long side's midpoint at one point and turned a third of a
    // turn from each other, a short side of each crossing a short side of the next at both midpoints
    // (the dots). The area covered? Nothing to drag: the midpoints meeting makes the rectangles sqrt(3)
    // times as long as wide. The overlaps in the colors she got by laying one pencil over the other.
    [Puzzle("three-rectangles-as-a-triangle")]
    public static void ThreeRectanglesAsATriangle(Lgf d)
    {
        const double b = 2, a = 1.7320508075688772;
        (double X, double Y) Turn((double X, double Y) p, double degrees)
        {
            double t = degrees * Math.PI / 180, c = Math.Cos(t), s = Math.Sin(t);
            return (c * p.X - s * p.Y, s * p.X + c * p.Y);
        }

        (double X, double Y)[] green = [(-a, -b), (a, -b), (a, 0), (-a, 0)];
        V[] Corners(double degrees) => green.Select(p => Turn(p, degrees)).Select(p => d.Pt(p.X, p.Y)).ToArray();
        string greenShape = d.Polygon(null, Corners(0));
        string yellowShape = d.Polygon(null, Corners(-120));
        string blueShape = d.Polygon(null, Corners(120));
        d.Region(d.Fill("Yellow", "#FFFCF57A", null), "Difference", yellowShape, greenShape, blueShape);
        d.Region(d.Fill("Blue", "#FFB0EAF5", null), "Difference", blueShape, greenShape, yellowShape);
        d.Region(d.Fill("Green", "#FF66F2C9", null), "Difference", greenShape, yellowShape, blueShape);
        d.Region(d.Fill("YellowGreen", "#FFD7FC9C", null), "Intersection", yellowShape, greenShape);
        d.Region(d.Fill("BlueGreen", "#FF97F5F1", null), "Intersection", blueShape, greenShape);
        d.Region(d.Fill("YellowBlue", "#FFE7FEBD", null), "Intersection", yellowShape, blueShape);
        d.Polygon(d.Outline("YellowRim", "#FFF8DC10", 3.5), Corners(-120));
        d.Polygon(d.Outline("BlueRim", "#FF0090DE", 3.5), Corners(120));
        d.Polygon(d.Outline("GreenRim", "#FF10D020", 3.5), Corners(0));
        foreach (double degrees in new[] { 0.0, 120, -120 })
        {
            var dot = Turn((-a, -b / 2), degrees);
            d.DotAt(d.Pt(dot.X, dot.Y), size: 16);
        }

        var all = new[] { 0.0, 120, -120 }.SelectMany(t => green.Select(p => Turn(p, t))).ToArray();
        d.Scene(all.Min(p => p.X), all.Max(p => p.Y), all.Max(p => p.X), all.Min(p => p.Y));
    }

    // A square of area 36 cut into four parts of area 9: a green triangle from the top left corner,
    // a blue one on the bottom, a red one on the right side, and the orange part between. The length
    // of the red one's long side? Nothing to drag: equal areas fix every corner (side 6).
    [Puzzle("partitioned-square")]
    public static void PartitionedSquare(Lgf d)
    {
        var bl = d.Pt(0, 0, "BL");
        var br = d.Pt(6, 0, "BR");
        var tr = d.Pt(6, 6, "TR");
        var tl = d.Pt(0, 6, "TL");
        var t = d.Pt(3, 6, "T");
        var j = d.Pt(1.5, 3, "J");
        var k = d.Pt(3, 2, "K");
        string ink = Lgf.Ink;
        d.Polygon(d.Fill("Green", "#FFDFF09A", ink), tl, t, bl);
        d.Polygon(d.Fill("Blue", "#FFB8DBE6", ink), bl, br, j);
        d.Polygon(d.Fill("Red", "#FFFCA895", ink), br, tr, k);
        d.Polygon(d.Fill("Orange", "#FFFDE08A", ink), t, tr, k, j);
        d.Box(d.Outline("Frame", ink, 3), V.At(0, 0), V.At(6, 6));
        string blue = "#FF1494DE";
        var middle = Mid(d, tr, k);
        string arrow = d.Line("Arrow", blue, 3, dark: "#FF0A4F8A");
        d.Vector(arrow, middle, d.Pt(tr + V.At(-0.12, -0.16)));
        d.Vector(arrow, middle, d.Pt(k + V.At(0.12, 0.16)));
        d.Label("?", V.At(3.97, 4.25), d.ShapeText("BlueMark", blue, 36), 36);
        d.Label("9", V.At(0.86, 4.25), d.ShapeText("GreenMark", "#FF6CC020", 30), 30);
        d.Label("9", V.At(2.8, 3.7), d.ShapeText("OrangeMark", "#FFF09A10", 30), 30);
        d.Label("9", V.At(5, 2.5), d.ShapeText("RedMark", "#FFF02810", 30), 30);
        d.Label("9", V.At(2, 1.05), d.ShapeText("CyanMark", "#FF10B8E0", 30), 30);
        d.Scene(0, 6, 6, 0);
    }

    // A semicircle and two squares: one on the diagonal from the left end of the diameter, the other on
    // the diagonal from the right end, sharing the corner where the diagonals meet, each with a further
    // corner on the arc, those two corners 10 apart. The semicircle's area? A corner on the arc sees the
    // diameter at a right angle, so it is in line with the shared corner and the far end, and the two
    // corners on the arc are a quarter turn apart. Dragging the left one along the arc moves the shared
    // corner and resizes the squares: the corners stay 10 apart, the semicircle stays.
    [Puzzle("two-squares-overlapping-a-semi-circle")]
    public static void TwoSquaresOverlappingASemiCircle(Lgf d)
    {
        double R = 5 * Math.Sqrt(2);
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0, "E");
        var west = d.Pt(-R, 0, "W");
        var p = d.ArcSlider(o, R, 114, 156, 129, (0, 0), "P");
        var q = d.Pt(new V("P.Y", "-P.X"), "Q");
        var m = d.Crossing(west, q, p, east, "M");
        var light = d.Pt(west + m - p);
        var deep = d.Pt(m + east - q);
        string red = d.Outline("Red", "#FFF00810", 3.5);
        d.Arc("CircleSegment", red, o, east, west);
        d.Polygon(d.Outline("Cyan", "#FF00B2C0", 3.5), west, p, m, light);
        d.Polygon(d.Outline("Navy", "#FF0408B4", 3.5), m, q, east, deep);
        d.Dimension(d.Line("OrangeArrow", "#FFF09018", 3), d.Text("OrangeMark", "#FFF09018", 30), p, q, 0, "10", 0.95);
        d.Scene(-R, R, R, -3.1);
    }

    // Five copies of a right triangle in a rectangle: on top two with their long sides end to end on
    // a line across, their right angles touching the top; below, one standing in each corner and one
    // lying between with its short side up. What fraction do they cover? Nothing to drag: fitting
    // across both ways (two long sides = two short legs and a long leg) makes them 3-4-5 triangles,
    // the rectangle 10 by 6.4.
    [Puzzle("five-triangles-in-a-rectangle")]
    public static void FiveTrianglesInARectangle(Lgf d)
    {
        const double W = 10, H = 6.4, line = 4;
        string pink = d.Fill("Pink", "#FFFCA8E2", "#FF8E2598");
        d.Polygon(pink, d.Pt(0, line), d.Pt(1.8, H), d.Pt(5, line));
        d.Polygon(pink, d.Pt(5, line), d.Pt(6.8, H), d.Pt(W, line));
        d.Polygon(pink, d.Pt(0, line), d.Pt(0, 0), d.Pt(3, 0));
        d.Polygon(pink, d.Pt(3, 0), d.Pt(7, 0), d.Pt(7, 3));
        d.Polygon(pink, d.Pt(7, 0), d.Pt(W, 0), d.Pt(W, line));
        d.Segment(d.Line("Purple", "#FF8E2598", 2.5), d.Pt(0, line), d.Pt(W, line));
        d.Inked(d.Outline("Frame", "#FF0800A4", 3.5), d.Pt(0, 0), d.Pt(W, 0), d.Pt(W, H), d.Pt(0, H));
        d.Scene(0, H, W, 0);
    }
}
