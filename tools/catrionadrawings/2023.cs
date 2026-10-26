#nullable disable

static partial class Puzzles
{
    // A trophy: a regular hexagon with a square on its top side and three squares hanging from its
    // bottom side, white lines from the top square's far corner and the bottom square's to the
    // hexagon's right corner; the angle between them? She drew it in gold and white on black: here
    // gold and ink on the paper, and her own gold and white under the dark theme. Nothing to drag: the
    // squares fix it (the angle would stay for rectangles whose heights multiply to three squares, but
    // she says squares).
    [Puzzle("four-rectangles-and-a-hexagon")]
    public static void FourRectanglesAndAHexagon(Lgf d)
    {
        const double s = 2, h = 2 * S3;
        var hexagon = HexagonFromLeft(d, V.At(-s, 0), Lgf.N(s));
        var r = hexagon[3];
        // drawn as segments: a line style can say its own color on the dark paper (her gold)
        string gold = d.Line("Gold", "#FFC9A430", 3, dark: "#FFDAC660");
        Drawing2023.Around(d, gold, hexagon);
        Drawing2023.Around(d, gold, d.Pt(-s / 2, h), d.Pt(s / 2, h), d.Pt(s / 2, h + s), d.Pt(-s / 2, h + s));
        for (int i = 1; i <= 3; i++)
        {
            Drawing2023.Around(d, gold, d.Pt(-s / 2, -h - (i - 1) * s), d.Pt(s / 2, -h - (i - 1) * s), d.Pt(s / 2, -h - i * s), d.Pt(-s / 2, -h - i * s));
        }

        // her lines were white on the black paper: ink on the light paper, near white on the dark one
        var a = d.Pt(-s / 2, h + s, "A");
        var b = d.Pt(-s / 2, -h - 3 * s, "B");
        string chalk = d.Line("Chalk", Lgf.Ink, 2.5, dark: "#FFF4F4F4");
        d.Segment(chalk, a, r);
        d.Segment(chalk, r, b);
        d.Angle(chalk, r, a, b, 22);
        d.Label("?", r + V.At(-0.75, 0.05), d.Text("Mark", Lgf.Ink, 24, dark: "#FFF4F4F4"));
        d.Scene(-2, h + s, 2, -h - 3 * s);
    }

    // Two quarter circles sharing a corner, the big one's top and left end and the small one's top on
    // the yellow circle of radius 4; their total area? The circle then also crosses the bottom line
    // as far left of the corner as the small one's radius, and its center is on the diagonal through
    // the corner. Dragging the big one's top along the circle trades size between the two: the
    // squares of their radii still add up to twice the circle's, so the total stays.
    [Puzzle("two-quarter-circles-and-a-circle")]
    public static void TwoQuarterCirclesAndACircle(Lgf d)
    {
        const double R = 4;
        var o = d.Pt(0, 0, "O");
        var rim = d.Pt(R, 0);
        // the circle centered at O: the corner is (m, -m) and the radii are m + n and m - n, (m, n) the top
        var t = d.ArcSlider(o, R, 16, 38, 28.3, (0, 0), "T");
        var p = d.Pt(new V("T.X", "-T.X"), "P");
        var left = d.Pt(new V("-T.Y", "-T.X"), "L");
        var smallTop = d.Pt(new V("T.X", "-T.Y"), "S");
        var smallEnd = d.Pt(new V("2 * T.X - T.Y", "-T.X"), "E");
        string green = d.Fill("Green", "#FFACDFBB", "#FF2F4A40");
        d.Arc("CircleSector", green, p, t, left);
        d.Arc("CircleSector", green, p, smallEnd, smallTop);
        d.Circle(d.Line("Yellow", "#FFF0C04C", 2.5), o, rim);
        d.Scene(-4, 4, 5.5, -4);
    }

    // A regular hexagon cut by its long diagonal, the line from its left corner to the top right one,
    // and a line from its top side to the bottom right corner, cut by the other two in the ratio
    // 1 : 2 : 3; what fraction is shaded? Nothing to drag: the ratio puts the top end a tenth of the
    // side left of the middle.
    [Puzzle("subdivided-hexagon-iv")]
    public static void SubdividedHexagonIV(Lgf d)
    {
        const double s = 2, h = 2 * S3;
        var hexagon = HexagonFromLeft(d, V.At(-s, 0), Lgf.N(s));
        var (left, topRight, right, bottomRight) = (hexagon[0], hexagon[2], hexagon[3], hexagon[4]);
        var x = d.Pt(-0.2, h, "X");
        var y = d.Pt(0, h * 2 / 3, "Y");
        var z = d.Pt(0.4, 0, "Z");
        string orange = d.Fill("Orange", "#FFFCBF7D", null);
        d.Polygon(orange, x, topRight, y);
        d.Polygon(orange, left, y, z);
        d.Polygon(orange, z, right, bottomRight);
        string line = d.Line("OrangeLine", "#FFEE8A4A", 2.5);
        d.Segment(line, left, topRight);
        d.Segment(line, left, right);
        d.Segment(line, x, bottomRight);
        d.Polygon(d.Outline("Frame", "#FF2E2E4A"), hexagon);
        string mark = d.Text("OrangeMark", "#FFE8743A", 24);
        d.Label("1", V.At(-0.32, 1.38), mark);
        d.Label("2", V.At(0.38, 0.62), mark);
        d.Label("3", V.At(0.42, -0.9), mark);
        d.Scene(-s, h, s, -h);
    }

    // A rectangle 4 wide, a semicircle on its left side and a quarter circle in its bottom right corner
    // as high as it is, touching; the shaded area? Nothing to drag: touching fixes the height at 2√2.
    [Puzzle("a-semi-circle-and-a-quarter-circle-in-a-rectangle")]
    public static void ASemiCircleAndAQuarterCircleInARectangle(Lgf d)
    {
        const double w = 4;
        double h = 2 * Math.Sqrt(2);
        var bottomLeft = d.Pt(0, 0);
        var topLeft = d.Pt(0, h);
        var middle = d.Pt(0, h / 2);
        var corner = d.Pt(w, 0);
        var topRight = d.Pt(w, h);
        var foot = d.Pt(w - h, 0);
        string purple = d.Fill("Purple", "#FF9C87C2", "#FF5E46B4");
        d.Arc("CircleSegment", purple, middle, bottomLeft, topLeft);
        d.Arc("CircleSector", purple, corner, topRight, foot);
        // inked after coloring: the rectangle runs over the shapes' straight sides (as segments: a
        // polygon is drawn under every arc and circle, whatever the order)
        string ink = d.Line("Ink");
        Drawing2023.Around(d, ink, bottomLeft, corner, topRight, topLeft);
        d.Dimension(ink, d.Text(), topLeft, topRight, 0.32, "4", 0.2);
        d.Scene(0, h + 0.5, w, 0);
    }

    // Three squares in a row, a line from the first one's top left corner to the last one's bottom
    // right, one from the first one's bottom left to the middle one's top right, and one from the
    // middle one's bottom left through where those two cross; the angle there? Nothing to drag: the
    // squares fix it. The answer is a right angle: a plain arc, not the square of an angle mark.
    [Puzzle("three-squares-viii")]
    public static void ThreeSquaresVIII(Lgf d)
    {
        const double s = 2;
        string blue = "#FF0A6DF5";
        d.Box(d.Fill("LightBlue", "#FFD4FBFC", blue, 3), V.At(0, 0), V.At(s, s));
        d.Box(d.Fill("LightGreen", "#FFEDFEDA", blue, 3), V.At(s, 0), V.At(2 * s, s));
        d.Box(d.Fill("Blue", "#FFBCE6FB", blue, 3), V.At(2 * s, 0), V.At(3 * s, s));
        var c = d.Pt(1.2 * s, 0.6 * s, "C");
        var foot = d.Pt(s, 0);
        var end = d.Pt(3 * s, 0);
        // every line is on the fills, which stay light in the dark theme: deep there too
        string purple = d.Line("Purple", "#FF9A3FC8", 2.5, dark: "#FF3E0F6E");
        d.Segment(purple, d.Pt(0, s), end);
        d.Segment(purple, d.Pt(0, 0), d.Pt(2 * s, s));
        d.Segment(purple, foot, d.Pt(4 * s / 3, s));
        d.ArcMark(purple, c, foot, end, 0.26);
        d.Label("?", V.At(1.31 * s, 0.32 * s), d.ShapeText("PurpleMark", "#FF9A3FC8", 28));
        d.Scene(0, s, 3 * s, 0);
    }

    // A semicircle with two lines up from its diameter to the arc, 4 and 2 long, their feet 6 apart;
    // its area? Nothing to drag: the lengths fix it (the center 2 right of the long one's foot, the
    // radius √20).
    [Puzzle("three-lengths-in-a-semi-circle")]
    public static void ThreeLengthsInASemiCircle(Lgf d)
    {
        double r = Math.Sqrt(20);
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(r, 0);
        var west = d.Pt(-r, 0);
        var redFoot = d.Pt(-2, 0);
        var redTop = d.Pt(-2, 4);
        var orangeFoot = d.Pt(4, 0);
        var orangeTop = d.Pt(4, 2);
        string navy = d.Line("Navy", "#FF2C4468", 3);
        d.Arc("CircleArc", navy, o, east, west);
        d.Segment(navy, west, redFoot);
        d.Segment(navy, orangeFoot, east);
        string purple = "#FF983F92", red = "#FFC8432E", orange = "#FFE88A2C";
        d.Segment(d.Line("Purple", purple, 2.5), redFoot, orangeFoot);
        d.Segment(d.Line("Red", red, 2.5), redFoot, redTop);
        d.Segment(d.Line("Orange", orange, 2.5), orangeFoot, orangeTop);
        string mark = d.Line("RightAngle", "#FF2C4468", 1.5);
        d.Angle(mark, redFoot, redTop, west, 12);
        d.Angle(mark, orangeFoot, east, orangeTop, 12);
        d.Label("4", V.At(-1.35, 1.9), d.Text("RedMark", red, 30));
        d.Label("2", V.At(3.4, 0.95), d.Text("OrangeMark", orange, 30));
        d.Label("6", V.At(1, 0.45), d.Text("PurpleMark", purple, 30));
        d.Scene(-r, r, r, 0);
    }

    // Two blue squares stacked, a big purple square of side 6 on their bottom left corner whose top
    // side runs through their top left corner, 2 from its end; the yellow triangle from the bottom
    // corner to the blue squares' middle right corner and the big square's right corner? Nothing to
    // drag: the 2 and 4 fix the tilt (1 in 3) and the blue squares (side √10).
    [Puzzle("three-squares-vii")]
    public static void ThreeSquaresVII(Lgf d)
    {
        d.Texture = "FeltTip";
        double a = Math.Sqrt(10), k = 1 / Math.Sqrt(10);
        var o = d.Pt(0, 0, "O");
        var right = d.Pt(18 * k, 6 * k, "P");
        var topRight = d.Pt(12 * k, 24 * k, "Q");
        var topLeft = d.Pt(-6 * k, 18 * k, "K");
        var split = d.Pt(0, 2 * a, "X");
        var middle = d.Pt(a, a, "M");
        d.Polygon(d.Fill("Yellow", "#FFFEFF5C", "#FFE0C440"), o, middle, right);
        // inked after coloring: the blue squares' sides run over the yellow
        string blue = d.Line("Blue", "#FF2996D5", 3);
        Drawing2023.Around(d, blue, o, d.Pt(a, 0), d.Pt(a, 2 * a), split);
        d.Segment(blue, middle, d.Pt(0, a));
        string purple = "#FF9C5AA6";
        Drawing2023.Around(d, d.Line("Purple", purple, 3), o, right, topRight, topLeft);
        string arrow = d.Line("PurpleArrow", purple, 2);
        string mark = d.Text("PurpleMark", purple, 26);
        d.Dimension(arrow, mark, topLeft, split, 0.45, "2", 0.3);
        d.Dimension(arrow, mark, split, topRight, 0.45, "4", 0.3);
        d.Scene(-6 * k - 0.3, 24 * k + 0.6, 18 * k, 0);
    }
}

file static class Drawing2023
{
    /// <summary>A closed outline as segments, one per side: a polygon's sides in a line style</summary>
    public static void Around(Lgf d, string lineStyle, params V[] points)
    {
        for (int i = 0; i < points.Length; i++)
        {
            d.Segment(lineStyle, points[i], points[(i + 1) % points.Length]);
        }
    }
}
