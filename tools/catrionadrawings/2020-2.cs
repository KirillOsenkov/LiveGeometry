#nullable disable

static partial class Puzzles
{
    // A circle cut by its vertical diameter into four semicircles stacked along it, two on the right
    // (the top small one inside a big one) and two on the left, and a yellow circle on the diameter
    // between them; the green area is twice the yellow, what fraction is shaded? That fixes the
    // yellow circle's diameter at a third of the big one's. Dragging the yellow circle's center up
    // and down the diameter trades size between the small semicircles: the fraction stays a third.
    [Puzzle("multiple-semi-circles-in-a-circle")]
    public static void MultipleSemiCirclesInACircle(Lgf d)
    {
        const double R = 3;
        var o = d.Pt(0, 0, "O");
        var top = d.Pt(0, R, "T");
        var bottom = d.Pt(0, -R, "B");
        var c = d.Slider(V.At(0, -1.65), V.At(0, 1.65), (0.483 + 1.65) / 3.3, (0, 0.483), "C");
        var upper = d.Pt(new V("0", "C.Y + 1"), "U");
        var lower = d.Pt(new V("0", "C.Y - 1"), "L");
        var bigRight = d.Pt(new V("0", $"(C.Y + {R - 1}) / 2"), "Mr");
        var smallTop = d.Pt(new V("0", $"(C.Y + {R + 1}) / 2"), "Mt");
        var bigLeft = d.Pt(new V("0", $"(C.Y - {R - 1}) / 2"), "Ml");
        var smallBottom = d.Pt(new V("0", $"(C.Y - {R + 1}) / 2"), "Mb");
        string rightShape = d.Arc("CircleSegment", null, bigRight, lower, top, visible: false);
        string topShape = d.Arc("CircleSegment", null, smallTop, upper, top, visible: false);
        string leftShape = d.Arc("CircleSegment", null, bigLeft, upper, bottom, visible: false);
        string bottomShape = d.Arc("CircleSegment", null, smallBottom, lower, bottom, visible: false);
        string teal = "#FF0E9F92";
        string yellow = d.Circle(d.Fill("Yellow", "#FFFCF866", null), c, upper);
        string green = d.Fill("Green", "#FFC4EA80", null);
        d.Region(green, "Difference", rightShape, topShape, yellow);
        d.Region(green, "Difference", leftShape, bottomShape, yellow);
        // inked after coloring
        string line = d.Line("Teal", teal, 3);
        d.Circle(line, o, top);
        d.Segment(line, top, bottom);
        d.Arc("CircleArc", line, bigRight, lower, top);
        d.Arc("CircleArc", line, smallTop, upper, top);
        d.Arc("CircleArc", line, bigLeft, upper, bottom);
        d.Arc("CircleArc", line, smallBottom, lower, bottom);
        d.Circle(line, c, upper);
        d.Scene(-R, R, R, -R);
    }

    // Three squares of a 2 by 2 grid and a big tilted square with a corner at the middle of the grid's
    // bottom and its right side through the grid's center right corner; the angle at its top corner,
    // between its top side and the line from the grid's top middle? The top corner is on the circle
    // around the missing square's corner through its sides' ends. Dragging it along that circle turns
    // and resizes the tilted square: the angle stays 45°.
    [Puzzle("four-squares-iv")]
    public static void FourSquaresIV(Lgf d)
    {
        const double s = 2;
        var p = d.Pt(s, 0, "P");
        var a = d.Pt(s, 2 * s, "A");
        var corner = d.Pt(s, s, "K");
        var q = d.ArcSlider(corner, s, 10, 70, 33, (s, s), "Q");
        var diagonal = q - p;
        var right = d.Pt(p + Lgf.N(Math.Sqrt(0.5)) * diagonal.Turn(-45), "R");
        var left = d.Pt(p + Lgf.N(Math.Sqrt(0.5)) * diagonal.Turn(45), "Lc");
        string light = d.Fill("LightBlue", "#FF08CCFB", null);
        d.Box(light, V.At(0, 0), V.At(s, s));
        d.Box(light, V.At(0, s), V.At(s, 2 * s));
        d.Box(light, V.At(s, 0), V.At(2 * s, s));
        d.Polygon(d.Fill("Blue", "#FF0555F5", null), p, right, q, left);
        // inked after coloring: the grid's lines run across the tilted square
        string frame = d.Outline("Frame", "#FF14185A", 4);
        d.Box(frame, V.At(0, 0), V.At(s, s));
        d.Box(frame, V.At(0, s), V.At(s, 2 * s));
        d.Box(frame, V.At(s, 0), V.At(2 * s, s));
        d.Polygon(frame, p, right, q, left);
        string blue = d.Line("BlueLine", "#FF1A2BB0", 3);
        d.Segment(blue, a, q);
        d.Angle(d.Line("BlueMark", "#FF1020D8", 2.5), q, a, left, 32);
        d.Label("?", q + V.At(-0.8, 0.05), d.Text("BlueText", "#FF1020D8", 32));
        d.Scene(0, 2 * s, 4.45, 0);
    }

    // Two rectangles side by side on one base, lines from the bottom left corner to the top right
    // corners of both and one between those corners; the purple, pink and orange regions have area 12
    // each, the white triangle over the low rectangle? That makes the rectangles equally wide and the
    // low one half as high. Dragging the foot of the line between them changes their width (the
    // heights follow, areas kept): the white triangle stays 6.
    [Puzzle("triangle-over-two-rectangles")]
    public static void TriangleOverTwoRectangles(Lgf d)
    {
        var o = d.Pt(0, 0, "O");
        var f = d.Slider(V.At(3.9, 0), V.At(4.85, 0), (4.67 - 3.9) / 0.95, (4.67, 0), "F");
        string h = d.Num("24 / F.X", "h");
        var topLeft = d.Pt(new V("0", h), "TL");
        var b = d.Pt(new V("F.X", h), "B");
        var k = d.Pt(new V("F.X", $"{h} / 2"), "K");
        var c = d.Pt(new V("2 * F.X", $"{h} / 2"), "C");
        var end = d.Pt(new V("2 * F.X", "0"), "E");
        d.Polygon(d.Fill("Purple", "#FF4611BA", null), o, b, topLeft);
        d.Polygon(d.Fill("Pink", "#FFFD71B8", null), o, c, k, b);
        d.Polygon(d.Fill("Orange", "#FFFEB950", null), o, end, c);
        string frame = d.Outline("Frame", Lgf.Ink, 3.5);
        d.Polygon(frame, o, f, b, topLeft);
        d.Polygon(frame, f, end, c, k);
        string purple = d.Line("PurpleLine", "#FF3B1D9C", 3);
        d.Segment(purple, o, b);
        d.Segment(purple, o, c);
        d.Segment(purple, b, c);
        d.Label("?", new V("1.22 * F.X", $"0.66 * {h}"), d.Text("Mark", Lgf.Ink, 40));
        d.Scene(0, 5.14, 9.34, 0);
    }

    // A regular decagon with flat top and bottom, cut by the two long chords parallel to them, each
    // cap cut again by a diagonal from its end; what fraction is shaded? Nothing to drag: the decagon
    // is regular.
    [Puzzle("subdivided-decagon")]
    public static void SubdividedDecagon(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 3;
        // vertices counterclockwise from the right one, every 36°
        var v = Enumerable.Range(0, 10).Select(i => d.Pt(R * Math.Cos(i * Math.PI / 5), R * Math.Sin(i * Math.PI / 5))).ToArray();
        string yellow = d.Fill("Yellow", "#FFFEFC05", null);
        d.Polygon(yellow, v[4], v[3], v[2]);
        d.Polygon(yellow, v[4], v[5], v[6], v[9], v[0], v[1]);
        d.Polygon(yellow, v[9], v[8], v[7]);
        string frame = d.Outline("Frame", Lgf.Ink, 3.5);
        d.Polygon(frame, v);
        string ink = d.Line("Ink", Lgf.Ink, 3.5);
        d.Segment(ink, v[4], v[1]);
        d.Segment(ink, v[4], v[2]);
        d.Segment(ink, v[6], v[9]);
        d.Segment(ink, v[9], v[7]);
        d.Scene(-R, R * Math.Sin(2 * Math.PI / 5), R, -R * Math.Sin(2 * Math.PI / 5));
    }

    // A semicircle and two congruent rectangles: one standing on the right half of the diameter, the
    // other tilted, a corner at the diameter's left end, the next on the arc and the third on the
    // first one's top left corner; the angle between them there? Nothing to drag: the conditions fix
    // the rectangles (the long side the radius, the short one the radius over √3).
    [Puzzle("two-congruent-rectangles-and-a-semi-circle")]
    public static void TwoCongruentRectanglesAndASemiCircle(Lgf d)
    {
        const double R = 4;
        double h = R / Math.Sqrt(3);
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0, "E");
        var west = d.Pt(-R, 0, "W");
        var t = d.Pt(-R / 2, R * S3, "T");
        var p = d.Pt(0, h, "P");
        var q = d.Pt(R, h, "Q");
        var foot = d.Pt(-R + h * S3, -h / 2, "F");
        string green = d.Fill("Green", "#FF3CC97A", null);
        d.Polygon(green, west, foot, p, t);
        d.Polygon(green, o, east, q, p);
        string line = d.Line("GreenLine", "#FF0A8F5E", 2.5);
        d.Arc("CircleArc", line, o, east, west);
        d.Segment(line, west, o);
        // inked after coloring, over the arc
        string frame = d.Outline("Frame", Lgf.Ink, 4);
        d.Inked(frame, west, foot, p, t);
        d.Inked(frame, o, east, q, p);
        string cyan = "#FF22C8E8";
        d.ArcMark(d.Line("CyanLine", cyan, 3), p, q, t, 0.5);
        d.Label("?", p + V.At(0.2, 0.9), d.Text("CyanMark", cyan, 36));
        d.Scene(-R, R, R, -h / 2);
    }

    // A blue rectangle in a square's corner, covering half of it; the square's diagonal crosses the
    // rectangle's top and right side, and lines from the corner run through those crossings. The angle
    // between them? Dragging the rectangle's top left corner up and down changes its shape (its area
    // kept): the angle stays 45°.
    [Puzzle("a-rectangle-in-a-square")]
    public static void ARectangleInASquare(Lgf d)
    {
        const double L = 6;
        var o = d.Pt(0, 0, "O");
        var rectangleTop = d.Slider(V.At(0, 3.3), V.At(0, 5.1), (3.956 - 3.3) / 1.8, (0, 3.956), "RT");
        string h = "RT.Y";
        string w = d.Num($"{L * L / 2} / {h}", "w");
        var f = d.Pt(new V(w, "0"), "F");
        var topLeft = d.Pt(0, L, "TL");
        var topRight = d.Pt(L, L, "TR");
        var bottomRight = d.Pt(L, 0, "BR");
        var corner = d.Pt(new V(w, h), "K");
        var top = d.Pt(new V($"{L} - {h}", h), "P1");
        var side = d.Pt(new V(w, $"{L} - {w}"), "P2");
        var topEnd = d.Pt(new V($"{L} * ({L} - {h}) / {h}", Lgf.N(L)), "E1");
        var sideEnd = d.Pt(new V(Lgf.N(L), $"{L} * ({L} - {w}) / {w}"), "E2");
        d.Polygon(d.Fill("Purple", "#FFB38FDD", null), o, bottomRight, topRight, topLeft);
        d.Polygon(d.Fill("Blue", "#FF45B6EF", null), o, f, corner, rectangleTop);
        string frame = d.Outline("Frame", Lgf.Ink, 4);
        d.Polygon(frame, o, bottomRight, topRight, topLeft);
        d.Polygon(frame, o, f, corner, rectangleTop);
        string blue = "#FF1E28C8";
        string line = d.Line("BlueLine", blue, 4);
        d.Segment(line, topLeft, bottomRight);
        d.Segment(line, o, topEnd);
        d.Segment(line, o, sideEnd);
        d.ArcMark(d.Line("BlueArc", blue, 3), o, side, top, 0.6);
        d.Label("?", V.At(0.95, 1.05), d.ShapeText("BlueMark", blue, 36));
        d.Scene(0, L, L, 0);
    }

    // Two half squares and a circle of radius 1: the orange one's base ends are on the circle, the red
    // one's top corners too, and their hypotenuses lie on one line. The total shaded area? That puts
    // the red one's top right corner at the orange one's bottom right turned a quarter turn about the
    // center. Dragging the orange one's left corner along the circle trades size between them: the
    // total stays 2.
    [Puzzle("two-half-squares-and-a-circle")]
    public static void TwoHalfSquaresAndACircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 3;
        var o = d.Pt(0, 0, "O");
        var rim = d.Pt(R, 0);
        var a = d.ArcSlider(o, R, 200, 236, 209.6, (0, 0), "A");
        string p = "(-A.X)", k = "(-A.Y)";
        var b = d.Pt(new V(p, "A.Y"), "B");
        var c = d.Pt(new V(p, $"A.Y + 2 * {p}"), "C");
        var topLeft = d.Pt(new V($"-{k}", p), "TL");
        var topRight = d.Pt(new V(k, p), "TR");
        var foot = d.Pt(new V($"-{k}", $"{p} - 2 * {k}"), "F");
        string ink = Lgf.Ink;
        d.Polygon(d.Fill("Orange", "#FFFD7403", ink, 3.5), a, b, c);
        d.Polygon(d.Fill("Red", "#FFFC1304", ink, 3.5), topLeft, topRight, foot);
        d.Circle(d.Outline("Frame", ink, 3.5), o, rim);
        d.Scene(-R, 1.3 * R, R, -R);
    }

    // Three congruent rectangles in a square: two standing along its left and right sides, the third
    // tilted with its corners at the left one's bottom right corner and the right one's top left.
    // What fraction of the square do they cover? Nothing to drag: congruence makes them a third of
    // the square wide.
    [Puzzle("three-congruent-rectangles-iii")]
    public static void ThreeCongruentRectanglesIII(Lgf d)
    {
        d.Texture = "FeltTip";
        const double w = 1.2, s = 3 * w;
        var b = d.Pt(w, 0, "B");
        var t = d.Pt(2 * w, s, "T");
        var left = d.Pt(0.2 * w, 0.6 * w, "Lf");
        var right = d.Pt(2.8 * w, 2.4 * w, "Rt");
        var leftTop = d.Pt(w, s);
        var rightBottom = d.Pt(2 * w, 0);
        var o = d.Pt(0, 0);
        var topLeft = d.Pt(0, s);
        var corner = d.Pt(s, s);
        var bottomRight = d.Pt(s, 0);
        string leftShape = d.Polygon(null, o, b, leftTop, topLeft);
        string rightShape = d.Polygon(null, rightBottom, bottomRight, corner, t);
        string tilted = d.Polygon(null, b, right, t, left);
        // one shape: three fills side by side would show seams under the dark theme
        d.Region(d.Fill("Teal", "#FF02C6C0", null), "Union", leftShape, rightShape, tilted);
        // across the teal, which stays bright in the dark theme, and the paper: a mid-tone there
        string frame = d.Outline("Frame", "#FF0A1446", 3.5, dark: "#FF3A4AD0");
        d.Polygon(frame, o, b, leftTop, topLeft);
        d.Polygon(frame, rightBottom, bottomRight, corner, t);
        d.Polygon(frame, b, right, t, left);
        d.Polygon(frame, o, bottomRight, corner, topLeft);
        d.Scene(0, s, s, 0);
    }

    // A checkerboard of 7 by 2 squares, 2 wide, and a semicircle on its top side; the perimeter of
    // the combined shape? Nothing to drag: the squares fix it.
    [Puzzle("squares-and-a-semi-circle")]
    public static void SquaresAndASemiCircle(Lgf d)
    {
        const double W = 2, a = W / 7;
        var o = d.Pt(1, 2 * a, "O");
        var east = d.Pt(W, 2 * a);
        var west = d.Pt(0, 2 * a);
        d.Arc("CircleSegment", d.Fill("Pink", "#FFFBC5F4", null), o, east, west);
        d.Arc("CircleArc", d.Line("Magenta", "#FFD601AA", 3), o, east, west);
        string green = "#FF28A886";
        string dark = d.Fill("Green", "#FF3ADFA0", green, 3);
        string light = d.Fill("LightGreen", "#FFD1FC7B", green, 3);
        for (int i = 0; i < 7; i++)
        {
            for (int j = 0; j < 2; j++)
            {
                d.Box((i + j) % 2 == 1 ? dark : light, V.At(i * a, j * a), V.At((i + 1) * a, (j + 1) * a));
            }
        }

        string purple = "#FF6A69C8";
        d.Dimension(d.Line("PurpleArrow", purple, 3), d.Text("PurpleMark", purple, 30), d.Pt(W, 0), d.Pt(0, 0), 0.12, "2", 0.08);
        d.Scene(0, 2 * a + 1, W, -0.2);
    }

    // A square in a quarter circle centered at its bottom left corner, the arc through the middles of
    // its top and right sides, and a semicircle on its bottom side; the orange area is 10, the blue?
    // Nothing to drag: the arc through the middles fixes the quarter circle's radius at √5 / 2 of the side.
    [Puzzle("semi-circle-in-a-square-in-a-quarter-circle")]
    public static void SemiCircleInASquareInAQuarterCircle(Lgf d)
    {
        const double a = 4;
        double r = a * Math.Sqrt(5) / 2;
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(r, 0);
        var north = d.Pt(0, r);
        var bottomRight = d.Pt(a, 0);
        var topRight = d.Pt(a, a);
        var topLeft = d.Pt(0, a);
        var middle = d.Pt(a / 2, 0);
        d.Arc("CircleSector", d.Fill("Blue", "#FF0AC2FA", "#FF1040E0", 3), o, east, north);
        d.Arc("CircleSegment", d.Fill("Orange", "#FFFCAD66", "#FFD0102A", 3), middle, bottomRight, o);
        // inked after coloring, over the arcs
        d.Inked(d.Outline("Frame", Lgf.Ink, 3.5), o, bottomRight, topRight, topLeft);
        var top = d.DotAt(V.At(a / 2, a), size: 13);
        var side = d.DotAt(V.At(a, a / 2), size: 13);
        string ink = d.Line("Ink", Lgf.Ink, 3.5);
        d.Segment(ink, topLeft, top, "TwoTicks");
        d.Segment(ink, top, topRight, "TwoTicks");
        d.Segment(ink, topRight, side, "TwoTicks");
        d.Segment(ink, side, bottomRight, "TwoTicks");
        d.Label("10", V.At(a / 2 - 0.1, a / 4), d.ShapeText("RedMark", "#FFE01020", 32));
        d.Scene(0, r, r, 0);
    }

    // Three concentric semicircles; the blue right triangle stands on the diameter at the middle one's
    // left end, its top on the big arc and its hypotenuse to the diameter's right end, touching the
    // middle arc; the red one runs from the same corner to that touching point and down, touching the
    // small arc. What fraction is yellow? Nothing to drag: the touching fixes the radii at 1, 1/2, 1/4.
    [Puzzle("two-triangles-and-three-concentric-semi-circles")]
    public static void TwoTrianglesAndThreeConcentricSemiCircles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4;
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0, "E");
        var west = d.Pt(-R, 0, "W");
        var middleEast = d.Pt(R / 2, 0);
        var f = d.Pt(-R / 2, 0, "F");
        var smallEast = d.Pt(R / 4, 0, "G");
        var smallWest = d.Pt(-R / 4, 0);
        var p = d.Pt(-R / 2, R * S3, "P");
        var t = d.Pt(R / 4, R / 2 * S3, "T");
        string yellow = d.Fill("Yellow", "#FFFFFE02", null);
        d.Arc("CircleSegment", yellow, o, east, west);
        d.Arc("CircleSegment", d.PaperFill("Hole", null), o, middleEast, f);
        d.Arc("CircleSegment", yellow, o, smallEast, smallWest);
        // the arcs part the yellow, which stays bright in the dark theme, from the paper, and the blue
        // lines cross the yellow: a mid-tone and a deep blue there
        string cyan = d.Line("Cyan", "#FF18B4FA", 3, dark: "#FF1C8CD8");
        d.Arc("CircleArc", cyan, o, east, west);
        d.Arc("CircleArc", cyan, o, middleEast, f);
        d.Arc("CircleArc", cyan, o, smallEast, smallWest);
        d.Segment(d.Line("Diameter", "#FF2E7FE8", 3), west, east);
        string blue = d.Line("Blue", "#FF0A18A8", 3.5, dark: "#FF1A2A90");
        d.Segment(blue, f, p);
        d.Segment(blue, p, east);
        string red = d.Line("Red", "#FFE25858", 3);
        d.Segment(red, f, t);
        d.Segment(red, t, smallEast);
        d.Scene(-R, R, R, 0);
    }

    // Five congruent rectangles: one lying, one standing, one lying with two standing on it, and lines
    // from the standing one's top left corner to the first one's top left corner and to the last one's
    // bottom right. The angle between them? Nothing to drag: the two standing on one lying make the
    // rectangles twice as long as wide.
    [Puzzle("five-congruent-rectangles")]
    public static void FiveCongruentRectangles(Lgf d)
    {
        const double w = 1.5, L = 2 * w;
        var a = d.Pt(0, w, "A");
        var v = d.Pt(L, L, "V");
        var b = d.Pt(2 * L + w, w, "B");
        string magenta = d.Fill("Magenta", "#FFF8025E", "#FF3A1840", 3.5);
        d.Box(magenta, V.At(0, 0), V.At(L, w));
        d.Box(magenta, V.At(L, 0), V.At(L + w, L));
        d.Box(magenta, V.At(L + w, 0), V.At(2 * L + w, w));
        d.Box(magenta, V.At(L + w, w), V.At(L + 2 * w, w + L));
        d.Box(magenta, V.At(L + 2 * w, w), V.At(2 * L + w, w + L));
        string purple = "#FF3A20A8";
        string line = d.Line("Purple", purple, 3);
        d.Segment(line, a, v);
        d.Segment(line, v, b);
        d.ArcMark(line, v, a, b, 0.33);
        d.Label("?", v + V.At(-0.22, -0.7), d.Text("PurpleMark", "#FF5A12BB", 34));
        d.Scene(0, w + L, 2 * L + w, 0);
    }

    // A square split into five rectangles of equal area: a strip down the right side, one 2 high
    // across the top beside it, one down the left side below that and two stacked in the middle. The
    // square's area? Nothing to drag: the equal areas make the side 8.
    [Puzzle("five-rectangles-tiling-a-square")]
    public static void FiveRectanglesTilingASquare(Lgf d)
    {
        const double s = 8, p = s / 5, q = s * s / 5 / (s - 2);
        // the fills stay bright in the dark theme: ink between them there too
        d.Box(d.Fill("Orange", "#FFFCD478", width: 3.5, darkOutline: Lgf.Ink), V.At(0, s - 2), V.At(s - p, s));
        d.Box(d.Fill("Pink", "#FFFBB2EA", width: 3.5, darkOutline: Lgf.Ink), V.At(s - p, 0), V.At(s, s));
        d.Box(d.Fill("Purple", "#FF8662D9", width: 3.5, darkOutline: Lgf.Ink), V.At(0, 0), V.At(q, s - 2));
        d.Box(d.Fill("Green", "#FFD6FB79", width: 3.5, darkOutline: Lgf.Ink), V.At(q, (s - 2) / 2), V.At(s - p, s - 2));
        d.Box(d.Fill("Blue", "#FF04C6EF", width: 3.5, darkOutline: Lgf.Ink), V.At(q, 0), V.At(s - p, (s - 2) / 2));
        string red = "#FFE82D52";
        d.Dimension(d.Line("RedArrow", red, 3), d.Text("RedMark", red, 30), d.Pt(0, s - 2), d.Pt(0, s), 0.5, "2", 0.35);
        d.Scene(-0.9, s, s, 0);
    }
}
