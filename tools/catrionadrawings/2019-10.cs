#nullable disable

static partial class Puzzles
{
    // A semicircle of diameter 6, an equilateral triangle in its middle with its top on the arc, and a
    // square standing on the diameter line beside it, its left side through the triangle's corner and
    // its top left corner on the arc. Its area? Nothing to drag: the diameter fixes it (the square's
    // side is √(9 - 3)).
    [Puzzle("triangle-in-semi-circle-with-square")]
    public static void TriangleInSemiCircleWithSquare(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 3;
        double half = R / Math.Sqrt(3), s = Math.Sqrt(R * R - half * half);
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0, "E");
        var west = d.Pt(-R, 0, "W");
        var top = d.Pt(0, R, "T");
        var left = d.Pt(-half, 0);
        var right = d.Pt(half, 0);
        d.Box(d.Fill("Orange", "#FFFE9A08", "#FFF2600A"), V.At(half, 0), V.At(half + s, s));
        string purple = "#FF5A2AA8";
        d.Arc("CircleSegment", d.Outline("Frame", purple), o, east, west);
        d.Inked(d.Outline("Pink", "#FFD84068"), left, right, top);
        var low = d.Pt(-R, -0.65);
        var high = d.Pt(R, -0.65);
        d.Dimension(d.Line("Arrow", purple), d.Text("PurpleMark", purple, 30), low, high, 0, "6", 0.5);
        d.Scene(-R, R, half + s, -0.9);
    }

    // Two overlapping equilateral triangles, one pointing left and one right, and three equal circles in
    // a row between them: the outer ones in the triangles' points, touching the other triangle's
    // upright side, the middle one between the two upright sides. The total area? Nothing to drag:
    // the circles fix the shape (each triangle is 5 radii high, the points 8 radii apart).
    [Puzzle("circles-in-overlapping-triangles")]
    public static void CirclesInOverlappingTriangles(Lgf d)
    {
        double half = 5 / Math.Sqrt(3);
        string cyan = d.Outline("Cyan", "#FF1CC4E2", 4);
        foreach (double x in new[] { 2.0, 4, 6 })
        {
            d.Circle(cyan, d.Pt(x, 0), d.Pt(x + 1, 0));
        }

        d.Inked(d.Outline("Orange", "#FFE8800A", 3), d.Pt(0, 0), d.Pt(5, -half), d.Pt(5, half));
        d.Inked(d.Outline("Red", "#FFC81408", 3), d.Pt(8, 0), d.Pt(3, half), d.Pt(3, -half));
        d.Scene(0, half, 8, -half);
    }

    // Two congruent equilateral triangles sharing a corner, and two green lines joining their other
    // corners in turn; the angle where the lines cross? Dragging the right triangle's top corner turns
    // it about the shared corner: the lines join each point to its image under that turn, and the
    // angle stays the triangles' 60°.
    [Puzzle("tilted-triangles")]
    public static void TiltedTriangles(Lgf d)
    {
        const double L = 4;
        var p = d.Pt(0, 0, "P");
        var a = d.Pt(L * Math.Cos(155.5 * Math.PI / 180), L * Math.Sin(155.5 * Math.PI / 180), "A");
        var b = d.Pt(L * Math.Cos(95.5 * Math.PI / 180), L * Math.Sin(95.5 * Math.PI / 180), "B");
        var c = d.ArcSlider(p, L, 45, 78, 59.5, (0, 0), "C");
        var e = d.Pt(c.Turn(-60), "D");
        var x = d.Crossing(a, c, b, e, "X");
        string green = d.Line("Green", "#FF9ADB4E", 2.5);
        d.Segment(green, a, c);
        d.Segment(green, b, e);
        string ink = d.Line("Ink", Lgf.Ink, 3);
        d.Segment(ink, p, a);
        d.Segment(ink, a, b);
        d.Segment(ink, b, p);
        d.Segment(ink, p, c);
        d.Segment(ink, c, e);
        d.Segment(ink, e, p);
        string blue = "#FF100DBC";
        d.ArcMark(d.Line("Blue", blue, 3), x, c, b, 0.4);
        string toC = d.Num("distance(X, C)");
        string toB = d.Num("distance(X, B)");
        var bisector = new V($"(C.X - X.X) / {toC} + (B.X - X.X) / {toB}", $"(C.Y - X.Y) / {toC} + (B.Y - X.Y) / {toB}");
        string length = d.Num($"sqrt(({bisector.X}) ^ 2 + ({bisector.Y}) ^ 2)");
        d.Label("?", x + new V($"0.72 * ({bisector.X}) / {length}", $"0.72 * ({bisector.Y}) / {length}"), d.Text("BlueMark", blue, 26));
        d.Scene(-3.7, 4.2, 4.1, -0.3);
    }

    // A square, a quarter circle centered at its bottom right corner from the bottom left corner to the
    // top right, and a small quarter circle centered on the right side, down to the bottom right corner,
    // its other end on the big arc. What fraction is shaded? Nothing to drag: the end on the arc makes
    // the small radius the side over √2.
    [Puzzle("two-quarter-circles-in-a-square")]
    public static void TwoQuarterCirclesInASquare(Lgf d)
    {
        const double S = 4;
        double r = S / Math.Sqrt(2);
        var corner = d.Pt(S, 0, "K");
        var topRight = d.Pt(S, S);
        var origin = d.Pt(0, 0);
        var center = d.Pt(S, r, "C");
        var end = d.Pt(S - r, r, "E");
        string big = d.Arc("CircleSector", null, corner, topRight, origin, visible: false);
        string small = d.Arc("CircleSector", null, center, end, corner, visible: false);
        string above = d.Box(null, V.At(S - r, r), V.At(S, S));
        d.Region(d.Fill("Blue", "#FF84D2FC", null), "Difference", big, small, above);
        d.Box(d.Outline("Brown", "#FF542004", 3), V.At(0, 0), V.At(S, S));
        d.Arc("CircleSector", d.Outline("BlueArc", "#FF5014E8", 3), corner, topRight, origin);
        d.Arc("CircleSector", d.Outline("Purple", "#FF9A0592", 3), center, end, corner);
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(-0.2, S, S + 0.2, -0.45);
    }

    // A regular octagon, a square of area 8 on each of two opposite slanted sides, inside it, and the
    // rectangle between the squares' inner sides from one of the other slanted sides to the one opposite.
    // Its area? Nothing to drag: the octagon's side is the squares', and the rectangle is
    // a(1 + √2) by a(√2 - 1).
    [Puzzle("two-squares-and-a-rectangle-in-an-octagon")]
    public static void TwoSquaresAndARectangleInAnOctagon(Lgf d)
    {
        d.Texture = "FeltTip";
        double a = 2 * Math.Sqrt(2), m = a * (1 + Math.Sqrt(2)) / 2, k = Math.Sqrt(0.5);
        // a point given along the up-right diagonal (u) and the up-left one (w)
        V At(double u, double w) => d.Pt(k * (u - w), k * (u + w));
        double radius = a / (2 * Math.Sin(Math.PI / 8));
        var octagon = Enumerable.Range(0, 8).Select(i => d.Pt(radius * Math.Cos((22.5 + 45 * i) * Math.PI / 180), radius * Math.Sin((22.5 + 45 * i) * Math.PI / 180))).ToArray();
        d.Polygon(d.Fill("Teal", "#FF01CFA8", "#FF3D3A33", 3), At(-m, m - a), At(m, m - a), At(m, a - m), At(-m, a - m));
        string green = d.Outline("Green", "#FF0C9030");
        d.Polygon(green, At(-a / 2, m), At(a / 2, m), At(a / 2, m - a), At(-a / 2, m - a));
        d.Polygon(green, At(-a / 2, -m), At(a / 2, -m), At(a / 2, a - m), At(-a / 2, a - m));
        d.Polygon(d.Outline("Octagon", "#FF8EC00A", 3), octagon);
        string text = d.Text("GreenMark", "#FF2FA808", 30);
        d.Label("8", d.Pt(k * (0 - (m - a / 2)), k * (m - a / 2)), text, 30);
        d.Label("8", d.Pt(k * (m - a / 2), -k * (m - a / 2)), text, 30);
        d.Scene(-radius, radius, radius, -radius);
    }

    // A square and a circle of radius 2 touching its left and bottom sides and through its top right
    // corner; two triangles from the bottom left corner to where the circle crosses the top and the right
    // sides. Their total area? Nothing to drag: the circle fixes the square (side 2 + √2).
    [Puzzle("square-overlapping-a-circle")]
    public static void SquareOverlappingACircle(Lgf d)
    {
        d.Texture = "FeltTip";
        double s = 2 + Math.Sqrt(2), c = 2 - Math.Sqrt(2);
        var bl = d.Pt(0, 0);
        string orange = d.Fill("Orange", "#FFFCCC0A", "#FFF07D18");
        d.Polygon(orange, bl, d.Pt(0, s), d.Pt(c, s));
        d.Polygon(orange, bl, d.Pt(s, 0), d.Pt(s, c));
        d.Box(d.Outline("Purple", "#FF8520A8"), V.At(0, 0), V.At(s, s));
        d.Circle(d.Outline("Red", "#FFDD1008", 3.5), d.Pt(2, 2), d.Pt(4, 2));
        d.Scene(0, 4, 4, 0);
    }

    // A regular hexagon in a circle and a semicircle on each side, inward. The shaded parts are where
    // neighboring semicircles overlap and the middle that none covers. What fraction of the circle?
    // Nothing to drag: the hexagon is regular.
    [Puzzle("six-semi-circles-in-a-circle")]
    public static void SixSemiCirclesInACircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4;
        var o = d.Pt(0, 0, "O");
        var corners = Enumerable.Range(0, 6).Select(i => d.Pt(R * Math.Cos(i * Math.PI / 3), R * Math.Sin(i * Math.PI / 3))).ToArray();
        var middles = Enumerable.Range(0, 6).Select(i => Mid(d, corners[i], corners[(i + 1) % 6])).ToArray();
        var halves = Enumerable.Range(0, 6).Select(i => d.Arc("CircleSegment", null, middles[i], corners[(i + 1) % 6], corners[i], visible: false)).ToArray();
        string hexagon = d.Polygon(null, corners);
        string yellow = d.Fill("Yellow", "#FFFED702", null);
        d.Region(yellow, "Difference", [hexagon, .. halves]);
        for (int i = 0; i < 6; i++)
        {
            d.Region(yellow, "Intersection", halves[i], halves[(i + 1) % 6]);
        }

        d.Polygon(d.Outline("Cyan", "#FF28C8E0"), corners);
        string red = d.Line("Red", "#FFE50A06", 3.5, dark: "#FFFF5A4F");
        for (int i = 0; i < 6; i++)
        {
            d.Arc("CircleArc", red, middles[i], corners[(i + 1) % 6], corners[i]);
        }

        d.Circle(d.Outline("Blue", "#FF0600DE", 4), o, corners[0]);
        d.Scene(-R, R, R, -R);
    }

    // A regular hexagon of area 30 and a rectangle with a corner at its top left corner, the next one on
    // its bottom side, and the far side through its top right corner. Its area? Dragging the corner along
    // the bottom side tilts the rectangle: it is twice the triangle on the top side up to that corner,
    // and stays 20.
    [Puzzle("rectangle-on-a-hexagon")]
    public static void RectangleOnAHexagon(Lgf d)
    {
        double h = Math.Sqrt(3);
        var hexagon = HexagonFromLeft(d, V.At(-2, 0), "2");
        var tl = d.Pt(-1, h, "TL");
        var p = d.Slider(V.At(-0.75, -h), V.At(0.95, -h), (0.24 + 0.75) / 1.7, (0.24, -h), "P");
        string length = d.Num("distance(TL, P)", "len");
        string nx = d.Num($"(TL.Y - P.Y) / {length}", "nx");
        string ny = d.Num($"(P.X - TL.X) / {length}", "ny");
        // the distance from the top right corner to the line from TL to P
        var across = new V($"2 * {nx} ^ 2", $"2 * {nx} * {ny}");
        d.Polygon(d.Fill("Pink", "#FFF7B0F0", "#FFF200B4", 3), tl, p, d.Pt(p + across), d.Pt(tl + across));
        d.Polygon(d.Outline("Blue", "#FF0A02D8", 3), hexagon);
        d.Scene(-2, 2.6, 2.4, -h);
    }

    // Two squares side by side on a line, and a circle of radius 4 through the small one's top left
    // corner, the corner they share at the bottom and the big one's top right corner. Their combined
    // area? Dragging the small square's corner along the circle trades size between the squares: the
    // two outer corners are ends of a diameter (the angle at the shared corner is right), and the areas
    // add up to 32.
    [Puzzle("two-squares-on-a-circle")]
    public static void TwoSquaresOnACircle(Lgf d)
    {
        const double R = 4;
        var c = d.Pt(0, 0, "C");
        var a = d.ArcSlider(c, R, 186, 219, 203.7, (0, 0), "A");
        string small = d.Num("A.Y - A.X", "a");
        string big = d.Num("-A.X - A.Y", "b");
        var o = d.Pt(new V($"A.X + {small}", $"A.Y - {small}"), "O");
        d.Box(d.Fill("Yellow", "#FFFBDA4F", "#FFDD2A05"), new V("A.X", "O.Y"), new V("O.X", "A.Y"));
        d.Box(d.Fill("Pink", "#FFF8A0BB", "#FFDD3874"), o, new V($"O.X + {big}", $"O.Y + {big}"));
        d.Circle(d.Outline("Blue", "#FF0AA8DC", 3), c, d.Pt(R, 0));
        d.Scene(-R, R, R, -R);
    }

    // A regular pentagon and two semicircles through its top corner, each touching one of the
    // pentagon's sides there, their bases (diameters) sharing an end. The angle between the bases?
    // Dragging that shared end resizes both semicircles (their centers stay on the normals to the
    // sides at the corner): the angle stays 72°, the angle between those normals.
    [Puzzle("pentagon-with-two-semi-circles")]
    public static void PentagonWithTwoSemiCircles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double S = 2.5;
        V Toward(V from, double degrees, string name = null) => d.Pt(from + V.At(S * Math.Cos(degrees * Math.PI / 180), S * Math.Sin(degrees * Math.PI / 180)), name);
        var t = d.Pt(0, 0, "T");
        var l1 = Toward(t, 216);
        var l2 = Toward(l1, 288);
        var r2 = Toward(l2, 0);
        var r1 = Toward(r2, 72);
        double gx = Math.Cos(126 * Math.PI / 180), gy = Math.Sin(126 * Math.PI / 180);
        double bx = Math.Cos(54 * Math.PI / 180), by = Math.Sin(54 * Math.PI / 180);
        var p = d.HandleIn(d.Circle(null, d.Pt(0.74, 4.76), d.Pt(0.74, 4.16)), 0.74, 4.76, "P");
        string rg = d.Num($"(P.X ^ 2 + P.Y ^ 2) / (2 * ({Lgf.N(gx)} * P.X + {Lgf.N(gy)} * P.Y))", "rg");
        string rb = d.Num($"(P.X ^ 2 + P.Y ^ 2) / (2 * ({Lgf.N(bx)} * P.X + {Lgf.N(by)} * P.Y))", "rb");
        var cg = d.Pt(new V($"{Lgf.N(gx)} * {rg}", $"{Lgf.N(gy)} * {rg}"), "Cg");
        var cb = d.Pt(new V($"{Lgf.N(bx)} * {rb}", $"{Lgf.N(by)} * {rb}"), "Cb");
        var g = d.Pt(2 * cg - p, "G");
        var b = d.Pt(2 * cb - p, "B");
        d.Arc("CircleSegment", d.Outline("Green", "#FF94C82A", 3), cg, g, p);
        d.Arc("CircleSegment", d.Outline("Blue", "#FF1AAAC6", 3), cb, p, b);
        d.Inked(d.Outline("Orange", "#FFCC3A02", 3), t, l1, l2, r2, r1);
        string red = "#FFC00808";
        d.ArcMark(d.Line("Red", red, 3), p, g, b, 0.55);
        string toG = d.Num("distance(P, G)");
        string toB = d.Num("distance(P, B)");
        var bisector = new V($"(G.X - P.X) / {toG} + (B.X - P.X) / {toB}", $"(G.Y - P.Y) / {toG} + (B.Y - P.Y) / {toB}");
        string length = d.Num($"sqrt(({bisector.X}) ^ 2 + ({bisector.Y}) ^ 2)");
        d.Label("?", p + new V($"1.05 * ({bisector.X}) / {length}", $"1.05 * ({bisector.Y}) / {length}"), d.Text("RedMark", red, 26));
        d.Scene(-4.8, 5.1, 2.4, -3.85);
    }

    // A regular hexagon cut into twelve horizontal stripes of equal height, colored blue, red, white from
    // the top down. More red or blue, and by how much? Nothing to drag.
    [Puzzle("striped-hexagon")]
    public static void StripedHexagon(Lgf d)
    {
        d.Texture = "FeltTip";
        const double A = 4;
        double h = A * Math.Sqrt(3) / 2, step = 2 * h / 12;
        double Half(double y) => A - Math.Abs(y) / Math.Sqrt(3);
        string blue = d.Fill("Blue", "#FF06C3B7", null);
        string red = d.Fill("Red", "#FFB80140", null);
        string ink = d.Line("Ink", Lgf.Ink, 3);
        for (int i = 0; i < 12; i++)
        {
            double top = h - i * step, bottom = top - step;
            if (i % 3 != 2)
            {
                d.Polygon(i % 3 == 0 ? blue : red, d.Pt(-Half(bottom), bottom), d.Pt(Half(bottom), bottom), d.Pt(Half(top), top), d.Pt(-Half(top), top));
            }

            if (i > 0)
            {
                d.Segment(ink, d.Pt(-Half(top), top), d.Pt(Half(top), top));
            }
        }

        d.Inked(d.Outline("Frame", Lgf.Ink, 3), HexagonFromLeft(d, V.At(-A, 0), Lgf.N(A)));
        d.Scene(-A, h, A, -h);
    }

    // A semicircle, a regular hexagon standing on its diameter with its top corners on the arc, and the
    // circle inscribed in the hexagon. What fraction of the semicircle is shaded? Nothing to drag: the
    // hexagon's side is 2R/√13.
    [Puzzle("circle-in-a-hexagon-in-a-semi-circle")]
    public static void CircleInAHexagonInASemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4;
        double a = 2 * R / Math.Sqrt(13), h = a * Math.Sqrt(3) / 2;
        var o = d.Pt(0, 0, "O");
        var hexagon = HexagonFromLeft(d, V.At(-a, h), Lgf.N(a));
        d.Polygon(d.Outline("Orange", "#FFF8843F", 3), hexagon);
        d.Arc("CircleSegment", d.Outline("Purple", "#FF5E0E98", 3), o, d.Pt(R, 0), d.Pt(-R, 0));
        d.Circle(d.Fill("Red", "#FFEB2050", "#FF4A1020", 3.5), d.Pt(0, h, "C"), d.Pt(h, h));
        d.Scene(-R, R, R, 0);
    }

    // A semicircle and a rectangle on its diameter line, its top side touching the arc and its bottom
    // right corner at the diameter's right end; from there to where the rectangle's left side crosses
    // the arc is 6. The rectangle's area? Dragging the diameter's left end resizes the semicircle and
    // the rectangle with it: 6 squared is twice the radius times the width, so the area stays 18.
    [Puzzle("overlapping-rectangle-and-semi-circle")]
    public static void OverlappingRectangleAndSemiCircle(Lgf d)
    {
        var e = d.Pt(0, 0, "E");
        var w = d.Slider(V.At(-6.1, 0), V.At(-7.3, 0), (6.36 - 6.1) / 1.2, (-6.36, 0), "W");
        string r = d.Num("-W.X / 2", "r");
        string width = d.Num($"18 / {r}", "w");
        var o = d.Pt(new V($"-{r}", "0"), "O");
        var x = d.Pt(new V($"-{width}", $"sqrt(36 - {width} ^ 2)"), "X");
        d.Arc("CircleSegment", d.Outline("Yellow", "#FFC8DC6A", 3), o, e, w);
        d.Box(d.Outline("Green", "#FF5EBE8F", 3), new V($"-{width}", "0"), new V("0", r));
        string navy = "#FF1505C4";
        d.Dimension(d.Line("Arrow", navy, 3), d.Text("NavyMark", navy, 30), x, e, 0, "6", 0.45);
        d.Scene(-6.9, 3.4, 0, -0.1);
    }
}
