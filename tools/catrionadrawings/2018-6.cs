#nullable disable

static partial class Puzzles
{
    // A regular hexagon with twelve equally spaced dots on it (its corners and the middles of its
    // sides) and three orange spikes between dots. What fraction is shaded? Nothing to drag: the dots
    // fix it.
    [Puzzle("shaded-spikes-in-a-hexagon")]
    public static void ShadedSpikesInAHexagon(Lgf d)
    {
        d.Texture = "FeltTip";
        var h = HexagonFromLeft(d, V.At(-4, 0), "4");
        var (l, tl, tr, r, br, bl) = (h[0], h[1], h[2], h[3], h[4], h[5]);
        var upperLeft = Mid(d, l, tl);
        var top = Mid(d, tl, tr);
        var upperRight = Mid(d, tr, r);
        var lowerRight = Mid(d, r, br);
        var bottom = Mid(d, br, bl);
        var lowerLeft = Mid(d, bl, l);
        string orange = d.Fill("Orange", "#FFEE6F2A", "#FFE0601E", 2);
        d.Polygon(orange, l, upperLeft, top);
        d.Polygon(orange, l, upperRight, lowerLeft);
        d.Polygon(orange, lowerLeft, lowerRight, bl);
        d.Polygon(d.Outline("Frame", Lgf.Ink, 3.5), h);
        foreach (var p in new[] { l, upperLeft, tl, top, tr, upperRight, r, lowerRight, br, bottom, bl, lowerLeft })
        {
            d.DotAt(p, size: 14);
        }

        d.Scene(-4, 4 * S3, 4, -4 * S3);
    }

    // A regular octagon, two squares on its long horizontal diagonal (dashed), each with a corner at
    // an end of it and its top and bottom corners on the pink lines between the octagon's corners,
    // which are 8 long. The perimeter of the shaded octagon the squares make? Nothing to drag: the
    // octagon fixes the squares.
    [Puzzle("squares-in-an-octagon")]
    public static void SquaresInAnOctagon(Lgf d)
    {
        double radius = 4 * Math.Sqrt(2);
        var octagon = Enumerable.Range(0, 8).Select(i => d.Pt(radius * Math.Cos(i * Math.PI / 4), radius * Math.Sin(i * Math.PI / 4))).ToArray();
        var east = octagon[0];
        var west = octagon[4];
        var left = new[] { west, d.Pt(4 - radius, 4), d.Pt(8 - radius, 0), d.Pt(4 - radius, -4) };
        var right = new[] { east, d.Pt(radius - 4, -4), d.Pt(radius - 8, 0), d.Pt(radius - 4, 4) };
        string leftShape = d.Polygon(null, left);
        string rightShape = d.Polygon(null, right);
        d.Region(d.Fill("Blue", "#FFB8C8DC", null), "Union", leftShape, rightShape);
        string gray = d.Outline("Square", "#FF6E7484", 2);
        d.Polygon(gray, left);
        d.Polygon(gray, right);
        d.Polygon(d.Outline("Frame", Lgf.Ink, 3), octagon);
        string pink = d.Line("Pink", "#FFE0559F", 3);
        d.Segment(pink, octagon[3], octagon[1]);
        d.Segment(pink, octagon[5], octagon[7]);
        d.Dashed(d.Line("Dashes", "#FF2E8FC0", 3), west, east, 14);
        foreach (var p in octagon.Concat(left.Skip(1)).Concat(right.Skip(1)))
        {
            d.DotAt(p, size: 7);
        }

        d.Scene(-radius, radius, radius, -radius);
    }

    // A square of side 6; two pink squares against its top and bottom sides, two blue squares on their
    // points between its left and right sides, meeting in the middle, the pink ones' corners on the
    // blue ones' sides. The total shaded area? Nothing to drag: the side fixes them all.
    [Puzzle("four-squares-in-a-big-square")]
    public static void FourSquaresInABigSquare(Lgf d)
    {
        string pink = d.Fill("Pink", "#FFF0B8AA");
        d.Box(pink, V.At(2, 4), V.At(4, 6));
        d.Box(pink, V.At(2, 0), V.At(4, 2));
        string blue = d.Fill("Blue", "#FFACD0D2");
        foreach (double x in new[] { 1.5, 4.5 })
        {
            d.Polygon(blue, d.Pt(x - 1.5, 3), d.Pt(x, 1.5), d.Pt(x + 1.5, 3), d.Pt(x, 4.5));
        }

        d.Box(d.Outline("Frame"), V.At(0, 0), V.At(6, 6));
        d.Dimension(d.Line("Arrow"), d.Text(), d.Pt(6, 0), d.Pt(0, 0), 0.6, "6", 0.45);
        // room under the arrow for the page's buttons
        d.Scene(0, 6, 6, -1.5);
    }

    // Three unit squares in a row (pink), a circle in each, and in the circles a square, two stacked
    // squares and three stacked squares, each stack with its four corners on the circle; the row goes
    // on. Nothing to drag: each stack is fixed by its circle.
    [Puzzle("stacked-boxes-in-circles-in-squares")]
    public static void StackedBoxesInCirclesInSquares(Lgf d)
    {
        const double L = 4;
        string pink = d.Line("Pink", "#FFE04A94", 3);
        d.Segment(pink, d.Pt(-0.12, L), d.Pt(3 * L + 0.4, L));
        d.Segment(pink, d.Pt(-0.12, 0), d.Pt(3 * L + 0.4, 0));
        for (int i = 0; i <= 3; i++)
        {
            d.Segment(pink, d.Pt(i * L, -0.08), d.Pt(i * L, L + 0.08));
        }

        string teal = d.Outline("Circle", "#FF2EA3B5", 4);
        string navy = d.Outline("Box", "#FF2D3FA6", 3);
        string navyLine = d.Line("BoxLine", "#FF2D3FA6", 3);
        for (int n = 1; n <= 3; n++)
        {
            double cx = (n - 0.5) * L, cy = L / 2;
            d.Circle(teal, d.Pt(cx, cy), d.Pt(cx + L / 2, cy));
            double a = L / Math.Sqrt(1 + n * n);
            d.Box(navy, V.At(cx - a / 2, cy - n * a / 2), V.At(cx + a / 2, cy + n * a / 2));
            for (int k = 1; k < n; k++)
            {
                double y = cy - n * a / 2 + k * a;
                d.Segment(navyLine, d.Pt(cx - a / 2, y), d.Pt(cx + a / 2, y));
            }
        }

        string dot = d.Wedge("PinkDot", "#FFE04A94");
        for (int k = 0; k < 3; k++)
        {
            double x = 3 * L + 0.75 + 0.42 * k;
            d.Circle(dot, d.Pt(x, L / 2 + 0.3), d.Pt(x + 0.1, L / 2 + 0.3));
        }

        d.Scene(-0.12, L, 3 * L + 1.7, 0);
    }

    // A circle in a square and four identical squares in it: three tilted ones stacked, the top one
    // with its top corners on the circle, the bottom one with its bottom corners on it (so they line
    // up, and the stack is three squares tall), the middle one pushed sideways; the fourth upright,
    // in the gap, a corner on the circle. What fraction of the big square is shaded? The middle
    // square is dragged along the stack: the squares stay, so does the fraction (2/5).
    [Puzzle("four-stacked-squares-in-a-circle-in-a-square")]
    public static void FourStackedSquaresInACircleInASquare(Lgf d)
    {
        const double R = 5;
        double s = 2 * R / Math.Sqrt(10);
        double angle = 21 * Math.PI / 180, cos = Math.Cos(angle), sin = Math.Sin(angle);
        // along the stack's squares (u) and up the stack (w)
        (double X, double Y) Off(double u, double w) => (u * cos - w * sin, u * sin + w * cos);
        V At(double u, double w) => V.At(Off(u, w).X, Off(u, w).Y);
        // the upright square: its right side through the bottom square's top left corner, its bottom
        // left corner on the circle, its top right corner on the middle square's left side
        double right = Off(-s / 2, -s / 2).X;
        double bottom = -Math.Sqrt(R * R - (right - s) * (right - s));
        double home = right * cos + (bottom + s) * sin + s / 2;
        double end = 2.6;
        var o = d.Pt(0, 0, "O");
        d.Box(d.Outline("Frame", "#FF3EB0D4", 3.5), V.At(-R, -R), V.At(R, R));
        d.Circle(d.Outline("Circle", "#FF3887D8", 4), o, d.Pt(R, 0));
        var b = d.Slider(At(home, 0), At(end, 0), 0, Off(home, 0), "B");
        string pink = d.Fill("Pink", "#FFF0C2D6", "#FFE87A50", 4, darkOutline: "#FFFFA27A");
        d.Box(pink, V.At(right - s, bottom), V.At(right, bottom + s));
        d.Polygon(pink, d.Pt(At(-s / 2, s / 2)), d.Pt(At(s / 2, s / 2)), d.Pt(At(s / 2, 3 * s / 2)), d.Pt(At(-s / 2, 3 * s / 2)));
        d.Polygon(pink, d.Pt(b + At(-s / 2, -s / 2)), d.Pt(b + At(s / 2, -s / 2)), d.Pt(b + At(s / 2, s / 2)), d.Pt(b + At(-s / 2, s / 2)));
        d.Polygon(pink, d.Pt(At(-s / 2, -3 * s / 2)), d.Pt(At(s / 2, -3 * s / 2)), d.Pt(At(s / 2, -s / 2)), d.Pt(At(-s / 2, -s / 2)));
        d.Scene(-R, R, R, -R);
    }

    // A square and a circle through its two right corners, touching its left side; the circle reaches
    // 1 past the square's right side. The square's area? Nothing to drag: the 1 fixes it (side 4,
    // radius 5/2).
    [Puzzle("overlapping-square-and-circle")]
    public static void OverlappingSquareAndCircle(Lgf d)
    {
        d.Box(d.Outline("Square", "#FF0074B1", 3.5), V.At(0, 0), V.At(4, 4));
        d.Circle(d.Outline("Circle", "#FFD5604F", 3.5), d.Pt(2.5, 2), d.Pt(5, 2));
        string teal = "#FF02A6C1";
        d.DoubleArrow(d.Line("Arrow", teal, 3), d.Text("TealMark", teal, 26), d.Pt(4, 2), d.Pt(5, 2), "1", V.At(4.5, 2.4));
        d.Scene(0, 4.5, 5, -0.5);
    }

    // A semicircle with center O; an equilateral triangle on the right half of its diameter, its top
    // on the arc, with a circle of area 8 in it; a chord from the left end to that top, and the other
    // circle in the part of the semicircle above the chord. Its area? Nothing to drag: the triangle
    // fixes everything.
    [Puzzle("two-circles-inscribed-in-triangles-in-a-semi-circle")]
    public static void TwoCirclesInscribedInTrianglesInASemiCircle(Lgf d)
    {
        d.Texture = null;
        const double R = 4;
        double root3 = Math.Sqrt(3);
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0, "E");
        var west = d.Pt(-R, 0, "W");
        var apex = d.Pt(R / 2, R * root3 / 2, "A");
        d.Arc("CircleArc", d.Line("Arc", "#FF434343", 3), o, east, west);
        string purple = d.Line("Purple", "#FF8E7CC3", 3);
        d.Segment(purple, west, east);
        d.Segment(purple, o, apex);
        d.Segment(purple, apex, east);
        d.Segment(d.Line("Blue", "#FF6FA8DC", 3), west, apex);
        double r = R / (2 * root3);
        var red = d.Pt(R / 2, r, "Cr");
        d.Circle(d.Outline("Red", "#FFE06666", 3), red, d.Pt(R / 2 + r, r));
        // the chord is R / 2 from O, the arc R beyond: the circle between them, on the line from O
        // through the chord's middle, has radius R / 4
        var orange = d.Pt(-3 * R / 8, 3 * R / 4 * S3, "Co");
        d.Circle(d.Outline("Orange", "#FFF6B26B", 3), orange, d.Pt(-3 * R / 8 + R / 4, 3 * R / 4 * S3));
        d.Label("8", red, d.Text("RedMark", "#FFE06666", 44), 44);
        d.Label("?", orange, d.Text("OrangeMark", "#FFF6B26B", 44), 44);
        d.DotAt(o, size: 14);
        d.Label("O", V.At(0, -0.6), d.Text("OMark", "#FF434343", 30), 30);
        d.Scene(-R, R, R, -0.85);
    }

    // A circle, three dots equally spaced on it, an arc centered on each dot (in its color) between the
    // other two, and three arcs through pairs of dots crossing at the center. The pink parts between
    // them: what fraction of the circle? Nothing to drag: the dots fix it.
    [Puzzle("arcs-in-a-circle")]
    public static void ArcsInACircle(Lgf d)
    {
        const double R = 3;
        double x = R * S3, y = R / 2;
        var o = d.Pt(0, 0, "O");
        var t = d.Pt(0, R, "T");
        var br = d.Pt(x, -y, "Br");
        var bl = d.Pt(-x, -y, "Bl");
        // the centers of the arcs through the center: the circle's points opposite the dots
        var k1 = d.Pt(-x, y, "K1");
        var k2 = d.Pt(x, y, "K2");
        var k3 = d.Pt(0, -R, "K3");
        string c1 = d.Circle(null, k1, t);
        string c2 = d.Circle(null, k2, t);
        string c3 = d.Circle(null, k3, bl);
        string pinkDisk = d.Circle(null, br, t);
        string navyDisk = d.Circle(null, bl, t);
        string tealDisk = d.Circle(null, t, bl);
        string fill = d.Fill("Pink", "#FFF8D8DF", null);
        d.Region(fill, "Difference", d.Region(null, "Intersection", pinkDisk, c1), c2, c3);
        d.Region(fill, "Difference", d.Region(null, "Intersection", navyDisk, c2), c1, c3);
        d.Region(fill, "Difference", d.Region(null, "Intersection", tealDisk, c3), c1, c2);
        d.Circle(d.Outline("Circle", "#FF5A6670", 3), o, d.Pt(R, 0));
        string gray = d.Line("Petal", "#FFC9C3B3", 3);
        d.Arc("CircleArc", gray, k1, bl, t);
        d.Arc("CircleArc", gray, k2, t, br);
        d.Arc("CircleArc", gray, k3, br, bl);
        string teal = "#FF4FB3C0", pink = "#FFE85F95", navy = "#FF1E40B0";
        d.Arc("CircleArc", d.Line("Teal", teal, 4), t, bl, br);
        d.Arc("CircleArc", d.Line("PinkArc", pink, 4), br, t, bl);
        d.Arc("CircleArc", d.Line("Navy", navy, 4), bl, br, t);
        const double dot = 0.13;
        d.Circle(d.Wedge("TealDot", "#FF3E8FAD"), t, d.Pt(dot, R));
        d.Circle(d.Wedge("PinkDot", "#FFD6327B"), br, d.Pt(x + dot, -y));
        d.Circle(d.Wedge("NavyDot", "#FF0C3AAA"), bl, d.Pt(-x + dot, -y));
        d.DotAt(o, size: 7);
        d.Scene(-R, R, R, -R);
    }

    // A regular hexagon, six identical squares each with a corner at one of its corners and its
    // diagonal along the radius, neighbours sharing a corner, and a rectangle between four of their
    // inner corners. What fraction of the hexagon do they cover? Nothing to drag: the hexagon fixes it.
    [Puzzle("six-squares-in-a-hexagon")]
    public static void SixSquaresInAHexagon(Lgf d)
    {
        const double R = 5;
        double g = R / (1 + Math.Sqrt(3)), inner = R - 2 * g;
        (double X, double Y) Polar(double radius, double degrees) => (radius * Math.Cos(degrees * Math.PI / 180), radius * Math.Sin(degrees * Math.PI / 180));
        var corners = Enumerable.Range(0, 6).Select(i => d.Pt(Polar(R, 60 * i).X, Polar(R, 60 * i).Y)).ToArray();
        d.Polygon(d.Outline("Hexagon", "#FF4A9CBE", 3.5), corners);
        string square = d.Outline("Square", "#FF2A2840", 3.5);
        for (int i = 0; i < 6; i++)
        {
            double a = 60 * i * Math.PI / 180, cos = Math.Cos(a), sin = Math.Sin(a);
            double px = R * cos, py = R * sin;
            d.Polygon(square, corners[i],
                d.Pt(px - g * cos - g * sin, py - g * sin + g * cos),
                d.Pt(px - 2 * g * cos, py - 2 * g * sin),
                d.Pt(px - g * cos + g * sin, py - g * sin - g * cos));
        }

        var rectangle = new[] { 0, 60, 180, 240 }.Select(a => d.Pt(Polar(inner, a).X, Polar(inner, a).Y)).ToArray();
        d.Polygon(d.Outline("Rectangle", "#FF2E6CC8", 3.5), rectangle);
        d.Scene(-R, R * S3, R, -R * S3);
    }

    // A square of area 12 and an equilateral triangle on each side; the triangles' tips make a bigger
    // square, and the four gaps between it and the triangles are shaded. Their total area? Nothing to
    // drag: the square fixes it.
    [Puzzle("four-equilateral-triangles-round-a-square")]
    public static void FourEquilateralTrianglesRoundASquare(Lgf d)
    {
        double half = Math.Sqrt(3), tip = half + 3;
        var east = d.Pt(tip, 0);
        var north = d.Pt(0, tip);
        var west = d.Pt(-tip, 0);
        var south = d.Pt(0, -tip);
        var ne = d.Pt(half, half);
        var nw = d.Pt(-half, half);
        var sw = d.Pt(-half, -half);
        var se = d.Pt(half, -half);
        string teal = d.Fill("Teal", "#FFC5E1EB", null);
        d.Polygon(teal, east, north, ne);
        d.Polygon(teal, north, west, nw);
        d.Polygon(teal, west, south, sw);
        d.Polygon(teal, south, east, se);
        d.Polygon(d.Outline("Tips", "#FF1FA8D0", 3.5), east, north, west, south);
        string blue = d.Outline("Blue", "#FF0745B8", 3.5);
        d.Polygon(blue, nw, ne, north);
        d.Polygon(blue, ne, se, east);
        d.Polygon(blue, se, sw, south);
        d.Polygon(blue, sw, nw, west);
        d.Polygon(blue, ne, nw, sw, se);
        d.Label("12", V.At(0, 0), d.Text("TealMark", "#FF1D9DC2", 36), 36);
        d.Scene(-tip, tip, tip, -tip);
    }

    // A rectangle, a point on its diagonal straight above the point of its bottom side as far from the
    // corner as the rectangle is high (the red ticks), and lines from the point to the other corners.
    // The yellow and the orange triangles: the same share of the rectangle? Dragging the top right
    // corner sideways changes the rectangle's shape: the two stay equal.
    [Puzzle("two-triangles-in-a-rectangle")]
    public static void TwoTrianglesInARectangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double H = 4;
        var bl = d.Pt(0, 0, "Bl");
        var tl = d.Pt(0, H, "Tl");
        var tr = d.Slider(V.At(4.6, H), V.At(6.8, H), (5.2 - 4.6) / 2.2, (5.2, H), "W");
        var br = d.Pt(new V("W.X", "0"), "Br");
        var p = d.Pt(new V(Lgf.N(H), $"{Lgf.N(H * H)} / W.X"), "P");
        var foot = d.Pt(H, 0, "F");
        d.Polygon(d.Fill("Yellow", "#FFF2C53A", null), tl, tr, p);
        d.Polygon(d.Fill("Orange", "#FFF0845F", null), p, tr, br);
        // across the yellow and the orange, bright in the dark theme: a mid gray there
        string gray = d.Line("Gray", "#FF5A6070", 2.5, dark: "#FF8C929E");
        d.Segment(gray, bl, tr);
        d.Segment(gray, tl, p);
        d.Segment(gray, p, br);
        d.Segment(gray, p, foot);
        d.Angle(gray, foot, bl, p, 14);
        string red = d.Line("Red", "#FFC8201A", 3);
        d.Segment(red, bl, tl, "TwoTicks");
        d.Segment(red, bl, foot, "TwoTicks");
        string ink = d.Line("Ink", Lgf.Ink, 3);
        d.Segment(ink, tl, tr);
        d.Segment(ink, tr, br);
        d.Segment(ink, foot, br);
        d.DotAt(p);
        d.Scene(0, H, 6.8, 0);
    }

    // Two squares sharing a corner, the tilted one with a corner on the upright one's side; the
    // upright one's top side drawn on to the tilted one's corner, and the pink triangle from there to
    // the tilted one's next corner and the upright one's corner has area 5. More green or blue, by how
    // much? Dragging the upright square's corner resizes it, the tilt following: the pink triangle
    // keeps its area, and blue stays 10 more than green.
    [Puzzle("two-overlapping-squares-iii")]
    public static void TwoOverlappingSquaresIII(Lgf d)
    {
        d.Texture = "FeltTip";
        string h = Lgf.N(Math.Sqrt(10));
        var o = d.Pt(0, 0, "D");
        var s = d.Slider(V.At(6.5, 6.5), V.At(10.2, 10.2), (9.4 - 6.5) / 3.7, (9.4, 9.4), "S");
        var tl = d.Pt(new V("0", "S.Y"), "Tl");
        var br = d.Pt(new V("S.X", "0"), "Br");
        var c = d.Pt(new V("S.X", h), "C");
        var b = d.Pt(new V($"S.X - {h}", $"S.Y + {h}"), "B");
        var a = d.Pt(new V($"-{h}", "S.Y"), "A");
        string upright = d.Polygon(null, o, br, s, tl);
        string tilted = d.Polygon(null, o, c, b, a);
        string pinkShape = d.Polygon(d.Fill("Pink", "#FFCF0170", null), a, b, tl);
        d.Region(d.Fill("Green", "#FF6EDC46", null), "Difference", upright, tilted);
        d.Region(d.Fill("Blue", "#FF0A3599", null), "Difference", tilted, upright, pinkShape);
        string frame = d.Outline("Frame", Lgf.Ink, 3.5);
        d.Polygon(frame, o, br, s, tl);
        d.Polygon(frame, o, c, b, a);
        d.Segment(d.Line("Ink", Lgf.Ink, 3.5), a, tl);
        string magenta = "#FFCF0170";
        var tip = d.Pt(a + 0.23 * (b - a) + V.At(0, 0.4));
        d.Vector(d.Line("PinkArrow", magenta, 3), d.Pt(tip + V.At(0, 1.3)), tip);
        d.Label("Area = 5", a + 0.4 * (b - a) + V.At(0, 2.6), d.Text("PinkMark", magenta, 34), 34);
        // room under the squares for the page's buttons
        d.Scene(-3.2, 14.4, 9.4, -0.9);
    }
}
