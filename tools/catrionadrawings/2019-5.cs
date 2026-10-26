#nullable disable

static partial class Puzzles
{
    // Three identical circles in a semicircle: two on the diameter touching the arc, one touching
    // the arc at the top and both of them. What fraction of the semicircle isn't covered? Nothing to
    // drag: the touching fixes the circles at a third of the radius, the outer ones' centers 1/√3 of
    // it from the middle.
    [Puzzle("three-circles-in-a-semi-circle")]
    public static void ThreeCirclesInASemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 3, r = R / 3;
        double a = R / Math.Sqrt(3);
        var o = d.Pt(0, 0);
        d.Arc("CircleSegment", d.Fill("Yellow", "#FFFEFEA5", "#FF0896E8", 3.5), o, d.Pt(R, 0), d.Pt(-R, 0));
        string circle = d.PaperFill("Circle", null);
        // on the yellow, which stays bright in the dark theme: a vivid orange there
        string rim = d.Outline("Rim", "#FFEC8A26", 3, dark: "#FFF07A1A");
        foreach (var (x, y) in new[] { (-a, r), (0, 2 * r), (a, r) })
        {
            var center = d.Pt(x, y);
            var through = d.Pt(x + r, y);
            d.Circle(circle, center, through);
            d.Circle(rim, center, through);
        }

        d.Scene(-R, R, R, 0);
    }

    // Three semicircles side by side, in each a stack of identical upside-down semicircles (one, two,
    // three), the top one's diameter a chord of the big arc, each standing on the one below's
    // diameter, the bottom one touching the base. What proportion is shaded? Nothing to drag: the
    // stack of n fixes the small radius at R / √(1 + n²). Her three big ones' sizes are kept.
    [Puzzle("upside-down-semi-circles-in-a-semi-circle")]
    public static void UpsideDownSemiCirclesInASemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        string[] fills =
        [
            d.Fill("Red", "#FFDC0303", "#FFC80000", 2),
            d.Fill("Orange", "#FFFE9504", "#FFF07800", 2),
            d.Fill("Yellow", "#FFFEFD0D", "#FFF0E000", 2),
        ];
        // her centers and radii, in her pixels / 25
        var diagrams = new[] { (X: 3.54, R: 2.46), (X: 11.6, R: 3.7), (X: 21.4, R: 4.7) };
        for (int n = 1; n <= 3; n++)
        {
            var (cx, big) = diagrams[n - 1];
            d.Arc("CircleSegment", frame, d.Pt(cx, 0), d.Pt(cx + big, 0), d.Pt(cx - big, 0));
            double r = big / Math.Sqrt(1 + n * n);
            for (int k = 1; k <= n; k++)
            {
                double y = k * r;
                d.Arc("CircleSegment", fills[n - 1], d.Pt(cx, y), d.Pt(cx - r, y), d.Pt(cx + r, y));
            }
        }

        d.Scene(1.08, 4.7, 26.1, 0);
    }

    // A stack of triangles on a line: a yellow and a blue equilateral one side by side, the green one
    // between their tops and the corner they share, and a pink equilateral one on the green one's top
    // side. Yellow + blue = 100, pink + green? Dragging the shared corner along the line trades the
    // yellow and the blue (their sum stays, a² + b² kept): pink + green stays equal to it.
    [Puzzle("stacked-triangles")]
    public static void StackedTriangles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double a0 = 1.88, b0 = 2.96, K = a0 * a0 + b0 * b0;
        var a = d.Pt(0, 0, "A");
        var b = d.Slider(V.At(1.1, 0), V.At(3, 0), (a0 - 1.1) / 1.9, (a0, 0), "B");
        string bs = d.Num($"sqrt({Lgf.N(K)} - B.X ^ 2)", "b");
        var c = d.Pt(new V("B.X / 2", $"{s3} * B.X"), "C");
        var e = d.Pt(new V($"B.X + {bs}", "0"), "E");
        var t = d.Pt(new V($"B.X + {bs} / 2", $"{s3} * {bs}"), "D");
        var p = d.Pt(c + (t - c).Turn(60), "P");
        string teal = "#FF14CDA8";
        string yellow = d.Fill("Yellow", "#FFFEFD20", teal, 3), blue = d.Fill("Blue", "#FF1EC4F8", teal, 3);
        string green = d.Fill("Green", "#FFA6EE48", teal, 3), pink = d.Fill("Pink", "#FFFF00E0", teal, 3);
        d.Polygon(yellow, a, b, c);
        d.Polygon(blue, b, e, t);
        d.Polygon(green, c, b, t);
        d.Polygon(pink, c, t, p);
        string purple = "#FF5A10E0";
        d.Segment(d.Line("Base", purple, 3.5), d.Pt(-1.56, 0), d.Pt(new V($"{e.X} + 1.3", "0")));
        // the legend under the line, in her pixels: x = (px - 218) / 50, y = (250 - py) / 50
        var blobs = new[] { ("YellowBlob", "#FFFEFD06", 222.0, 292.0), ("BlueBlob", "#FF03A2FC", 290.0, 292.0), ("PinkBlob", "#FFF503C8", 257.0, 338.0), ("GreenBlob", "#FF59E704", 327.0, 338.0) };
        foreach (var (name, color, x, y) in blobs)
        {
            double cx = (x - 218) / 50, cy = (250 - y) / 50;
            d.Circle(d.Fill(name, color, null), d.Pt(cx, cy), d.Pt(cx + 0.27, cy));
        }

        string mark = d.Text("PurpleMark", purple, 30);
        foreach (var (text, x, y) in new[] { ("+", 257.0, 292.0), ("=", 322.0, 292.0), ("100", 373.0, 292.0), ("+", 295.0, 338.0), ("=", 360.0, 338.0), ("?", 403.0, 338.0) })
        {
            d.Label(text, V.At((x - 218) / 50, (250 - y) / 50), mark, 30);
        }

        d.Scene(-1.56, 4.2, 6.2, -2.1);
    }

    // A regular dodecagon, a corner at the top; lines from the top corner to the left one and from
    // the next corner up on the left to the right one. The orange triangle (top, crossing, right
    // corner) is 2 more than the yellow one (the two left corners, crossing): the dodecagon's area?
    // Nothing to drag: the dodecagon fixes it.
    [Puzzle("triangles-in-a-dodecagon")]
    public static void TrianglesInADodecagon(Lgf d)
    {
        const double R = 4;
        var corners = Enumerable.Range(0, 12).Select(i => d.Pt(R * Math.Cos((90 + 30 * i) * Math.PI / 180), R * Math.Sin((90 + 30 * i) * Math.PI / 180))).ToArray();
        // corners[0] the top, [2] at 150°, [3] the left one at 180°, [9] the right one at 0°
        var top = corners[0];
        var x = d.Crossing(top, corners[3], corners[2], corners[9], "X");
        d.Polygon(d.Fill("Orange", "#FFFC8C05", null), top, x, corners[9]);
        d.Polygon(d.Fill("Yellow", "#FFFDDF20", null), corners[2], x, corners[3]);
        string line = d.Line("LightBlue", "#FF0C9CF0", 2.5);
        d.Segment(line, top, corners[3]);
        d.Segment(line, corners[2], corners[9]);
        d.Segment(line, top, corners[9]);
        d.Inked(d.Outline("Frame", "#FF0513D0", 3.5), corners);
        d.Scene(-R, R, R, -R);
    }

    // Two squares the same size sharing a corner, one upright under it, one tilted above. Lines
    // from the tilted one's far corner to the upright one's bottom left and top right corners; the
    // angle between them? Dragging that corner around the shared one turns the tilted square: the
    // angle stays (all three corners are on a circle around the shared one).
    [Puzzle("angle-formed-by-two-squares-ii")]
    public static void AngleFormedByTwoSquaresII(Lgf d)
    {
        const double s = 4;
        var p = d.Pt(0, 0, "P");
        var bl = d.Pt(0, -s);
        var br = d.Pt(s, -s);
        var tr = d.Pt(s, 0);
        var t = d.ArcSlider(p, s, 35, 95, 56, (0, 0), "T");
        var w = new V("T.Y", "-T.X");
        var far = d.Pt(t + w);
        var side = d.Pt(p + w);
        string frame = d.Outline("Frame", "#FF3002F8", 4.5);
        d.Polygon(frame, bl, br, tr, p);
        d.Polygon(frame, p, t, far, side);
        string magenta = d.Line("Magenta", "#FFF305EC", 3.5);
        d.Segment(magenta, t, bl);
        d.Segment(magenta, t, tr);
        string red = "#FFF20484";
        d.ArcMark(d.Line("Red", red, 3), t, bl, tr, 0.85);
        d.Label("?", t + V.At(0.2, -1.4), d.Text("RedMark", red, 36), 36);
        d.Scene(0, 3.4, 5.5, -s);
    }

    // Three semicircles with horizontal bases: a big one, a chord across it, a semicircle on the
    // chord, and one on the big one's base at its right end, touching the chord. The yellow lune
    // (the top one outside the big circle) is 6, the blue (the big one under the chord, outside
    // the small one)? Dragging the chord's end along the big arc moves the chord up and down,
    // resizing the other two: the blue stays equal to the yellow.
    [Puzzle("three-parallel-semi-circles")]
    public static void ThreeParallelSemiCircles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4;
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0);
        var west = d.Pt(-R, 0);
        var l = d.ArcSlider(o, R, 115, 158, 138, (0, 0), "L");
        var r = d.Pt(new V("-L.X", "L.Y"), "Rc");
        var m = d.Pt(new V("0", "L.Y"), "M");
        var foot = d.Pt(new V($"{Lgf.N(R)} - 2 * L.Y", "0"), "F");
        var middle = d.Pt(new V($"{Lgf.N(R)} - L.Y", "0"), "C");
        string big = d.Arc("CircleSegment", null, o, east, west, visible: false);
        string cap = d.Arc("CircleSegment", null, o, r, l, visible: false);
        string top = d.Arc("CircleSegment", null, m, r, l, visible: false);
        string small = d.Arc("CircleSegment", null, middle, east, foot, visible: false);
        d.Region(d.Fill("Yellow", "#FFFFFF01", null), "Difference", top, cap);
        d.Region(d.Fill("Blue", "#FF03C4F8", null), "Difference", big, cap, small);
        // between the yellow, which stays bright in the dark theme, and the dark paper: a mid blue there
        string line = d.Line("Line", "#FF0A48E8", 3.5, dark: "#FF6B7BFF");
        d.Arc("CircleArc", line, o, east, west);
        d.Arc("CircleArc", line, m, r, l);
        d.Arc("CircleArc", line, middle, east, foot);
        d.Segment(line, west, east);
        d.Segment(line, l, r);
        d.Scene(-R, 2.67 + R * Math.Sin(Math.Acos(2.67 / R)), R, 0);
    }

    // A semicircle of diameter 8 on the hypotenuse of a right triangle, touching both legs; the
    // shaded triangle between the right angle and the two touching points? Dragging a touching point
    // around the semicircle turns the triangle about it: the shaded triangle is always half the
    // square of the radius.
    [Puzzle("semi-circle-in-a-right-angled-triangle")]
    public static void SemiCircleInARightAngledTriangle(Lgf d)
    {
        const double r = 4;
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(r, 0);
        var west = d.Pt(-r, 0);
        var p = d.ArcSlider(o, r, 136, 157, 154, (0, 0), "P");
        var q = d.Pt(new V("P.Y", "-P.X"), "Q");
        var c = d.Pt(p + q, "C");
        var a = d.Crossing(c, p, west, east, "A");
        var b = d.Crossing(c, q, west, east, "B");
        d.Polygon(d.Fill("Green", "#FF03D80A", null), c, p, q);
        d.Arc("CircleArc", d.Line("Arc", "#FF9EF020", 2.5), o, east, west);
        d.Inked(d.Outline("Frame", "#FF07DB86", 3.5), a, b, c);
        d.Segment(d.Line("Navy", "#FF0905EA", 3.5), p, q);
        d.DotAt(p, size: 11);
        d.DotAt(q, size: 11);
        d.Dimension(d.Line("Arrow", "#FF6EE83A", 3), d.Text("GreenMark", "#FF14C81A", 40), east, west, 0.75, "8", 0.75);
        d.Scene(-4.8, 5.36, 9.4, -1.1);
    }

    // Three identical squares side by side across a rectangle, the first in its top left corner, the
    // last in its bottom right corner, the diagonal through their inner corners. What fraction do
    // they cover? The diagonal makes the rectangle 3 by 1.5 squares. The middle square's height
    // isn't fixed: dragging it up and down, the fraction stays.
    [Puzzle("three-congruent-squares-in-a-rectangle")]
    public static void ThreeCongruentSquaresInARectangle(Lgf d)
    {
        const double s = 3, W = 3 * s, H = 1.5 * s;
        var m = d.Slider(V.At(s, 0.05 * s), V.At(s, H - s - 0.05 * s), 0.5, (s, 0.25 * s), "M");
        // the squares side by side stay bright in the dark theme: a deep gold between them there
        string yellow = d.Fill("Yellow", "#FFFEF998", "#FFF4BC05", 3, darkOutline: "#FF7A5A00");
        d.Box(yellow, V.At(0, H - s), V.At(s, H));
        d.Box(yellow, m, new V(Lgf.N(2 * s), $"M.Y + {Lgf.N(s)}"));
        d.Box(yellow, V.At(2 * s, 0), V.At(W, s));
        string red = "#FFF00401";
        d.Inked(d.Outline("Frame", red, 3.5), d.Pt(0, 0), d.Pt(W, 0), d.Pt(W, H), d.Pt(0, H));
        d.Segment(d.Line("Diagonal", red, 3.5), d.Pt(0, 0), d.Pt(W, H));
        d.Scene(0, H, W, 0);
    }

    // Three regular hexagons in a trapezium with 60° base angles: one in its top left corner, one on
    // its base under that one, and a big one in the notch between them against the right side. The
    // perimeters add up to 36; the width? The point where all three meet is dragged anywhere in
    // range: the hexagons and the trapezium's height and top change, the width stays 12 (twice the
    // sum of the sides).
    [Puzzle("three-hexagons-in-a-trapezium")]
    public static void ThreeHexagonsInATrapezium(Lgf d)
    {
        const double sum = 6, sqrt3 = 1.7320508075688772;
        // the sides a (top), b (bottom), c = 6 - a - b (big) stay where the hexagons fit: 3a + b ≥ 6
        // (c ≤ 2a, the big one under the top), a + 3b ≥ 6 (c ≤ 2b, over the base), b ≤ 2a (the bottom
        // one inside the left side), 3a + 2b ≤ 12 (a ≤ 2c, the top one inside the right side); the
        // meeting point is (2a + b, √3 b): its range is that polygon shrunk toward its middle
        (double A, double B)[] corners = [(1.5, 1.5), (24.0 / 7, 6.0 / 7), (12.0 / 7, 24.0 / 7), (1.2, 2.4)];
        double ma = corners.Average(p => p.A), mb = corners.Average(p => p.B);
        var range = corners.Select(p => (A: ma + 0.7 * (p.A - ma), B: mb + 0.7 * (p.B - mb))).Select(p => d.Pt(2 * p.A + p.B, sqrt3 * p.B)).ToArray();
        string box = d.Polygon(null, range);
        double a0 = 1.65, b0 = 1.85;
        var p = d.HandleIn(box, 2 * a0 + b0, sqrt3 * b0, "P");
        string b = d.Num($"P.Y / {Lgf.N(sqrt3)}", "b");
        string a = d.Num($"(P.X - {b}) / 2", "a");
        string c = d.Num($"{Lgf.N(sum)} - {a} - {b}", "c");
        string h = d.Num($"P.Y + {Lgf.N(sqrt3)} * {a}", "h");
        var bl = d.Pt(0, 0);
        var br = d.Pt(2 * sum, 0);
        var tr = d.Pt(new V($"{Lgf.N(2 * sum)} - {a} - {b}", h));
        var tl = d.Pt(new V($"{a} + {b}", h), "TL");
        // the top one hangs from the top left corner, the bottom one stands under the meeting point, the big one starts at it
        var top = HexagonFromLeft(d, d.Pt(new V($"TL.X - {a} / 2", $"TL.Y - {s3} * {a}")), a);
        var bottom = HexagonFromLeft(d, d.Pt(new V($"P.X - 1.5 * {b}", $"{s3} * {b}")), b);
        var big = HexagonFromLeft(d, p, c);
        string outline = "#FF0AB9EE";
        // side by side they stay light in the dark theme: a deep teal between them there
        d.Polygon(d.Fill("Blue", "#FF8FE3FB", outline, 3, darkOutline: "#FF0A5A78"), top);
        d.Polygon(d.Fill("Cyan", "#FF52F0FC", outline, 3, darkOutline: "#FF0A5A78"), bottom);
        d.Polygon(d.Fill("LightCyan", "#FFA6FEFE", outline, 3, darkOutline: "#FF0A5A78"), big);
        d.Inked(d.Outline("Frame", Lgf.Ink, 3.5), bl, br, tr, tl);
        string green = "#FF55C040";
        d.Dimension(d.Line("Arrow", green, 3), d.Text("GreenMark", green, 34), br, bl, 0.75, "?", 0.5);
        d.Scene(0, sqrt3 * (a0 + b0), 2 * sum, -1.1);
    }

    // A regular hexagon and a rectangle on its long diagonal from the top left corner to the bottom
    // right one; a line from the rectangle's top corner to the hexagon's right corner. The angle
    // between it and the rectangle's side? Dragging that corner around the hexagon's circumcircle
    // (it is on it: the diagonal is a diameter) changes the rectangle's width: the angle stays.
    [Puzzle("regular-hexagon-and-rectangle")]
    public static void RegularHexagonAndRectangle(Lgf d)
    {
        const double R = 4;
        var o = d.Pt(0, 0, "O");
        var hexagon = Enumerable.Range(0, 6).Select(i => d.Pt(R * Math.Cos(i * Math.PI / 3), R * Math.Sin(i * Math.PI / 3))).ToArray();
        // hexagon[0] the right corner, [2] the top left, [5] the bottom right
        var q = d.ArcSlider(o, R, 68, 114, 97, (0, 0), "Q");
        var opposite = d.Pt(new V("-Q.X", "-Q.Y"));
        V[] rectangle = [hexagon[2], q, hexagon[5], opposite];
        d.Polygon(d.Fill("Green", "#FFE6FE9A", null), hexagon);
        d.Polygon(d.Fill("Pink", "#FFFEDCFD", null), rectangle);
        d.Inked(d.Outline("Frame", "#FF05DB7E", 3.5), hexagon);
        d.Inked(d.Outline("Purple", "#FF7A30C0", 3), rectangle);
        string red = "#FFD00A05";
        // on the green, which stays bright in the dark theme: deep there too
        string line = d.Line("Red", red, 3.5, dark: "#FFA01010");
        d.Segment(line, q, hexagon[0]);
        d.ArcMark(line, q, hexagon[5], hexagon[0], 0.9);
        d.Label("?", q + V.At(0.85, -1.7), d.Text("RedMark", red, 36, dark: "#FFA01010"), 36);
        d.Scene(-R, R * Math.Sin(97 * Math.PI / 180), R, -R * Math.Sin(97 * Math.PI / 180));
    }

    // Six congruent rectangles stacked 22 high: an L of two at the top, two tilted ones from its
    // bottom corners meeting corner to corner, and an L of two at the bottom. The total area?
    // Nothing to drag: the tilted ones meeting makes the rectangles 4 times as long as wide, tilted
    // 3-4-5, and the stack 8.8 widths high.
    [Puzzle("stack-of-congruent-rectangles")]
    public static void StackOfCongruentRectangles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double w = 2.5, L = 4 * w;
        // along the left tilted one's long side, and across it
        double ux = 0.6, uy = -0.8, nx = -0.8, ny = -0.6;
        V P(double x, double y) => d.Pt(x, y);
        // side by side they stay bright in the dark theme: a deep orange between them there
        string yellow = d.Fill("Yellow", "#FFFFFE05", "#FFE47D00", 3, darkOutline: "#FF8A3A00");
        d.Polygon(yellow, P(0, 0), P(w, 0), P(w, L), P(0, L));
        d.Polygon(yellow, P(w, 0), P(w + L, 0), P(w + L, w), P(w, w));
        d.Polygon(yellow, P(0, 0), P(w * nx, w * ny), P(w * nx + L * ux, w * ny + L * uy), P(L * ux, L * uy));
        double bx = w + L;
        d.Polygon(yellow, P(bx, 0), P(bx + w * ux, w * uy), P(bx + w * ux + L * nx, w * uy + L * ny), P(bx + L * nx, L * ny));
        double cx = w * nx + L * ux, cy = w * ny + L * uy, ex = bx + w * ux, ey = w * uy;
        d.Polygon(yellow, P(cx, cy - w), P(ex, cy - w), P(ex, cy), P(cx, cy));
        d.Polygon(yellow, P(ex, ey - L), P(ex + w, ey - L), P(ex + w, ey), P(ex, ey));
        string red = "#FFC00505";
        d.Dimension(d.Line("Arrow", red, 3), d.Text("RedMark", red, 30), P(-4.2, L), P(-4.2, cy - w), 0, "22", 1.5);
        d.Scene(-4.8, L, ex + w, cy - w);
    }

    // Six points on a circle and lines between them: two lines from a point outside through two
    // neighboring points, one to each of the far points beyond, and chords from those to the rest.
    // The sum of the three marked angles? Dragging the top point or the lower left one along the
    // circle changes the angles: their sum stays 180°.
    [Puzzle("angles-in-a-circle")]
    public static void AnglesInACircle(Lgf d)
    {
        const double R = 4;
        var o = d.Pt(0, 0, "O");
        V On(double degrees, string name) => d.Pt(R * Math.Cos(degrees * Math.PI / 180), R * Math.Sin(degrees * Math.PI / 180), name);
        var a = On(140.7, "A");
        var c = On(39.8, "C");
        var e2 = On(6, "D");
        var f = On(-85, "F");
        var b = d.ArcSlider(o, R, 72, 126, 97, (0, 0), "B");
        var e = d.ArcSlider(o, R, 198, 242, 217.6, (0, 0), "E");
        var p = d.Crossing(e, a, c, b, "P");
        d.Circle(d.Line("Orange", "#FFE06A10", 3), o, d.Pt(R, 0));
        string wedge = d.Wedge("Red", "#FFE8124E");
        d.SectorMark(wedge, p, e, c, 0.6);
        d.SectorMark(wedge, e2, c, e, 0.6);
        d.SectorMark(wedge, f, b, a, 1.4);
        string ink = d.Line("Ink", Lgf.Ink, 4);
        d.Segment(ink, p, e);
        d.Segment(ink, p, c);
        d.Segment(ink, a, f);
        d.Segment(ink, b, f);
        d.Segment(ink, c, e2);
        d.Segment(ink, e2, e);
        foreach (var point in new[] { a, c, e2, f })
        {
            d.DotAt(point, size: 11);
        }

        d.Scene(-R, 4.9, R, -R);
    }

    // A regular dodecagon, a corner at the top, with a long diagonal of 12; two stripes between
    // chords across it, each from the second to the third corner down (or up) from the top (bottom).
    // The total shaded area? Nothing to drag: the diagonal fixes it.
    [Puzzle("striped-dodecagon")]
    public static void StripedDodecagon(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 6;
        var corners = Enumerable.Range(0, 12).Select(i => d.Pt(R * Math.Cos(i * Math.PI / 6), R * Math.Sin(i * Math.PI / 6))).ToArray();
        // corners[i] at 30i degrees: [0] the right one, [3] the top
        string magenta = d.Fill("Magenta", "#FFC8005F", null);
        d.Polygon(magenta, corners[1], corners[2], corners[4], corners[5]);
        d.Polygon(magenta, corners[7], corners[8], corners[10], corners[11]);
        d.Polygon(d.Outline("Frame", "#FF6040C4", 3.5), corners);
        string brown = d.Line("Brown", "#FF8A3A10", 3.5);
        d.Segment(brown, corners[2], corners[4]);
        d.Segment(brown, corners[1], corners[5]);
        d.Segment(brown, corners[7], corners[11]);
        d.Segment(brown, corners[8], corners[10]);
        string lilac = "#FFB060C8";
        d.Dimension(d.Line("Arrow", lilac, 3), d.Text("LilacMark", lilac, 32), corners[7], corners[1], 0, "12", 0.9);
        d.Scene(-R, R, R, -R);
    }
}
