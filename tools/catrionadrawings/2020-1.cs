#nullable disable

static partial class Puzzles
{
    // A square on its corner and four white equilateral triangles meeting at its center, each with its
    // base ends on two sides of the square; green corners beyond the bases, yellow triangles between the
    // white ones. The ratio of yellow to green? Nothing to drag: the triangles fix it.
    [Puzzle("four-triangles-in-a-square")]
    public static void FourTrianglesInASquare(Lgf d)
    {
        const double R = 4;
        double h = R / (1 + 1 / Math.Sqrt(3)), w = h / Math.Sqrt(3);
        var o = d.Pt(0, 0, "O");
        string green = d.Fill("Green", "#FFC4F5D6");
        string yellow = d.Fill("Yellow", "#FFFEF87C");
        string frame = d.Outline("Frame");
        var corners = new List<V>();
        var white = new List<(V, V)>();
        for (int k = 0; k < 4; k++)
        {
            double c = Math.Round(Math.Cos(k * Math.PI / 2)), s = Math.Round(Math.Sin(k * Math.PI / 2));
            V Turned(double x, double y) => d.Pt(c * x - s * y, s * x + c * y);
            corners.Add(Turned(0, R));
            white.Add((Turned(w, h), Turned(-w, h)));
        }

        for (int k = 0; k < 4; k++)
        {
            var (right, left) = white[k];
            d.Polygon(green, corners[k], left, right);
            d.Polygon(yellow, o, white[(k + 3) % 4].Item2, right);
        }

        foreach (var (right, left) in white)
        {
            d.Polygon(frame, o, right, left);
        }

        d.Polygon(frame, corners.ToArray());
        d.Scene(-R, R, R, -R);
    }

    // A rectangle cut across into a blue band over a purple one, each band cut into two rectangles that
    // are similar (one turned), and a zig-zag from the middle line's left end down to the foot of the
    // purple cut, up to the top of the blue cut and down to the middle line's right end. The sum of the
    // three marked angles? A cut at x across a band of width W is similar when the band's height h has
    // h^2 = x (W - x): the zig-zag's bottom and top corners are on the circle on the middle line. Both
    // are dragged along it: the bands change, the angles still add up the same.
    [Puzzle("zig-zag-inside-a-rectangle")]
    public static void ZigZagInsideARectangle(Lgf d)
    {
        const double W = 6;
        var c = d.Pt(W / 2, 0, "C");
        var o = d.Pt(0, 0, "O");
        var e = d.Pt(W, 0, "E");
        var q = d.ArcSlider(c, W / 2, 240, 266, 255.7, (W / 2, 0), "Q");
        var t = d.ArcSlider(c, W / 2, 30, 52, 42.6, (W / 2, 0), "T");
        string blueLine = "#FF1428B8", purpleLine = "#FF4A08A4";
        string blue = d.Fill("Blue", "#FFB4D5F6", blueLine);
        string purple = d.Fill("Purple", "#FFD6C4EE", purpleLine);
        d.Box(blue, V.At(0, 0), t);
        d.Box(blue, new V("T.X", "0"), new V(Lgf.N(W), "T.Y"));
        d.Box(purple, new V("0", "Q.Y"), new V("Q.X", "0"));
        d.Box(purple, q, V.At(W, 0));
        string ink = d.Line("Ink");
        d.Segment(ink, o, q);
        d.Segment(ink, q, t);
        d.Segment(ink, t, e);
        string mark = d.Fill("AngleMark", "#FFB4ACB8");
        d.Angle(mark, o, d.Pt(new V("0", "Q.Y")), q, 34);
        d.Angle(mark, q, d.Pt(new V("Q.X", "0")), t, 34);
        d.Angle(mark, e, d.Pt(new V(Lgf.N(W), "T.Y")), t, 34);
        d.Scene(0, 2.35, W, -3);
    }

    // A square of side 13 and a row of five squares across it: the row's lower long side runs from the
    // big square's bottom side to its top right corner, the row's first short side ends on its left
    // side. The total shaded area? Nothing to drag: that fixes the row (its foot is a fifth of the way
    // along the bottom).
    [Puzzle("five-squares-across-a-square")]
    public static void FiveSquaresAcrossASquare(Lgf d)
    {
        const double L = 13, b = L / 5;
        // along the row and across it, each a square's side
        double ux = (L - b) / 5, uy = L / 5, nx = -uy, ny = ux;
        string red = d.Fill("Red", "#FFFDC2B2", "#FFE00818");
        for (int k = 0; k < 5; k++)
        {
            double x = b + k * ux, y = k * uy;
            d.Polygon(red, d.Pt(x, y), d.Pt(x + ux, y + uy), d.Pt(x + ux + nx, y + uy + ny), d.Pt(x + nx, y + ny));
        }

        // inked after coloring: the big square's top runs over the row
        d.Box(d.Outline("Frame"), V.At(0, 0), V.At(L, L));
        d.Dimension(d.Line("Arrow"), d.Text(), d.Pt(0, 0), d.Pt(0, L), 0.9, "13", 1.1);
        d.Scene(-1.6, L + ny, L, 0);
    }

    // Two squares side by side in a semicircle of diameter 10, tilted at 45 degrees: the bottom corner on
    // the diameter, the far corners of the pair and the corner they share on the arc. Their total
    // area? Nothing to drag: the arc fixes them (each has area 10).
    [Puzzle("two-squares-and-a-semi-circle")]
    public static void TwoSquaresAndASemiCircle(Lgf d)
    {
        const double R = 5;
        double a = Math.Sqrt(10), k = a / Math.Sqrt(2);
        string blue = "#FF4A8EEA";
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0, "E");
        var west = d.Pt(-R, 0, "W");
        var p0 = d.Pt(-k, 0, "P0");
        var left = d.Pt(-2 * k, k, "A");
        var shared = d.Pt(-k, 2 * k, "B");
        var middle = d.Pt(0, k, "M");
        var right = d.Pt(k, 2 * k, "C");
        var top = d.Pt(0, 3 * k, "T");
        string fill = d.Fill("Blue", "#FFD8F3FE", null);
        d.Polygon(fill, p0, middle, shared, left);
        d.Polygon(fill, middle, right, top, shared);
        d.Arc("CircleSegment", d.Outline("Semicircle", blue), o, east, west);
        string frame = d.Outline("Frame");
        d.Inked(frame, p0, middle, shared, left);
        d.Inked(frame, middle, right, top, shared);
        d.Dimension(d.Line("Arrow", blue), d.Text("BlueMark", blue), west, east, -0.55, "10", 0.5);
        d.Scene(-R, 3 * k, R, -0.8);
    }

    // A square and its diagonal, two touching circles above it: the small one on the top side and the
    // diagonal, the big one on the right side and the diagonal. Lines from the small one's point on the
    // top to where the circles touch and on to the big one's point on the right side; the angle between
    // them? The small circle's point on the top is dragged: the circles change size, the angle stays.
    [Puzzle("two-circles-in-a-square-ii")]
    public static void TwoCirclesInASquareII(Lgf d)
    {
        const double L = 6;
        // a circle of radius r in a 45 degree corner touches the sides k r from the corner
        double k = 1 + Math.Sqrt(2);
        string kk = Lgf.N(k);
        var t1 = d.Slider(V.At(0.82 * k, L), V.At(1.3 * k, L), (0.93 - 0.82) / 0.48, (0.93 * k, L), "T1");
        string r1 = d.Num($"T1.X / {kk}", "r1");
        // along the diagonal the touching points are k r1 and k r2 from its ends and 2 sqrt(r1 r2) apart
        string u = d.Num($"(-sqrt({r1}) + sqrt({r1} - {kk} * ({kk} * {r1} - {Lgf.N(L * Math.Sqrt(2))}))) / {kk}", "u");
        string r2 = d.Num($"{u} ^ 2", "r2");
        var c1 = d.Pt(new V("T1.X", $"{Lgf.N(L)} - {r1}"), "C1");
        var c2 = d.Pt(new V($"{Lgf.N(L)} - {r2}", $"{kk} * {r2}"), "C2");
        var t2 = d.Pt(new V(Lgf.N(L), $"{kk} * {r2}"), "T2");
        var touch = d.Pt(c1 + $"{r1} / ({r1} + {r2})" * (c2 - c1), "K");
        d.Polygon(d.Fill("Green", "#FFA2DDA9", "#FF079565"), d.Pt(0, L),d.Pt(0, 0), d.Pt(L, 0));
        string light = d.Fill("Light", "#FFEAFCA8", "#FF72BC10");
        d.Circle(light, c1, t1);
        d.Circle(light, c2, t2);
        d.Box(d.Outline("Frame", "#FF079565"), V.At(0, 0), V.At(L, L));
        // inside the circles, which stay bright in the dark theme: ink there too; the mark is on the paper
        string ink = d.Line("Ink", Lgf.Ink, 2.5, dark: Lgf.Ink);
        d.Segment(ink, t1, touch);
        d.Segment(ink, touch, t2);
        d.Angle(d.Line("AngleInk"), touch, t1, t2, 30);
        d.Label("?", touch + V.At(0.3, 0.8));
        d.Scene(0, L, L, 0);
    }

    // Two rectangles side by side on a line, a line from a point on it through the small one's top left
    // corner to the big one's, and lines from that point to the big one's top right corner, from the
    // small one's top left corner and from the big one's top left corner to its bottom right corner. The
    // yellow triangle is 15, the orange total? Orange equals yellow when the rectangles are similar (as
    // she drew them, though the question doesn't say it): the orange total less the yellow is half the
    // big one's width times the small one's height less the other way round. Dragging the big one's top
    // right corner changes the shape of both, the small one's top left corner its size: orange stays
    // yellow's.
    [Puzzle("triangles-formed-from-rectangles")]
    public static void TrianglesFormedFromRectangles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double H = 7;
        var o = d.Pt(0, 0, "O");
        var r = d.Slider(V.At(3.5, H), V.At(4.5, H), 0.5, (4, H), "R");
        // the small rectangle is the big one scaled down from its bottom right corner's place
        var s = d.Slider(new V("-0.45 * R.X", Lgf.N(0.45 * H)), new V("-0.59 * R.X", Lgf.N(0.59 * H)), (0.565 - 0.45) / 0.14, (-0.565 * 4, 0.565 * H), "S");
        var t = d.Pt(new V("0", Lgf.N(H)), "T");
        var b = d.Pt(new V("R.X", "0"), "B");
        var a = d.Crossing(t, s, o, b, "A");
        var x1 = d.Crossing(a, r, s, b, "X1");
        var x2 = d.Crossing(a, r, t, b, "X2");
        d.Polygon(d.Fill("Orange", "#FFFBA306", null), a, s, x1);
        d.Polygon(d.Fill("Gold", "#FFFCC924", null), t, r, x2);
        d.Polygon(d.Fill("Yellow", "#FFFFFE58", null), x1, x2, b);
        // along and across the bright fills and the paper alike: a vivid orange in the dark theme
        string orange = d.Line("OrangeLine", "#FFF06A04", 2.5, dark: "#FFF07A1A");
        d.Segment(orange, a, t);
        d.Segment(orange, a, r);
        d.Segment(orange, s, b);
        d.Segment(orange, t, b);
        string frame = d.Outline("Frame", Lgf.Ink, 3.5);
        d.Inked(frame, d.Pt(new V("S.X", "0")), o, d.Pt(new V("0", "S.Y")), s);
        d.Inked(frame, o, b, r, t);
        d.Segment(d.Line("Ground", Lgf.Ink, 3.5), d.Pt(new V("A.X - 0.3", "0")), d.Pt(new V("R.X + 0.4", "0")));
        d.Scene(-6.1, H, 4.9, 0);
    }

    // An equilateral triangle, its incircle, and three equilateral triangles inside: a chord across the
    // circle; on it the blue one pointing down from the chord's left end to a corner on the circle, the
    // orange one pointing down from the blue one's right corner to the chord's right end, and the yellow
    // one pointing up from that corner, its left side along the blue one's and its right corner on the
    // circle. What fraction of the big triangle is shaded? The chord's left end is dragged along the
    // circle: the three change, the fraction stays. (The shared corner is at 30 degrees from the center,
    // so the yellow side is sqrt(r^2 - 4 h^2) for a chord at height h.)
    [Puzzle("three-triangles-in-a-circle-in-a-triangle")]
    public static void ThreeTrianglesInACircleInATriangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4;
        var o = d.Pt(0, 0, "O");
        var rim = d.Pt(0, -R);
        var l = d.ArcSlider(o, R, 157, 177, 169.3, (0, 0), "L");
        var p = d.Pt(new V($"2 * {s3} * L.Y", "L.Y"), "P");
        var r = d.Pt(new V("-L.X", "L.Y"), "Rt");
        string b = d.Num("P.X - L.X", "b");
        string y = d.Num($"sqrt({Lgf.N(R * R)} - 4 * L.Y ^ 2)", "y");
        string k = d.Num("Rt.X - P.X", "k");
        V Down(V from, string side) => d.Pt(from + new V($"0.5 * {side}", $"-{s3} * {side}"));
        var blueBottom = Down(l, b);
        var orangeBottom = Down(p, k);
        var yellowRight = Down(p, y);
        var yellowLeft = d.Pt(p + new V($"-0.5 * {y}", $"-{s3} * {y}"));
        d.Polygon(d.Outline("Frame", Lgf.Ink, 3.5), d.Pt(0, 2 * R), d.Pt(-2 * S3 * R, -R), d.Pt(2 * S3 * R, -R));
        d.Circle(d.Outline("Circle", "#FFF84040"), o, rim);
        // inked over the circle, as she drew them
        d.Inked(d.Fill("Blue", "#FF000094", Lgf.Ink, 4), l, p, blueBottom);
        d.Inked(d.Fill("Yellow", "#FFFFFF01", Lgf.Ink, 4), p, yellowLeft, yellowRight);
        d.Inked(d.Fill("Orange", "#FFFE3301", Lgf.Ink, 4), p, r, orangeBottom);
        d.Scene(-2 * S3 * R, 2 * R, 2 * S3 * R, -R);
    }

    // Two quarter circles of the same radius in a semicircle: one standing on the diameter from its left
    // end, the other hanging from the first one's top with its corner there, reaching down to the first
    // one's corner, its far end on the arc. What fraction of the semicircle do they cover? Nothing to
    // drag: that fixes them (radius 0.8 of the semicircle's).
    [Puzzle("two-quarter-circles-in-a-semi-circle")]
    public static void TwoQuarterCirclesInASemiCircle(Lgf d)
    {
        const double R = 5, a = 0.8 * R, c = a - R;
        string pink = "#FFF04DBE";
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0, "E");
        var west = d.Pt(-R, 0, "W");
        var corner = d.Pt(c, 0, "C");
        var top = d.Pt(c, a, "T");
        var far = d.Pt(c + a, a, "F");
        string fill = d.Fill("Pink", "#FFFDD5FD", pink);
        d.Arc("CircleSector", fill, corner, top, west);
        d.Arc("CircleSector", fill, top, corner, far);
        d.Arc("CircleSegment", d.Outline("Semicircle", pink), o, east, west);
        d.Scene(-R, R, R, 0);
    }

    // A 13-14-15 triangle and its incircle. The circle's area? Nothing to drag: the sides fix it (the
    // incircle has radius 4, its center at (7, 4) over the side of 15).
    [Puzzle("circle-inside-a-triangle")]
    public static void CircleInsideATriangle(Lgf d)
    {
        d.Texture = "FeltTip";
        var b = d.Pt(0, 0, "B");
        var c = d.Pt(15, 0, "C");
        var a = d.Pt(6.6, 11.2, "A");
        var center = d.Pt(7, 4, "I");
        d.Circle(d.Fill("Teal", "#FF07F2C1", "#FF05A598"), center, d.Pt(7, 0));
        // inked after coloring: the sides run over the circle's rim where it touches them
        d.Inked(d.Outline("Frame"), b, c, a);
        string mark = d.Text("Mark", Lgf.Ink, 30);
        d.Label("13", V.At(1.9, 6.3), mark);
        d.Label("14", V.At(11.7, 6.6), mark);
        d.Label("15", V.At(7, -1.3), mark);
        d.Scene(0, 11.2, 15, -1.9);
    }

    // A rectangle cut into similar rectangles: a strip across the top of three red and two blue ones and a
    // yellow one, under it a big yellow one and a tall red one. What fraction is red? Nothing to drag:
    // similarity fixes it (the rectangles are sqrt 3 to 1; the strip is 1 high, the big yellow 3).
    [Puzzle("similar-rectangles")]
    public static void SimilarRectangles(Lgf d)
    {
        d.Texture = "FeltTip";
        double k = Math.Sqrt(3), red = 1 / k, blue = k, side = k, below = k * k, wide = k * k * k;
        string redFill = d.Fill("Red", "#FFD80002", Lgf.Ink, 5);
        string blueFill = d.Fill("Blue", "#FF0100B2", Lgf.Ink, 5);
        string yellowFill = d.Fill("Yellow", "#FFFFFF01", Lgf.Ink, 5);
        double x = 0;
        foreach (var (width, style) in new[] { (red, redFill), (blue, blueFill), (red, redFill), (blue, blueFill), (red, redFill), (side, yellowFill) })
        {
            d.Box(style, V.At(x, below), V.At(x + width, below + 1));
            x += width;
        }

        d.Box(yellowFill, V.At(0, 0), V.At(wide, below));
        d.Box(redFill, V.At(wide, 0), V.At(wide + side, below));
        d.Scene(0, below + 1, wide + side, 0);
    }

    // Two touching semicircles in a circle, their bases parallel chords of it, facing each other. What
    // fraction is shaded? The small one's top corner is dragged along the circle: touching, the big one's
    // top corner is the small one's turned a quarter turn around the center, so the radii's squares add
    // up to the circle's and the fraction stays a half.
    [Puzzle("two-semi-circles-in-a-circle")]
    public static void TwoSemiCirclesInACircle(Lgf d)
    {
        const double R = 5;
        string blue = "#FF2440E0";
        var o = d.Pt(0, 0, "O");
        var rim = d.Pt(R, 0);
        var l = d.ArcSlider(o, R, 125, 163, 149.6, (0, 0), "L");
        var smallBottom = d.Pt(new V("L.X", "-L.Y"));
        var smallCenter = d.Pt(new V("L.X", "0"));
        var bigTop = d.Pt(new V("L.Y", "-L.X"));
        var bigBottom = d.Pt(new V("L.Y", "L.X"));
        var bigCenter = d.Pt(new V("L.Y", "0"));
        string fill = d.Fill("Blue", "#FF52BCFB", blue);
        d.Arc("CircleSegment", fill, smallCenter, smallBottom, l);
        d.Arc("CircleSegment", fill, bigCenter, bigTop, bigBottom);
        d.Circle(d.Outline("Circle", blue), o, rim);
        d.Scene(-R, R, R, -R);
    }

    // A square, a semicircle of diameter 10 hanging from its top side at the left corner and a quarter
    // circle of radius 8 in its bottom right corner, touching; the black line is their tangent where they
    // touch, across the square. How long is it? Nothing to drag: the touching fixes the square (its side
    // s has (s - 5)^2 + s^2 = 13^2).
    [Puzzle("semi-circle-and-quarter-circle-in-a-square")]
    public static void SemiCircleAndQuarterCircleInASquare(Lgf d)
    {
        double s = (5 + Math.Sqrt(313)) / 2;
        string brown = "#FFC04E04";
        var a = d.Pt(5, s, "A");
        var b = d.Pt(s, 0, "B");
        // the touching point, and the tangent's direction (across the line of centers)
        double px = 5 + 5.0 / 13 * (s - 5), py = s - 5.0 / 13 * s, ux = s / 13, uy = (s - 5) / 13;
        var left = d.Pt(0, py - px / ux * uy, "P0");
        var right = d.Pt(s, py + (s - px) / ux * uy, "P1");
        d.Box(d.Outline("Frame", brown), V.At(0, 0), V.At(s, s));
        string fill = d.Fill("Orange", "#FFF8B612", brown);
        d.Arc("CircleSegment", fill, a, d.Pt(0, s), d.Pt(10, s));
        d.Arc("CircleSector", fill, b, d.Pt(s, 8), d.Pt(s - 8, 0));
        d.Segment(d.Line("Tangent", Lgf.Ink, 4), left, right);
        string arrow = d.Line("Arrow", brown);
        string mark = d.Text("BrownMark", brown, 30);
        d.Dimension(arrow, mark, d.Pt(0, s), d.Pt(10, s), 0.6, "10", 0.8);
        d.Dimension(arrow, mark, d.Pt(s, 0), d.Pt(s, 8), -1, "8", 0.7);
        d.Scene(0, s + 1, s + 1.4, 0);
    }
}
