#nullable disable

static partial class Puzzles
{
    // The sequel to circles-in-triangle-and-rhombus: two copies of a rhombus, the cyan one cut by its
    // horizontal diagonal into two triangles, each with its incircle (blue), the yellow one with its
    // own incircle (red). Can a rhombus be made where the two blue circles cover the same share as the
    // red one? Dragging the yellow rhombus's right corner changes the shape of both (the vertical
    // diagonal stays): the shares are equal when the horizontal diagonal is √((√2 - 1) / 2) ≈ 0.455 of
    // the vertical one; in her drawing (about 0.585) the red circle covers more.
    [Puzzle("circles-in-rhombii-sequal")]
    public static void CirclesInRhombiiSequal(Lgf d)
    {
        d.Texture = "FeltTip";
        const double Q = 2, gap = 0.375;
        var h = d.Slider(V.At(2 * 0.5 + gap, 0), V.At(2 * 1.6 + gap, 0), (2 * 1.17 - 2 * 0.5) / 2.2, (2 * 1.17 + gap, 0), "H");
        // the half diagonals: p across, Q up
        string p = d.Num($"(H.X - {gap}) / 2", "p");
        string side = d.Num($"sqrt({p} ^ 2 + {Q * Q})", "side");
        var l = d.Pt(new V($"-{gap} - {p}", "0"), "L");
        var r = d.Pt(new V($"{gap} + {p}", "0"), "R");
        var cyan = new[] { d.Pt(l + new V($"-{p}", "0")), d.Pt(l + V.At(0, Q)), d.Pt(l + new V(p, "0")), d.Pt(l + V.At(0, -Q)) };
        var yellow = new[] { d.Pt(r + new V($"-{p}", "0")), d.Pt(r + V.At(0, Q)), h, d.Pt(r + V.At(0, -Q)) };
        string cyanColor = "#FF14D8F4";
        d.Polygon(d.Outline("Cyan", cyanColor, 3.5), cyan);
        d.Segment(d.Line("Diagonal", cyanColor, 3.5), cyan[0], cyan[2]);
        d.Polygon(d.Outline("Yellow", "#FFFAEA20", 3.5), yellow);
        // each triangle's incircle touches the diagonal at its middle
        string inner = d.Num($"{p} * {Q} / ({p} + {side})", "rb");
        string blue = d.Line("Blue", "#FF0A30EC", 3.5);
        d.Circle(blue, d.Pt(l + new V("0", inner)), l);
        d.Circle(blue, d.Pt(l - new V("0", inner)), l);
        string rhombus = d.Num($"{p} * {Q} / {side}", "rr");
        d.Circle(d.Line("Red", "#FFEE1A10", 3.5), r, d.Pt(r + new V("0", rhombus)));
        d.Scene(-2 * 1.6 - gap, Q, 2 * 1.6 + gap, -Q);
    }

    // An equilateral triangle and a rhombus on its right side, made of two such triangles, each with
    // its incircle; the triangle's has area 16, the rhombus's? Nothing to drag: the triangle fixes it.
    [Puzzle("circles-in-triangle-and-rhombus")]
    public static void CirclesInTriangleAndRhombus(Lgf d)
    {
        d.Texture = "FeltTip";
        double h = 2 * S3;
        var a = d.Pt(0, 0);
        var b = d.Pt(2, 0);
        var t = d.Pt(1, h);
        d.Polygon(d.Outline("Yellow", "#FFFAEA20", 3.5), b, d.Pt(4, 0), d.Pt(3, h), t);
        d.Polygon(d.Outline("Cyan", "#FF22DCEA", 3.5), a, b, t);
        string blueColor = "#FF0820EC", redColor = "#FFE8160E";
        var c = d.Pt(1, h / 3);
        d.Circle(d.Line("Blue", blueColor, 3.5), c, d.Pt(1, 0));
        var m = d.Pt(2.5, h / 2);
        d.Circle(d.Line("Red", redColor, 3.5), m, d.Pt(2.5, 0));
        d.Label("16", c, d.Text("BlueMark", blueColor, 28), 28);
        d.Label("?", m, d.Text("RedMark", redColor, 30), 30);
        d.Scene(0, h, 4, 0);
    }

    // Three identical right triangles in a row: the first lying, the second its mirror image standing
    // on the same leg, the third standing on its short leg. Semicircles on the first one's hypotenuse
    // (area 14) and from its top corner to the third one's; the missing area? The second diameter is
    // √2 times the hypotenuse, so the area is 28. Dragging the first triangle's top corner changes the
    // triangles' shape (the hypotenuse stays): the big semicircle stays twice the small one.
    [Puzzle("circles-resting-on-triangles")]
    public static void CirclesRestingOnTriangles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double C = 4;
        var o = d.Pt(0, 0, "O");
        var h = d.ArcSlider(o, C, 18, 44, 33.7, (0, 0), "H");
        var foot = d.Pt(new V("H.X", "0"));
        var e = d.Pt(new V("2 * H.X", "0"));
        var f = d.Pt(new V("2 * H.X + H.Y", "0"));
        var q = d.Pt(new V("2 * H.X + H.Y", "H.X"), "Q");
        string pink = d.Outline("Pink", "#FFEE1288", 3.5);
        d.Polygon(pink, o, foot, h);
        d.Polygon(pink, h, foot, e);
        d.Polygon(pink, e, f, q);
        var m1 = Mid(d, o, h);
        var m2 = Mid(d, h, q);
        string blueColor = "#FF1462E4", cyanColor = "#FF14CCEC";
        d.Arc("CircleSegment", d.Outline("Blue", blueColor, 3.5), m1, h, o);
        d.Arc("CircleSegment", d.Outline("Cyan", cyanColor, 3.5), m2, q, h);
        d.Label("14", m1 + new V("-0.2 * H.Y", "0.2 * H.X"), d.Text("BlueMark", "#FF1446C8", 30), 30);
        d.Label("?", m2 + new V("-0.23 * (H.X - H.Y)", "0.23 * (H.X + H.Y)"), d.Text("CyanMark", cyanColor, 34), 34);
        d.Scene(-0.6, 5.7, 8.9, 0);
    }

    // An arbelos: a semicircle and two inside it on parts of its diameter; the red line from where they
    // meet up to the big arc is 2. The shaded area? Dragging the meeting point along the diameter
    // changes the split and the big semicircle with it (the red line stays 2, its top on the arc):
    // the shaded area stays the circle on the red line, π.
    [Puzzle("arbelos")]
    public static void Arbelos(Lgf d)
    {
        var o = d.Pt(0, 0, "O");
        var c = d.Slider(V.At(-1.5, 0), V.At(0.9, 0), (1.5 - 0.85) / 2.4, (-0.85, 0), "C");
        string radius = d.Num("sqrt(C.X ^ 2 + 4)", "R");
        var a = d.Pt(new V($"-{radius}", "0"), "A");
        var b = d.Pt(new V(radius, "0"), "B");
        var top = d.Pt(new V("C.X", "2"), "D");
        var m1 = Mid(d, a, c);
        var m2 = Mid(d, c, b);
        string big = d.Arc("CircleSegment", null, o, b, a, visible: false);
        string left = d.Arc("CircleSegment", null, m1, c, a, visible: false);
        string right = d.Arc("CircleSegment", null, m2, b, c, visible: false);
        d.Region(d.Fill("Yellow", "#FFFFF808", null), "Difference", big, left, right);
        string ink = d.Line("Ink", Lgf.Ink, 4.5);
        d.Arc("CircleSegment", d.Outline("Frame", Lgf.Ink, 4.5), o, b, a);
        d.Arc("CircleArc", ink, m1, c, a);
        d.Arc("CircleArc", ink, m2, b, c);
        // across the yellow, which stays bright in the dark theme: deep there too
        string red = "#FFE0060C";
        d.Segment(d.Line("Red", red, 3.5, dark: "#FFB0101A"), c, top);
        d.Label("2", c + V.At(-0.3, 1.1), d.ShapeText("RedMark", red, 36), 36);
        d.Scene(-2.5, 2.5, 2.5, 0);
    }

    // An equilateral triangle in its circumcircle and four identical circles in it: one at the center,
    // three in the corners touching two sides and the middle one. What fraction of the big circle is
    // shaded? Nothing to drag: the circles fix the triangle (their radius is a quarter of the big one's).
    [Puzzle("four-circles-in-a-triangle-in-a-circle")]
    public static void FourCirclesInATriangleInACircle(Lgf d)
    {
        var o = d.Pt(0, 0);
        V Toward(double radius, int k) => d.Pt(radius * Math.Cos((90 + 120 * k) * Math.PI / 180), radius * Math.Sin((90 + 120 * k) * Math.PI / 180));
        var corners = Enumerable.Range(0, 3).Select(k => Toward(4, k)).ToArray();
        string yellow = d.Fill("Yellow", "#FFFEEC06", null);
        d.Circle(yellow, o, d.Pt(0, 1));
        for (int k = 0; k < 3; k++)
        {
            var center = Toward(2, k);
            d.Circle(yellow, center, d.Pt(new V(center.X, $"{center.Y} + 1")));
        }

        d.Circle(d.Line("Magenta", "#FFFF14C9", 3.5), o, corners[0]);
        d.Inked(d.Outline("Red", "#FFE41A1A", 3.5), corners);
        d.Scene(-4, 4, 4, -4);
    }

    // A triangle and a parallelogram with a corner at its top and two sides along its sides; the
    // parallelogram's other sides cut off two small triangles (areas 5 and 20) and make one below the
    // base (45). The parallelogram's area? The three are similar to the big one, with sides 1/6, 2/6 and
    // 3/6 of it. Dragging the top along a line parallel to the base changes the triangle's shape (its
    // area stays): every area stays.
    [Puzzle("triangles-with-a-parallelogram")]
    public static void TrianglesWithAParallelogram(Lgf d)
    {
        const double B = 6, H = 4.26;
        var l = d.Pt(0, 0, "L");
        var r = d.Pt(B, 0, "R");
        var t = d.Slider(V.At(0.8, H), V.At(5.6, H), (3.71 - 0.8) / 4.8, (3.71, H), "T");
        var p = d.Pt(l + (1.0 / 6) * (t - l), "P");
        var s = d.Pt(r + (1.0 / 3) * (t - r), "S");
        var q = d.Pt(p + s - t, "Q");
        var x = d.Pt(1, 0, "X");
        var y = d.Pt(4, 0, "Y");
        d.Polygon(d.Outline("Pencil", "#FF7E8696", 2.5), l, r, t);
        string navy = d.Line("Navy", "#FF16167E", 3);
        d.Segment(navy, p, q);
        d.Segment(navy, q, s);
        string ink = d.Line("Ink", Lgf.Ink, 3);
        d.Segment(ink, t, p);
        d.Segment(ink, t, s);
        string mark = d.Text("Mark", Lgf.Ink, 26);
        d.Label("5", Mid(d, l, p, x), mark, 26);
        d.Label("20", Mid(d, y, s, r), mark, 26);
        d.Label("45", Mid(d, x, q, y), d.Text("BlueMark", "#FF1E6CC0", 30), 30);
        d.Scene(0, H, B, -H / 2);
    }

    // A square with three equal squares in three corners; the lines from the fourth corner to the top
    // right square's far corners run through the other two squares' inner corners. The shaded area
    // (the squares have area 5)? Nothing to drag: the lines fix the squares at (3 - √5) / 2 of the side.
    [Puzzle("three-squares-in-a-square")]
    public static void ThreeSquaresInASquare(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 6;
        double a = (3 - Math.Sqrt(5)) / 2 * L;
        var o = d.Pt(0, 0);
        d.Polygon(d.Fill("Lime", "#FFC8FC28", null), o, d.Pt(a, L - a), d.Pt(L - a, L - a), d.Pt(L - a, a));
        string greenColor = "#FF06A008";
        string green = d.Outline("Green", greenColor, 3.5);
        d.Box(green, V.At(0, L - a), V.At(a, L));
        d.Box(green, V.At(L - a, L - a), V.At(L, L));
        d.Box(green, V.At(L - a, 0), V.At(L, a));
        string ink = d.Line("Ink", Lgf.Ink, 4);
        d.Segment(ink, o, d.Pt(L - a, L));
        d.Segment(ink, o, d.Pt(L, L - a));
        d.Inked(d.Outline("Frame", Lgf.Ink, 4.5), o, d.Pt(L, 0), d.Pt(L, L), d.Pt(0, L));
        string mark = d.Text("GreenMark", greenColor, 44);
        d.Label("5", V.At(a / 2, L - a / 2), mark, 44);
        d.Label("5", V.At(L - a / 2, L - a / 2), mark, 44);
        d.Label("5", V.At(L - a / 2, a / 2), mark, 44);
        // a little paper under it: the page's buttons sit at the bottom right corner
        d.Scene(0, L, L, -0.5);
    }

    // An equilateral triangle and two circles of area 3: one in the bottom left corner touching two
    // sides, the other touching it and the right side. The side? Read as the circles' centers on the
    // line from that corner, perpendicular to the right side (as she drew it; the largest triangle
    // in which the second circle can touch both): the side is 10/√3 radii. Nothing to drag. (Her
    // triangle is a little larger.)
    [Puzzle("two-circles-in-an-equilateral-triangle")]
    public static void TwoCirclesInAnEquilateralTriangle(Lgf d)
    {
        double r3 = Math.Sqrt(3), s = 10 / r3;
        var a = d.Pt(0, 0);
        var b = d.Pt(s, 0);
        var t = d.Pt(s / 2, s * S3);
        d.Polygon(d.Outline("Blue", "#FF0A5CB0", 3.5), a, b, t);
        string redColor = "#FFD40A0A";
        string red = d.Line("Red", redColor, 3.5);
        var c1 = d.Pt(r3, 1);
        var c2 = d.Pt(2 * r3, 2);
        d.Circle(red, c1, d.Pt(r3, 0));
        d.Circle(red, c2, d.Pt(2 * r3, 1));
        string mark = d.Text("RedMark", redColor, 36);
        d.Label("3", c1, mark, 36);
        d.Label("3", c2, mark, 36);
        string navyColor = "#FF0C10A8";
        string navy = d.Line("Navy", navyColor, 3);
        var middle = d.Pt(s / 2, -0.5);
        d.Vector(navy, d.Pt(s / 2 - 0.4, -0.5), d.Pt(0, -0.5));
        d.Vector(navy, d.Pt(s / 2 + 0.4, -0.5), d.Pt(s, -0.5));
        d.Label("?", middle, d.Text("NavyMark", navyColor, 36), 36);
        // a little paper under it: the page's buttons sit at the bottom right corner
        d.Scene(0, s * S3, s, -1.2);
    }

    // A square standing on a corner, its diagonals, and four semicircles of radius 5 on the half
    // diagonals, one bulging out of each quarter like a pinwheel. The shaded area (each quarter's part
    // outside its semicircle and the semicircle's part outside the square)? Nothing to drag: the radius
    // fixes it.
    [Puzzle("semi-circles-in-a-square")]
    public static void SemiCirclesInASquare(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 5;
        var o = d.Pt(0, 0);
        var corners = new[] { d.Pt(0, 2 * R), d.Pt(-2 * R, 0), d.Pt(0, -2 * R), d.Pt(2 * R, 0) };
        string square = d.Polygon(null, corners);
        string yellow = d.Fill("Yellow", "#FFF7E204", null);
        var centers = new V[4];
        for (int k = 0; k < 4; k++)
        {
            // the semicircle on the half diagonal to this corner bulges into the quarter toward the next corner
            var corner = corners[k];
            centers[k] = Mid(d, o, corner);
            string semicircle = d.Arc("CircleSegment", null, centers[k], corner, o, visible: false);
            string quarter = d.Polygon(null, o, corner, corners[(k + 1) % 4]);
            d.Region(yellow, "Difference", quarter, semicircle);
            d.Region(yellow, "Difference", semicircle, square);
        }

        d.Polygon(d.Outline("Pink", "#FFF04896", 3.5), corners);
        string orangeColor = "#FFF58A35";
        string orange = d.Line("Orange", orangeColor, 3.5);
        d.Segment(orange, corners[0], corners[2]);
        d.Segment(orange, corners[1], corners[3]);
        for (int k = 0; k < 4; k++)
        {
            d.Arc("CircleArc", orange, centers[k], corners[k], o);
        }

        d.Scene(-2 * R, 2 * R, 2 * R, -2 * R);
    }

    // Five identical rectangles: four in a pinwheel around the square's sides, the fifth across the
    // hole in the middle with a corner on each of its sides (so at 45°). What fraction of the square do
    // they cover? Nothing to drag: the fifth one fixes the rectangles' shape (3 + 2√2 to 1).
    [Puzzle("five-rectangles-in-a-square")]
    public static void FiveRectanglesInASquare(Lgf d)
    {
        const double W = 1;
        double l = 3 + 2 * Math.Sqrt(2), s = l + W, t = W / Math.Sqrt(2);
        string pink = d.Fill("Pink", "#FFE6C2D4", "#FFD47A9C", 2.5);
        d.Box(pink, V.At(0, s - W), V.At(l, s));
        d.Box(pink, V.At(l, W), V.At(s, s));
        d.Box(pink, V.At(W, 0), V.At(s, W));
        d.Box(pink, V.At(0, 0), V.At(W, l));
        d.Polygon(pink, d.Pt(W, l - t), d.Pt(W + t, l), d.Pt(l, W + t), d.Pt(l - t, W));
        d.Inked(d.Outline("Frame", "#FFAE8035", 3.5), d.Pt(0, 0), d.Pt(s, 0), d.Pt(s, s), d.Pt(0, s));
        // a little paper under it: the page's buttons sit at the bottom right corner
        d.Scene(0, s, s, -0.6);
    }

    // Three equal circles in a row, each through its neighbors' centers; the cyan circle (area 36) fills
    // the lens of the left two, the smallest touches the middle one inside and the outer ones outside.
    // Its area? Nothing to drag: the circles fix it (a quarter of the big radius, half the cyan one's).
    [Puzzle("five-overlapping-circles")]
    public static void FiveOverlappingCircles(Lgf d)
    {
        d.Texture = "FeltTip";
        string magenta = d.Line("Magenta", "#FFDC1A7E", 3.5);
        foreach (double x in new[] { -1.0, 0, 1 })
        {
            d.Circle(magenta, d.Pt(x, 0), d.Pt(x + 1, 0));
        }

        string cyanColor = "#FF18CCE0", navyColor = "#FF0806A0";
        var cyan = d.Pt(-0.5, 0);
        d.Circle(d.Line("Cyan", cyanColor, 3.5), cyan, d.Pt(0, 0));
        var small = d.Pt(0, -0.75);
        d.Circle(d.Line("Navy", navyColor, 3.5), small, d.Pt(0, -1));
        d.Label("36", cyan, d.Text("CyanMark", cyanColor, 30), 30);
        d.Label("?", small, d.Text("NavyMark", navyColor, 22), 22);
        d.Scene(-2, 1, 2, -1);
    }

    // Two regular pentagons sharing their top corner (the cyan one, area 12, turned a little) and two
    // equilateral triangles: the purple one (area 3) from the yellow one's bottom left corner to the cyan
    // one's, the red one from the bottom right to the bottom right. The red one's area? Turning about the
    // top corner takes one triangle's side to the other's, so it is 3 too. Dragging the cyan pentagon's
    // bottom left corner turns it; the yellow one grows or shrinks to keep the purple triangle's area,
    // and the red one stays the same.
    [Puzzle("four-regular-polygons")]
    public static void FourRegularPolygons(Lgf d)
    {
        d.Texture = "FeltTip";
        double side = Math.Sqrt(12 / (1.25 / Math.Tan(36 * Math.PI / 180)));
        double far = side * (1 + Math.Sqrt(5)) / 2;
        double triangle = Math.Sqrt(3 * 4 / Math.Sqrt(3));
        double ux = Math.Cos(252 * Math.PI / 180), uy = Math.Sin(252 * Math.PI / 180);
        double half = 2 * Math.Cos(18 * Math.PI / 180);
        var a = d.Pt(0, 0, "A");
        var h = d.ArcSlider(a, far, 228, 276, 262, (0, 0), "H");
        string along = d.Num($"H.X * {Lgf.N(ux)} + H.Y * {Lgf.N(uy)}", "along");
        string across = d.Num($"H.Y * {Lgf.N(ux)} - H.X * {Lgf.N(uy)}", "across");
        // the yellow pentagon's diagonal: the purple triangle's side stays its side
        string diagonal = d.Num($"{along} + sqrt({Lgf.N(triangle * triangle)} - {across} ^ 2)", "U");
        string radius = d.Num($"{diagonal} / {Lgf.N(half)}", "Ry");
        var yellow = Enumerable.Range(0, 5).Select(k => k == 0 ? a : d.Pt(new V(
            $"{radius} * {Lgf.N(Math.Cos((90 + 72 * k) * Math.PI / 180))}",
            $"{radius} * ({Lgf.N(Math.Sin((90 + 72 * k) * Math.PI / 180))} - 1)"))).ToArray();
        var bottomLeft = yellow[2];
        var bottomRight = yellow[3];
        var center = d.Pt((1 / half) * V.Of("H").Turn(18), "Cc");
        var cyan = Enumerable.Range(0, 5).Select(k => k == 0 ? a : d.Pt(center + (V.At(0, 0) - center).Turn(72 * k))).ToArray();
        var q = d.Pt(V.Of("H").Turn(36), "Q");
        var purpleCorner = d.Pt(bottomLeft + (h - bottomLeft).Turn(60));
        var redCorner = d.Pt(bottomRight + (q - bottomRight).Turn(-60));
        d.Polygon(d.Outline("Yellow", "#FFF1E028", 3.5), yellow);
        string cyanColor = "#FF28B4CC", purpleColor = "#FF6E64B0", redColor = "#FFD8581E";
        d.Polygon(d.Outline("Cyan", cyanColor, 3.5), cyan);
        d.Polygon(d.Outline("Purple", purpleColor, 3.5), bottomLeft, h, purpleCorner);
        d.Polygon(d.Outline("Red", redColor, 3.5), bottomRight, q, redCorner);
        d.Label("12", center, d.Text("CyanMark", cyanColor, 32), 32);
        d.Label("3", Mid(d, bottomLeft, h, purpleCorner), d.Text("PurpleMark", purpleColor, 32), 32);
        d.Label("?", Mid(d, bottomRight, q, redCorner), d.Text("RedMark", redColor, 32), 32);
        d.Scene(-4.6, 0, 6, -6.9);
    }
}
