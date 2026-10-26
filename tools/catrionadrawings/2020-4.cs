#nullable disable

static partial class Puzzles
{
    // Five circles in a square: two parallel lines, from the bottom left corner and to the top right
    // one, cut a band holding three equal orange circles in a row, the outer two against the sides;
    // the red circles are the two corner triangles' incircles. The red area is 24, the orange? Nothing
    // to drag: the three circles touching fix the slope of the band.
    [Puzzle("five-circles-in-a-square")]
    public static void FiveCirclesInASquare(Lgf d)
    {
        d.Texture = "FeltTip";
        const double S = 6;
        // slope m of the band (in a unit square): the outer circles' centers are 2r from the middle one along it,
        // (1/2 - r) k = 2r with k = sqrt(1 + m²) and r = (1 - m) / 2k, so (1 + m²)(1 - m)² = (m² + 2m - 1)², k > 0
        double low = 0.45, high = 0.7;
        for (int i = 0; i < 100; i++)
        {
            double mid = (low + high) / 2;
            double f = 1 + mid * mid - (1 - mid) * (2 + Math.Sqrt(1 + mid * mid));
            if (f < 0)
            {
                low = mid;
            }
            else
            {
                high = mid;
            }
        }

        double m = (low + high) / 2, k = Math.Sqrt(1 + m * m);
        double r = S * (1 - m) / (2 * k), rho = S * (m + 1 - k) / 2, h = S * (1 - m);
        var bl = d.Pt(0, 0);
        var br = d.Pt(S, 0);
        var tr = d.Pt(S, S);
        var tl = d.Pt(0, S);
        var leftEnd = d.Pt(0, h);
        var rightEnd = d.Pt(S, S - h);
        string red = d.Fill("Red", "#FFFC5960", "#FFD93445", 2);
        string orange = d.Fill("Orange", "#FFFD9A4A", "#FFE8782A", 2);
        foreach (var (x, y) in new[] { (rho, S - rho), (S - rho, rho) })
        {
            d.Circle(red, d.Pt(x, y), d.Pt(x + rho, y));
        }

        foreach (var x in new[] { r, S / 2, S - r })
        {
            double y = m * x + h / 2;
            d.Circle(orange, d.Pt(x, y), d.Pt(x + r, y));
        }

        string navy = d.Line("Navy", "#FF0A0A6E", 3);
        d.Segment(navy, bl, rightEnd);
        d.Segment(navy, leftEnd, tr);
        d.Polygon(d.Outline("Frame", "#FF0A0A6E", 3), bl, br, tr, tl);
        d.Scene(0, S, S, 0);
    }

    // A rectangle in a semicircle, its top corners on the arc; two quarter circles as high as it, one
    // from its top left corner, one from its bottom right, touching in the middle. What fraction of
    // the semicircle is shaded? Nothing to drag: the touching makes the rectangle √3 times as wide
    // as high.
    [Puzzle("two-quarter-circles-in-a-rectangle-in-a-semi-circle")]
    public static void TwoQuarterCirclesInARectangleInASemiCircle(Lgf d)
    {
        const double h = 4;
        double w = Math.Sqrt(3) * h / 2, R = Math.Sqrt(w * w + h * h);
        var o = d.Pt(0, 0);
        var east = d.Pt(R, 0);
        var west = d.Pt(-R, 0);
        var tl = d.Pt(-w, h);
        var tr = d.Pt(w, h);
        var br = d.Pt(w, 0);
        var leftEnd = d.Pt(-w + h, h);
        var leftFoot = d.Pt(-w, 0);
        var rightFoot = d.Pt(w - h, 0);
        string blue = "#FF0634F0";
        string fill = d.Fill("Blue", "#FF04E2FD", blue);
        d.Arc("CircleSector", fill, tl, leftFoot, leftEnd);
        d.Arc("CircleSector", fill, br, tr, rightFoot);
        string line = d.Line("BlueLine", blue);
        d.Arc("CircleArc", line, o, east, west);
        d.Segment(line, west, east);
        d.Dashed(line, leftEnd, tr, 8);
        d.Scene(-R, R, R, 0);
    }

    // Four squares: a big one, a yellow one on its left side as diagonal, a green one from its top
    // right corner with the far corner on the yellow one's side, and a blue one (area 2) whose
    // diagonal runs from the green one's corner to the big one's bottom right corner. The green
    // area? That corner is always halfway up: dragging it sideways tilts the green and the blue
    // squares, the green stays twice the blue.
    [Puzzle("four-squares-iii")]
    public static void FourSquaresIII(Lgf d)
    {
        const double L = 6;
        var bl = d.Pt(0, 0);
        var br = d.Pt(L, 0, "B");
        var tr = d.Pt(L, L, "T");
        var tl = d.Pt(0, L);
        var p = d.Slider(V.At(0.72 * L, L / 2), V.At(0.93 * L, L / 2), (0.84 - 0.72) / 0.21, (0.84 * L, L / 2), "P");
        string l = Lgf.N(L);
        var q = d.Pt(new V($"P.X + P.Y - {l}", $"P.Y - P.X + {l}"), "Q");
        var a = d.Pt(new V("P.Y", $"2 * {l} - P.X"), "A");
        var m = Mid(d, p, br);
        var half = new V($"({l} - P.X) / 2", "-P.Y / 2");
        var turned = new V(V.Neg(half.Y), half.X);
        var b1 = d.Pt(m + turned);
        var b2 = d.Pt(m - turned);
        d.Polygon(d.Fill("Yellow", "#FFFFFE87", "#FFF5D800"), bl, d.Pt(L / 2, L / 2), tl, d.Pt(-L / 2, L / 2));
        d.Polygon(d.Fill("Green", "#FFE2FD98", "#FF8CD611"), a, tr, p, q);
        d.Polygon(d.Fill("Blue", "#FF10D3F7", "#FF0BA6DC"), p, b1, br, b2);
        // inked after coloring: the big square's sides run over the yellow, the green and the blue
        // across the paper and the yellow, which stays bright in the dark theme: a blue between light and deep there
        d.Polygon(d.Outline("Frame", "#FF0A28C0", 3, dark: "#FF6B7BFF"), bl, br, tr, tl);
        d.Label("?", Mid(d, a, p), d.ShapeText("GreenMark", "#FF5BAE1A", 40), size: 40);
        d.Label("2", m, d.ShapeText("BlueMark", "#FF1A5FC8", 40), size: 40);
        d.Scene(-L / 2, 1.22 * L, 1.2 * L, 0);
    }

    // Four squares: a big one made of a small one (5) and a middle one (20) in its bottom corners, and a
    // tilted pink one from its top right corner, one side through the middle square's top left
    // corner and the next through the big square's top left corner. The pink area? Nothing to drag:
    // the given areas fix it.
    [Puzzle("four-squares-ii")]
    public static void FourSquaresII(Lgf d)
    {
        double a = Math.Sqrt(5), b = 2 * a, L = a + b;
        // from the top right corner toward the middle square's corner, and square to it toward the top left corner
        double ux = -b / Math.Sqrt(a * a + b * b), uy = -a / Math.Sqrt(a * a + b * b);
        double side = L * -ux;
        var t = d.Pt(L, L);
        var c = d.Pt(L + side * ux, L + side * uy);
        var e = d.Pt(L + side * ux + side * uy, L + side * uy - side * ux);
        var f = d.Pt(L + side * uy, L - side * ux);
        d.Polygon(d.Fill("Pink", "#FFFC88BA", "#FFC0284A"), t, c, e, f);
        // inked after coloring: the big square's top runs over the pink
        string frame = d.Outline("Frame");
        d.Box(frame, V.At(0, 0), V.At(L, L));
        d.Box(frame, V.At(0, 0), V.At(a, a));
        d.Box(frame, V.At(a, 0), V.At(L, b));
        string red = "#FFB72E36";
        d.Label("?", Mid(d, t, e), d.ShapeText("PinkMark", red, 44), size: 44);
        string mark = d.Text("RedMark", red, 36);
        d.Label("5", V.At(a / 2, a / 2), mark, 36);
        d.Label("20", V.At(a + b / 2, b / 2), mark, 36);
        d.Scene(L + side * (ux + uy), L - side * ux, L, 0);
    }

    // An equilateral triangle and two circles touching each other and its base, the green one its
    // left side, the blue one its right side. Lines from where they touch to where they touch the
    // sides; the angle between them? Dragging the green circle's center along the bisector of its
    // corner trades the circles' sizes: the angle stays (150°).
    [Puzzle("two-circles-inside-an-equilateral-triangle")]
    public static void TwoCirclesInsideAnEquilateralTriangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 10;
        const string sqrt3 = "1.7320508075688772";
        var bl = d.Pt(0, 0);
        var br = d.Pt(L, 0);
        var apex = d.Pt(L / 2, L * S3);
        // from where the blue circle is nearly the incircle to where the green one is bigger than it
        var g = d.Slider(V.At(0.115 * L * Math.Sqrt(3), 0.115 * L), V.At(0.23 * L * Math.Sqrt(3), 0.23 * L), (0.143 - 0.115) / 0.115, (0.143 * L * Math.Sqrt(3), 0.143 * L), "G");
        // the blue radius: (L - √3 r1 - √3 r2)² = 4 r1 r2, the smaller root
        string a = d.Num($"{Lgf.N(L)} - {sqrt3} * G.Y", "a");
        string b = d.Num($"2 * {sqrt3} * {a} + 4 * G.Y", "b");
        string r2 = d.Num($"({b} - sqrt({b} ^ 2 - 12 * {a} ^ 2)) / 6", "r2");
        var o2 = d.Pt(new V($"{Lgf.N(L)} - {sqrt3} * {r2}", r2), "O2");
        var t1 = d.Pt(new V($"G.X - {s3} * G.Y", "1.5 * G.Y"), "T1");
        var t2 = d.Pt(new V($"O2.X + {s3} * {r2}", $"1.5 * {r2}"), "T2");
        var x = d.Pt(g + $"G.Y / (G.Y + {r2})" * (o2 - g), "X");
        d.Circle(d.Line("Green", "#FF7ADB14", 2.5), g, d.Pt(new V("G.X", "0")));
        d.Circle(d.Line("Blue", "#FF2FA0DC", 2.5), o2, d.Pt(new V("O2.X", "0")));
        d.Polygon(d.Outline("Frame", Lgf.Ink, 3), bl, br, apex);
        string navy = "#FF000095";
        string line = d.Line("Navy", navy, 3);
        d.Segment(line, x, t1);
        d.Segment(line, x, t2);
        d.Angle(line, x, t2, t1, 18);
        d.Label("?", x + V.At(-0.15, 0.95), d.Text("NavyMark", navy, 30), 30);
        d.Scene(0, L * S3, L, 0);
    }

    // Three congruent rectangles (area 9) in a staircase: one lying, one standing at its right end,
    // one lying on top of that. The triangle from the bottom left corner to the top one's top right
    // corner and its corner over the standing one, whose far side passes through the standing one's
    // top right corner. Its area? Nothing to drag: that corner makes the rectangles √2 + 1 times as
    // long as wide.
    [Puzzle("three-congruent-rectangles-ii")]
    public static void ThreeCongruentRectanglesII(Lgf d)
    {
        d.Texture = "FeltTip";
        double a = Math.Sqrt(9 * (Math.Sqrt(2) + 1)), b = 9 / a;
        d.Polygon(d.Fill("Teal", "#FF4CCBB8", null), d.Pt(0, 0), d.Pt(2 * a, a + b), d.Pt(a + b, a + b));
        // inked after coloring: the rectangles' sides run over the triangle
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        d.Box(frame, V.At(0, 0), V.At(a, b));
        d.Box(frame, V.At(a, 0), V.At(a + b, a));
        d.Box(frame, V.At(a, a), V.At(2 * a, a + b));
        d.Scene(0, a + b, 2 * a, 0);
    }

    // Five squares: two small white ones in opposite corners of the big one, two more like them side by
    // side, tilted in the middle, in a band between two parallel red lines from the big square's
    // other corners through the small squares' inner corners. The red is the band outside the tilted
    // squares, the purple the rest; their ratio? Nothing to drag: the band as wide as the small
    // squares fixes the figure.
    [Puzzle("four-congruent-squares-inside-a-larger-square")]
    public static void FourCongruentSquaresInsideALargerSquare(Lgf d)
    {
        const double L = 6;
        // the small side t (of a unit square): the band (slope t / (1 - t)) is t wide,
        // (1 - 2t)² = t² ((1 - t)² + t²)
        double low = 0.2, high = 0.45;
        for (int i = 0; i < 100; i++)
        {
            double mid = (low + high) / 2;
            double f = (1 - 2 * mid) * (1 - 2 * mid) - mid * mid * ((1 - mid) * (1 - mid) + mid * mid);
            if (f > 0)
            {
                low = mid;
            }
            else
            {
                high = mid;
            }
        }

        double t = (low + high) / 2 * L, m = t / (L - t), k = Math.Sqrt(1 + m * m);
        double ux = 1 / k, uy = m / k, nx = -uy, ny = ux;
        var bl = d.Pt(0, 0);
        var br = d.Pt(L, 0);
        var tr = d.Pt(L, L);
        var tl = d.Pt(0, L);
        var lowEnd = d.Pt(L, m * L);
        var highEnd = d.Pt(0, L - m * L);
        // the tilted squares: the band's middle, t along it each way, t / 2 across
        double cx = L / 2, cy = L / 2;
        V Corner(double along, double across) => d.Pt(cx + along * ux + across * nx, cy + along * uy + across * ny);
        V[] rectangle = [Corner(-t, -t / 2), Corner(t, -t / 2), Corner(t, t / 2), Corner(-t, t / 2)];
        string square = d.Polygon(null, bl, br, tr, tl);
        string band = d.Polygon(null, bl, lowEnd, tr, highEnd);
        string topLeft = d.Polygon(null, d.Pt(0, L - t), d.Pt(t, L - t), d.Pt(t, L), tl);
        string bottomRight = d.Polygon(null, d.Pt(L - t, 0), br, d.Pt(L, t), d.Pt(L - t, t));
        d.Region(d.Fill("Purple", "#FFEAAEF8", null), "Difference", square, band, topLeft, bottomRight);
        d.Region(d.Fill("Red", "#FFF487B2", null), "Difference", band, d.Polygon(null, rectangle));
        string red = d.Line("RedLine", "#FFD81E50", 3);
        d.Segment(red, bl, lowEnd);
        d.Segment(red, highEnd, tr);
        string frame = d.Outline("Frame", "#FF1A1A50", 3.5);
        d.Inked(frame, rectangle);
        d.Segment(d.Line("FrameLine", "#FF1A1A50", 3.5), Corner(0, -t / 2), Corner(0, t / 2));
        d.Inked(frame, d.Pt(0, L - t), d.Pt(t, L - t), d.Pt(t, L), tl);
        d.Inked(frame, d.Pt(L - t, 0), br, d.Pt(L, t), d.Pt(L - t, t));
        d.Inked(frame, bl, br, tr, tl);
        d.Scene(0, L, L, 0);
    }

    // An equilateral triangle of side 4 and two semicircles on the line through the middles of its
    // sides, both centered under its top: the orange one up through the top, the yellow one down
    // through the bottom corners. Their total area? Nothing to drag: the side fixes it.
    [Puzzle("two-semi-circles-and-an-equilateral-triangle")]
    public static void TwoSemiCirclesAndAnEquilateralTriangle(Lgf d)
    {
        d.Texture = "FeltTip";
        double r3 = Math.Sqrt(3), r7 = Math.Sqrt(7);
        var a = d.Pt(-2, 0);
        var b = d.Pt(2, 0);
        var top = d.Pt(0, 2 * r3);
        var o = d.Pt(0, r3);
        d.Arc("CircleSegment", d.Fill("Orange", "#FFFDD802", "#FFF2A200", 2), o, d.Pt(r3, r3), d.Pt(-r3, r3));
        d.Arc("CircleSegment", d.Fill("Yellow", "#FFFCFF00", null), o, d.Pt(-r7, r3), d.Pt(r7, r3));
        // all on the orange and the yellow, which stay bright in the dark theme: deep there too
        string purple = "#FFB04F92", deep = "#FF6A1A80";
        string line = d.Line("Purple", purple, 2.5, dark: deep);
        var left = d.DotAt(Mid(d, a, top));
        var right = d.DotAt(Mid(d, b, top));
        d.Segment(line, a, left, "TwoTicks");
        d.Segment(line, left, top, "TwoTicks");
        d.Segment(line, top, right, "TwoTicks");
        d.Segment(line, right, b, "TwoTicks");
        d.Segment(line, a, b);
        d.Dimension(line, d.Text("PurpleMark", purple, 30, dark: deep), d.Pt(-1.85, 0), d.Pt(1.85, 0), 0.3, "4", 0.3);
        d.Scene(-r7, 2 * r3, r7, r3 - r7);
    }

    // Two quarter circles with the same center on a line 12 long, back to back, and a semicircle cut
    // out of the big one on the rest of its vertical radius. The shaded area? Dragging the shared
    // center along the line trades the quarter circles' sizes: the shaded area stays.
    [Puzzle("two-quarter-circles-and-a-semi-circle")]
    public static void TwoQuarterCirclesAndASemiCircle(Lgf d)
    {
        const double W = 12;
        var left = d.Pt(0, 0);
        var right = d.Pt(W, 0);
        var c = d.Slider(V.At(2.6, 0), V.At(5.4, 0), (3.32 - 2.6) / 2.8, (3.32, 0), "C");
        string big = Lgf.N(W) + " - C.X";
        var top = d.Pt(new V("C.X", big), "Top");
        var knee = d.Pt(new V("C.X", "C.X"), "K");
        var middle = d.Pt(new V("C.X", Lgf.N(W / 2)), "M");
        string blue = "#FF1E4FC0";
        string bigQuarter = d.Arc("CircleSector", null, c, right, top, visible: false);
        string smallQuarter = d.Arc("CircleSector", null, c, knee, left, visible: false);
        string bite = d.Arc("CircleSegment", null, middle, knee, top, visible: false);
        d.Region(d.Fill("Blue", "#FF1BA6F1", blue), "Difference", d.Region(null, "Union", bigQuarter, smallQuarter), bite);
        string line = d.Line("BlueLine", blue, 3);
        d.Segment(line, c, top);
        d.Segment(line, left, right);
        d.Dimension(line, d.Text("BlueMark", blue, 32), right, left, 0.9, "12", 0.9);
        d.Scene(0, 9.2, W, -1.3);
    }

    // Three squares: a purple one and a blue one side by side, 5 wide together, and a pink one on the
    // purple one against the blue one, level with its top. The semicircle on the bottom passes
    // through the pink one's bottom left corner. The total area? Nothing to drag: that makes the
    // sides 2, 3 and 1.
    [Puzzle("three-stacked-squares")]
    public static void ThreeStackedSquares(Lgf d)
    {
        const double p = 2, b = 3, k = 1, W = p + b;
        d.Box(d.Fill("Purple", "#FFBD7AE3"), V.At(0, 0), V.At(p, p));
        d.Box(d.Fill("Pink", "#FFFEB7F9"), V.At(p - k, p), V.At(p, b));
        d.Box(d.Fill("Blue", "#FF10DCFB"), V.At(p, 0), V.At(W, b));
        string blue = "#FF1E2AC0";
        // across the purple and the blue: a blue between light and deep shows on both in the dark theme
        string line = d.Line("BlueLine", blue, 3, dark: "#FF6B7BFF");
        d.Arc("CircleArc", line, d.Pt(W / 2, 0), d.Pt(W, 0), d.Pt(0, 0));
        d.Dimension(d.Line("Arrow", blue, 3), d.Text("BlueMark", blue, 34), d.Pt(W, 0), d.Pt(0, 0), 0.38, "5", 0.3);
        d.Scene(0, b, W, -0.6);
    }

    // A square, a line from its top left corner to its right side, a circle of diameter 18 in the
    // bottom left corner under the line and one of diameter 4 in the top right corner over it, both
    // touching it. What fraction is shaded? Nothing to drag: the two circles fix the square and the
    // line (its slope is 2/9, so 8/9 is shaded). Her small circle is drawn far bigger than 4 would be;
    // here it is to scale.
    [Puzzle("two-circles-in-a-square")]
    public static void TwoCirclesInASquare(Lgf d)
    {
        const double R = 9, r = 2, m = 2.0 / 9;
        double s = R + r + R * Math.Sqrt(1 + m * m);
        var bl = d.Pt(0, 0);
        var br = d.Pt(s, 0);
        var tr = d.Pt(s, s);
        var tl = d.Pt(0, s);
        var end = d.Pt(s, s - m * s);
        d.Polygon(d.Fill("Green", "#FF9EE88E", null), bl, br, end, tl);
        string navy = "#FF05246F", teal = "#FF1A8FBF";
        var big = d.Pt(R, R);
        var small = d.Pt(s - r, s - r);
        // all on the green, which stays bright in the dark theme: deep there too
        d.Circle(d.Line("Navy", navy, 3, dark: "#FF14286E"), big, d.Pt(2 * R, R));
        d.Circle(d.Line("Teal", teal, 3), small, d.Pt(s, s - r));
        d.Segment(d.Line("GreenLine", "#FF2F8F45", 3), tl, end);
        d.Dimension(d.Line("NavyArrow", navy, 2.5, dark: "#FF14286E"),d.ShapeText("NavyMark", navy, 30), d.Pt(0, R), d.Pt(2 * R, R), 0, "18", 1.7);
        d.Dimension(d.Line("TealArrow", teal, 2.5), d.Text("TealMark", teal, 24), d.Pt(s - 2 * r, s - r), d.Pt(s, s - r), 0, "4", 0.75);
        d.Polygon(d.Outline("Frame", Lgf.Ink, 3.5), bl, br, tr, tl);
        d.Scene(0, s, s, 0);
    }

    // Two congruent rectangles of area 12, one standing, one lying at its foot, and a big one tilted 45°
    // on the standing one's top left corner, the lying one's top right corner and the bottom line.
    // The big one's area? Dragging the corner the small ones share along the bottom changes their
    // shape (the area stays 12): the big one stays twice as big.
    [Puzzle("three-overlapping-rectangles-ii")]
    public static void ThreeOverlappingRectanglesII(Lgf d)
    {
        d.Texture = "FeltTip";
        // up to where the white gap between them closes (b = √6)
        var foot = d.Slider(V.At(1.95, 0), V.At(2.42, 0), (2.325 - 1.95) / 0.47, (2.325, 0), "S");
        string b = "S.X";
        string a = d.Num("12 / S.X", "a");
        var o = d.Pt(0, 0);
        var standingTop = d.Pt(new V(b, a));
        var l = d.Pt(new V("0", a), "L");
        var bottom = d.Pt(new V(a, "0"), "B");
        var r = d.Pt(new V($"{a} + {b}", b), "R");
        var t = d.Pt(new V(b, $"{a} + {b}"), "T");
        var lyingRight = d.Pt(new V($"{a} + {b}", "0"));
        var lyingLeftTop = d.Pt(new V(b, b));
        V[] standing = [o, foot, standingTop, l];
        V[] lying = [foot, lyingRight, r, lyingLeftTop];
        V[] big = [l, bottom, r, t];
        string bigShape = d.Polygon(null, big);
        d.Region(d.Fill("Orange", "#FFF18310", null), "Difference", d.Region(null, "Union", d.Polygon(null, standing), d.Polygon(null, lying)), bigShape);
        d.Polygon(d.Fill("Magenta", "#FFDA0086", null), big);
        string frame = d.Outline("Frame", Lgf.Ink, 3.5);
        d.Inked(frame, standing);
        d.Inked(frame, lying);
        d.Inked(frame, big);
        d.Scene(0, 7.8, 7.8, 0);
    }
}
