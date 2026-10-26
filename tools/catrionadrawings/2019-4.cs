#nullable disable

static partial class Puzzles
{
    // Two squares side by side on the diameter of a semicircle (16), each with its outer top corner on
    // the arc. Their total area? Dragging the blue square's corner along the arc trades their sizes:
    // the yellow one's corner is the blue one's turned a quarter turn, and the total stays the radius
    // squared.
    [Puzzle("two-squares-in-a-semi-circle")]
    public static void TwoSquaresInASemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 8;
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0);
        var west = d.Pt(-R, 0);
        // the blue square's top left corner (-b, a); the yellow one's top right corner is (a, b)
        var p = d.ArcSlider(o, R, 118, 162, 142, (0, 0), "P");
        d.Box(d.Fill("Blue", "#FF01FFE8"), new V("P.X", "0"), new V("(P.X + P.Y)", "P.Y"));
        d.Box(d.Fill("Yellow", "#FFFEF702"), new V("(P.X + P.Y)", "0"), new V("P.Y", "-P.X"));
        string red = "#FFE3201E";
        d.Arc("CircleSegment", d.Outline("Frame", red, 3), o, east, west);
        d.Dimension(d.Line("Arrow", red, 2.5), d.Text("RedMark", red, 30), east, west, 0.9, "16", 0.8);
        d.Scene(-R, R, R, -1.5);
    }

    // A big square on the diameter of a semicircle with two small squares (area 6 each) on its top,
    // the pair's top corners and the big square's top right corner on the arc. The red line from the
    // pair's corner to the big square's? Nothing to drag: the corners on the arc fix the shape (the
    // pair is centered, and the big square is 2 + √5 times as wide as a small one).
    [Puzzle("three-squares-in-a-semi-circle")]
    public static void ThreeSquaresInASemiCircle(Lgf d)
    {
        const double s = 1;
        double S = s * (2 + Math.Sqrt(5)), R = Math.Sqrt(s * s + (S + s) * (S + s));
        var o = d.Pt(0, 0);
        var east = d.Pt(R, 0);
        var west = d.Pt(-R, 0);
        string teal = "#FF2BB592";
        d.Box(d.Fill("Big", "#FFD8FCDF", teal, 3), V.At(-s, 0), V.At(S - s, S));
        string small = d.Fill("Small", "#FFAFF1D2", teal, 3);
        d.Box(small, V.At(-s, S), V.At(0, S + s));
        d.Box(small, V.At(0, S), V.At(s, S + s));
        d.Arc("CircleSegment", d.Outline("Frame", "#FF0A0A55", 3), o, east, west);
        var a = d.Pt(s, S + s);
        var b = d.Pt(S - s, S);
        var m = Mid(d, a, b);
        string red = d.Line("Red", "#FFD02050", 3);
        d.Vector(red, m, a);
        d.Vector(red, m, b);
        d.Label("?", m + V.At(0.15, 0.7), d.Text("RedMark", "#FFD02050", 34), 34);
        d.Scene(-R, R, R, 0);
    }

    // Three congruent squares, together as big as the big square, in a chain from its top left corner
    // to its bottom right one, the middle one with a corner on its top right corner. The angle at its
    // bottom left corner up to the middle square's lowest corner? Nothing to drag: the areas fix the
    // chain (a small square's side is the big one's over √3).
    [Puzzle("three-congruent-squares-overlapping-a-square")]
    public static void ThreeCongruentSquaresOverlappingASquare(Lgf d)
    {
        const double L = 4;
        double h = L / (2 * Math.Sqrt(3));
        var bl = d.Pt(0, 0);
        var br = d.Pt(L, 0);
        var tr = d.Pt(L, L);
        var tl = d.Pt(0, L);
        var p1 = d.Pt(L / 2, L - h);
        var p2 = d.Pt(L / 2 + h, L / 2 - h);
        var q = d.Pt(L + h, L / 2);
        string cyan = d.Fill("Cyan", "#FF7FDBFB", "#FF1FBCEB", 3);
        d.Polygon(cyan, tl, p1, d.Pt(L / 2 + h, 1.5 * L - h), d.Pt(h, 1.5 * L));
        d.Polygon(cyan, p1, tr, q, p2);
        d.Polygon(cyan, q, br, d.Pt(1.5 * L, -h), d.Pt(1.5 * L + h, L / 2 - h));
        // inked after coloring: the big square's sides run over the cyan squares
        d.Polygon(d.Outline("Frame", "#FF2F4BDB", 3), bl, br, tr, tl);
        string purple = d.Line("Purple", "#FFA77BEB", 2.5);
        d.Segment(purple, bl, p2);
        d.ArcMark(purple, bl, br, p2, 0.4);
        d.Label("?", V.At(0.4, 0.7), d.Text("PurpleMark", "#FFA77BEB", 26));
        d.Scene(0, 1.5 * L, 1.5 * L + h, -h);
    }

    // A square split by a line from its left side (4 up) to its top right corner and a white triangle
    // on its bottom with the apex on that line: the orange, yellow and pink triangles have equal areas.
    // The red line? Nothing to drag: the equal areas fix the square (side 2 + 2√5, the apex 4 from the
    // left side).
    [Puzzle("three-triangles-in-a-square")]
    public static void ThreeTrianglesInASquare(Lgf d)
    {
        double L = 2 + 2 * Math.Sqrt(5);
        var bl = d.Pt(0, 0);
        var br = d.Pt(L, 0);
        var tr = d.Pt(L, L);
        var tl = d.Pt(0, L);
        var e = d.Pt(0, 4);
        var p = d.Pt(4, 4 + 4 * (L - 4) / L);
        d.Polygon(d.Fill("Orange", "#FFFDC048"), tl, e, tr);
        d.Polygon(d.Fill("Yellow", "#FFFEF52A"), e, bl, p);
        d.Polygon(d.Fill("Pink", "#FFFC8686"), p, br, tr);
        string ink = d.Line("Ink");
        d.Segment(ink, bl, p);
        d.Segment(ink, p, br);
        d.Polygon(d.Outline("Frame", Lgf.Ink, 3), bl, br, tr, tl);
        // on the orange, the yellow and the pink, which stay bright in the dark theme: deep there too
        string red = "#FFEC0A0A";
        var m = Mid(d, e, tr);
        string line = d.Line("Red", red, 3, dark: "#FFA01010");
        d.Vector(line, m, e);
        d.Vector(line, m, tr);
        d.Label("?", V.At(2.5, L - 0.78), d.ShapeText("RedMark", red, 34), 34);
        d.Dimension(d.Line("Arrow"), d.Text(), e, bl, -0.3, "4", 0.42);
        d.Scene(-0.8, L, L, 0);
    }

    // A regular hexagon split into thirds by two lines from its bottom left corner. The angle between
    // them? Nothing to drag: equal thirds put their ends at the middles of the top side and the upper
    // right side.
    [Puzzle("trisected-hexagon")]
    public static void TrisectedHexagon(Lgf d)
    {
        var h = HexagonUp(d, V.At(0, 0), "2");
        var p = Mid(d, h[3], h[4]);
        var q = Mid(d, h[2], h[3]);
        d.Polygon(d.Fill("Yellow", "#FFFFFE4A", null), h[0], p, h[4], h[5]);
        d.Polygon(d.Fill("Blue", "#FF66ECFC", null), h[0], q, h[3], p);
        d.Polygon(d.Fill("Pink", "#FFFC5EA7", null), h[0], h[1], h[2], q);
        string blue = "#FF0204B0";
        // only on the fills, which the dark theme keeps mid-tone: deep there too
        string line = d.Line("Navy", blue, 3, dark: "#FF10106E");
        d.Segment(line, h[0], p);
        d.Segment(line, h[0], q);
        d.Polygon(d.Outline("Frame", Lgf.Ink, 3), h);
        d.ArcMark(line, h[0], q, p, 0.55);
        d.Label("?", V.At(0.56, 1.02), d.ShapeText("BlueMark", blue, 36), 36);
        d.Scene(-1, 4 * S3, 3, 0);
    }

    // Two congruent quadrilaterals and a rectangle: the left one has two corners on the rectangle's top
    // and one on its bottom; the right one is it turned half a turn about a point of its long side, so
    // that its far corner lands on the rectangle's bottom too. The black line from the left one's
    // corner to the right one's (through that point) cuts off the yellow and the green; which is
    // larger? Dragging the left one's top corner (anywhere in a box) and its bottom corner (along the
    // rectangle's bottom) changes both: they stay equal.
    [Puzzle("two-congruent-quadrilaterals-and-a-rectangle")]
    public static void TwoCongruentQuadrilateralsAndARectangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double h = 1.6, c = 5.8;
        var a = d.Pt(0, 0, "A");
        var cc = d.Pt(c, 0, "C");
        var box = d.Polygon(null, d.Pt(-0.6, 2.6), d.Pt(1.8, 2.6), d.Pt(1.8, 4.3), d.Pt(-0.6, 4.3));
        var t = d.HandleIn(box, 0.64, 3.56, "T");
        var dd = d.Slider(V.At(2.6, -h), V.At(5.2, -h), (4.04 - 2.6) / 2.6, (4.04, -h), "D");
        // the center of the half turn: on the long side, halfway between the top corner's height and the bottom
        string k = d.Num($"(T.Y + {Lgf.N(h)}) / (2 * T.Y)", "k");
        var x = d.Pt(t + k * (cc - t), "X");
        var hh = d.Pt(2 * x - a, "H");
        var kk = d.Pt(2 * x - t, "K");
        var f = d.Pt(2 * x - cc, "F");
        var g = d.Pt(2 * x - dd, "G");
        d.Polygon(d.Fill("Yellow", "#FFFEFE2D", null), a, t, x);
        d.Polygon(d.Fill("Green", "#FFE6FE74", null), f, g, hh, x);
        d.Box(d.Outline("Rectangle", "#FFA8D82A", 3), V.At(-1, -h), new V("(H.X + 1)", "0"));
        string frame = d.Outline("Frame", "#FF1F8A45", 3.5);
        d.Polygon(frame, a, t, cc, dd);
        d.Polygon(frame, f, g, hh, kk);
        d.Segment(d.Line("Ink", Lgf.Ink, 3.5), a, hh);
        d.Scene(-1.1, 4.3, 9.6, -h);
    }

    // A semicircle in a square on a diameter from the left side to the bottom; the lines from a point of
    // its arc through the diameter's ends run across the square, the blue one (25) from top to bottom,
    // the green one from side to side. The green one? Dragging the diameter's ends along the sides moves
    // the arc and the lines; they cross at a right angle on the arc and stay the same length.
    [Puzzle("semi-circle-in-a-square-ii")]
    public static void SemiCircleInASquareII(Lgf d)
    {
        const double L = 10, tan = 0.2145;
        double c = 1 / Math.Sqrt(1 + tan * tan), s = tan * c;
        var p = d.Slider(V.At(0, 0.8), V.At(0, 2.6), (1.87 - 0.8) / 1.8, (0, 1.87), "P");
        var q = d.Slider(V.At(3.2, 0), V.At(6.2, 0), (4.9 - 3.2) / 3, (4.9, 0), "Q");
        // the green line from P along (c, s) and the blue one through Q square to it: they cross on the arc, at W
        string a = d.Num($"Q.X * {Lgf.N(c)} - P.Y * {Lgf.N(s)}", "a");
        var w = d.Pt(p + a * V.At(c, s), "W");
        var top = d.Pt(new V($"Q.X - {Lgf.N(L * tan)}", Lgf.N(L)), "Top");
        var right = d.Pt(new V(Lgf.N(L), $"P.Y + {Lgf.N(L * tan)}"), "Right");
        var m = Mid(d, p, q);
        d.Box(d.Outline("Frame", Lgf.Ink, 3), V.At(0, 0), V.At(L, L));
        d.Arc("CircleSegment", d.Outline("LightBlue", "#FF1E9BE6", 3), m, q, p);
        string blue = "#FF1A1AD8", green = "#FF7CD62E";
        // the numbers in gaps away from the crossing
        DimensionPartway(d, d.Line("Blue", blue, 3), d.Text("BlueMark", blue, 30), top, q, top + 0.62 * (w - top), "25", 0.75, 30);
        DimensionPartway(d, d.Line("Green", green, 3), d.Text("GreenMark", green, 30), p, right, w + 0.35 * (right - w), "?", 0.5, 30);
        d.Scene(0, L, L, 0);
    }

    // A regular pentagon standing on the diameter of a semicircle, its top corner on the arc, and the
    // circle in it. What fraction of the semicircle is shaded? Nothing to drag: it is all one shape.
    [Puzzle("circle-in-pentagon-in-semi-circle")]
    public static void CircleInPentagonInSemiCircle(Lgf d)
    {
        const double a = 2;
        double r = a / (2 * Math.Tan(Math.PI / 5)), big = a / (2 * Math.Sin(Math.PI / 5)), R = r + big;
        var o = d.Pt(0, 0);
        var east = d.Pt(R, 0);
        var west = d.Pt(-R, 0);
        var center = d.Pt(0, r);
        var corners = Enumerable.Range(0, 5).Select(i => d.Pt(big * Math.Cos((90 + 72 * i) * Math.PI / 180), r + big * Math.Sin((90 + 72 * i) * Math.PI / 180))).ToArray();
        d.Arc("CircleSegment", d.Outline("Frame", "#FFE8083A", 3), o, east, west);
        d.Circle(d.Fill("Orange", "#FFFCD29A", "#FFF54515", 3), center, o);
        d.Inked(d.Outline("Pentagon", Lgf.Ink, 3.5), corners);
        d.Scene(-R, R, R, 0);
    }

    // Four concentric circles and two rectangles, each between two parallel tangents of the circles:
    // the right one from the innermost circle's tangent to the second's, its inner corners on the
    // second circle and its outer ones on the third; the left one from the innermost circle's tangent to
    // the third's, its corners on the third and the fourth. The blue ring is 4, the purple? Dragging the
    // innermost circle's edge (the right rectangle's side) changes the rings, the outer circle stays:
    // the purple ring stays twice the blue.
    [Puzzle("rectangles-in-concentric-rings")]
    public static void RectanglesInConcentricRings(Lgf d)
    {
        const double R = 5;
        var o = d.Pt(0, 0, "O");
        var h = d.Slider(V.At(0.2 * R, 0), V.At(0.5 * R, 0), (0.35 - 0.2) / 0.3, (0.35 * R, 0), "H");
        string r1 = "H.X";
        string r2 = d.Num($"sqrt(({Lgf.N(R * R)} + 3 * H.X ^ 2) / 4)", "r2");
        string r3 = d.Num($"sqrt(2 * {r2} ^ 2 - H.X ^ 2)", "r3");
        string y1 = d.Num($"sqrt({r2} ^ 2 - H.X ^ 2)", "y1");
        string y2 = d.Num($"sqrt({r3} ^ 2 - H.X ^ 2)", "y2");
        var c1 = d.Pt(new V(r1, "0"));
        var c2 = d.Pt(new V(r2, "0"));
        var c3 = d.Pt(new V(r3, "0"));
        var c4 = d.Pt(R, 0);
        string first = d.Circle(null, o, c1);
        string second = d.Circle(null, o, c2);
        string third = d.Circle(null, o, c3);
        string fourth = d.Circle(null, o, c4);
        d.Region(d.Fill("Purple", "#FFAA80E8", null), "Difference", fourth, third);
        d.Region(d.Fill("Blue", "#FF8EDFF9", null), "Difference", second, first);
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        foreach (var c in new[] { c1, c2, c3, c4 })
        {
            d.Circle(frame, o, c);
        }

        string blue = d.Outline("Rectangle", "#FF0A08CC", 3);
        d.Inked(blue, d.Pt(new V(r1, $"-{y1}")), d.Pt(new V(r2, $"-{y1}")), d.Pt(new V(r2, y1)), d.Pt(new V(r1, y1)));
        d.Inked(blue, d.Pt(new V($"-{r3}", $"-{y2}")), d.Pt(new V($"-{r1}", $"-{y2}")), d.Pt(new V($"-{r1}", y2)), d.Pt(new V($"-{r3}", y2)));
        d.Scene(-R, R, R, -R);
    }

    // Three regular hexagons: the big one (area 9) has the left one's bottom right corner at its bottom
    // left corner and the right one's top left corner at its top right corner. The pink quadrilateral
    // from the left one's top left corner over the big one's top left corner and the right one's bottom
    // right corner to the big one's bottom right corner? Dragging the left and the right hexagons' far
    // corners resizes them: the pink stays two thirds of the big one.
    [Puzzle("three-regular-hexagons-ii")]
    public static void ThreeRegularHexagonsII(Lgf d)
    {
        const double b = 4;
        const string sqrt3 = "1.7320508075688772";
        var big = HexagonUp(d, V.At(0, 0), Lgf.N(b));
        // the left one's top left corner (-a, √3 a) and the right one's bottom right corner (b + c, √3 (b - c))
        var p = d.Slider(V.At(-0.45 * b, 0.45 * b * Math.Sqrt(3)), V.At(-1.1 * b, 1.1 * b * Math.Sqrt(3)), (0.89 - 0.45) / 0.65, (-0.89 * b, 0.89 * b * Math.Sqrt(3)), "P");
        var q = d.Slider(V.At(1.4 * b, 0.6 * b * Math.Sqrt(3)), V.At(2 * b, 0), (0.63 - 0.4) / 0.6, (1.63 * b, 0.37 * b * Math.Sqrt(3)), "Q");
        string a = d.Num("-P.X", "a");
        string c = d.Num($"Q.X - {Lgf.N(b)}", "c");
        var left = HexagonUp(d, new V($"-{a}", "0"), a);
        var right = HexagonUp(d, new V(Lgf.N(b), $"{sqrt3} * ({Lgf.N(b)} - {c})"), c);
        d.Polygon(d.Fill("Pink", "#FFF8C6F6", "#FFEA14A3"), p, big[4], q, big[1]);
        d.Polygon(d.Outline("LeftHexagon", "#FF1278EA", 3), left);
        d.Polygon(d.Outline("RightHexagon", "#FF7A1AE0", 3), right);
        d.Polygon(d.Outline("BigHexagon", "#FF1010DA", 3), big);
        d.Scene(-1.4 * b, 2 * b * S3, 2.1 * b, 0);
    }

    // A rectangle twice as wide as high with a semicircle on each side, the long ones reaching the
    // opposite side; the yellow rectangle has its corners where the short sides' semicircles cross the
    // long sides' ones. What fraction of the outer rectangle is shaded? Nothing to drag: the long
    // semicircles touching the far side fix the shape (the corners are a fifth of the way in).
    [Puzzle("one-rectangle-inside-another-with-semi-circles")]
    public static void OneRectangleInsideAnotherWithSemiCircles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double W = 10, H = 5;
        var bl = d.Pt(0, 0);
        var br = d.Pt(W, 0);
        var tr = d.Pt(W, H);
        var tl = d.Pt(0, H);
        var inner = new[] { d.Pt(W / 5, H / 5), d.Pt(4 * W / 5, H / 5), d.Pt(4 * W / 5, 4 * H / 5), d.Pt(W / 5, 4 * H / 5) };
        d.Polygon(d.Fill("Yellow", "#FFFFFE5C", null), inner);
        d.Arc("CircleArc", d.Line("Orange", "#FFF07010", 3), d.Pt(W / 2, 0), br, bl);
        d.Arc("CircleArc", d.Line("YellowLine", "#FFF2DE00", 3), d.Pt(W / 2, H), tl, tr);
        d.Arc("CircleArc", d.Line("Red", "#FFE8205A", 3), d.Pt(0, H / 2), bl, tl);
        d.Arc("CircleArc", d.Line("Blue", "#FF4A10D0", 3), d.Pt(W, H / 2), tr, br);
        string frame = d.Outline("Frame", Lgf.Ink, 3.5);
        d.Inked(frame, inner);
        d.Inked(frame, bl, br, tr, tl);
        d.Scene(0, H, W, 0);
    }

    // A semicircle, a green rectangle as high as it from the diameter's right end to a point of the arc,
    // and a red rectangle (3 wide) on the line from that end to that point. The green area? The
    // statement leaves where the point is and how long the red rectangle is; drawn as she drew it: the
    // point at 4/5 and 3/5 of the radius (left, up), the red rectangle twice as long as wide (so the
    // green, half the square of that line, is as big as the red). Nothing to drag.
    [Puzzle("two-rectangles-and-a-semi-circle")]
    public static void TwoRectanglesAndASemiCircle(Lgf d)
    {
        double R = Math.Sqrt(10);
        var o = d.Pt(0, 0);
        var e = d.Pt(R, 0);
        var w = d.Pt(-R, 0);
        var p = d.Pt(-0.8 * R, 0.6 * R);
        var f = d.Pt(1.3 * R, 0.9 * R);
        var g = d.Pt(-0.5 * R, 1.5 * R);
        d.Box(d.Fill("Green", "#FFDAFC96", "#FF9BD72A", 3), V.At(-0.8 * R, 0), V.At(R, R));
        // across the green (bright in the dark theme) and the paper: mid-tones there
        string blue = d.Line("Blue", "#FF1A3CA8", 3, dark: "#FF6B7BFF");
        d.Arc("CircleArc", blue, o, e, w);
        d.Segment(blue, w, e);
        string red = "#FFE8483C";
        d.Inked(d.Outline("Red", red, 3, dark: "#FFFF5A4F"), p, e, f, g);
        d.Dimension(d.Line("RedArrow", red, 2.5), d.Text("RedMark", red, 30), f, e, 0.35, "3", 0.4);
        d.Scene(-R, 1.5 * R, 1.3 * R + 0.8, 0);
    }

    // Three equilateral triangles of side 12 in a row on one line, each half a side along from the last;
    // a semicircle with an upright diameter from the first one's right side to the second one's left
    // side, touching the third one's left side. Its area? Nothing to drag: the triangles fix it.
    [Puzzle("three-equilateral-triangles-and-a-semi-circle")]
    public static void ThreeEquilateralTrianglesAndASemiCircle(Lgf d)
    {
        double h = 6 * Math.Sqrt(3), q = Math.Sqrt(3);
        var center = d.Pt(11, 3 * q);
        var bottom = d.Pt(11, q);
        var top = d.Pt(11, 5 * q);
        var a1 = d.Pt(0, 0);
        var a2 = d.Pt(12, 0);
        var apex1 = d.Pt(6, h);
        string lightGreen = "#FF8EE02A";
        d.Polygon(d.Outline("LightGreen", lightGreen, 3), a1, a2, apex1);
        d.Polygon(d.Outline("Green", "#FF0C9D4F", 3), d.Pt(6, 0), d.Pt(18, 0), d.Pt(12, h));
        d.Polygon(d.Outline("Brown", "#FF3A1C14", 3), a2, d.Pt(24, 0), d.Pt(18, h));
        d.Arc("CircleSegment", d.Fill("Blue", "#FF05DFFC", "#FF1A78D8", 3), center, bottom, top);
        d.Label("?", V.At(12.2, 3 * q), d.ShapeText("BlueMark", "#FF1A60C8", 36), 36);
        // her 12 beside an arrow along the first triangle's left side
        var n = V.At(-S3, 0.5);
        var from = d.Pt(a1 + 1.1 * n);
        var to = d.Pt(apex1 + 1.1 * n);
        var middle = Mid(d, from, to);
        string arrow = d.Line("LightGreenArrow", lightGreen, 3);
        d.Vector(arrow, middle, from);
        d.Vector(arrow, middle, to);
        d.Label("12", V.At(3, h / 2) + 2.3 * n, d.Text("LightGreenMark", lightGreen, 30), 30);
        d.Scene(-2.6, h, 24, 0);
    }

    /// <summary>A dimension along the segment from a to b itself: arrows out to its ends from a gap around a point of it, with the text in the gap</summary>
    private static void DimensionPartway(Lgf d, string lineStyle, string textStyle, V a, V b, V at, string text, double gap, double size)
    {
        string length = d.Num($"sqrt(({b.X} - {a.X}) ^ 2 + ({b.Y} - {a.Y}) ^ 2)");
        var along = d.Pt($"1 / {length}" * (b - a));
        var middle = d.Pt(at);
        var from = d.Pt(a);
        var to = d.Pt(b);
        d.Vector(lineStyle, d.Pt(middle - gap * along), from);
        d.Vector(lineStyle, d.Pt(middle + gap * along), to);
        d.Label(text, middle, textStyle, size);
    }
}
