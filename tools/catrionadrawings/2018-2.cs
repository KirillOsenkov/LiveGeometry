#nullable disable

static partial class Puzzles
{
    // An equilateral triangle in a circle, the circle in it, and a red circle touching the outer one at
    // the triangle's top corner and the triangle's base at its middle. What fraction is shaded (the
    // crescent between the outer and the red circle, and the inner circle)? Nothing to drag: the
    // triangle fixes it all (the red circle's radius is three quarters of the outer one's, the inner
    // circle's a half).
    [Puzzle("nested-circles-and-a-triangle")]
    public static void NestedCirclesAndATriangle(Lgf d)
    {
        const double R = 4;
        var o = d.Pt(0, 0);
        var top = d.Pt(0, R);
        var left = d.Pt(-R * S3, -R / 2);
        var right = d.Pt(R * S3, -R / 2);
        var foot = d.Pt(0, -R / 2);
        var redCenter = d.Pt(0, R / 4);
        string outer = d.Circle(null, o, top);
        string red = d.Circle(null, redCenter, top);
        string yellow = d.Fill("Yellow", "#FFFEEE45", null);
        d.Region(yellow, "Difference", outer, red);
        d.Circle(d.Fill("Inner", "#FFFEEE45", "#FFEE9A20"), o, foot);
        d.Circle(d.Outline("Red", "#FFE2231A", 3), redCenter, top);
        d.Circle(d.Outline("Pink", "#FFF8045B", 3), o, top);
        d.Inked(d.Outline("Frame", Lgf.Ink, 3), top, left, right);
        d.Scene(-R, R, R, -R);
    }

    // A square of side 3, a blue triangle from its bottom right corner to a point of its top and one of
    // its left side, and a white square of side 1, parallel to it, with three corners on the triangle's
    // sides. The blue area to the yellow? Dragging the triangle's top corner along the top: the white
    // square follows (its left side is half as far in as that corner, B.X / 2, and the triangle's left
    // corner is 3 / (3 - B.X) up), and the triangle's area stays 3, so blue to yellow stays 1 to 3.
    [Puzzle("square-in-a-triangle-in-a-square")]
    public static void SquareInATriangleInASquare(Lgf d)
    {
        var c = d.Pt(3, 0);
        var b = d.Slider(V.At(0.3, 3), V.At(1.7, 3), (0.86 - 0.3) / 1.4, (0.86, 3), "B");
        var a = d.Pt(new V("0", "3 / (3 - B.X)"), "A");
        string bottom = d.Num("(3 - B.X / 2) / (3 - B.X)", "y");
        var p = d.Pt(new V("B.X / 2", bottom), "P");
        var q = d.Pt(new V("B.X / 2 + 1", bottom), "Q");
        d.Box(d.Fill("Yellow", "#FFFEFE88"), V.At(0, 0), V.At(3, 3));
        d.Polygon(d.Fill("Blue", "#FFA4ECF4", "#FF19B5C9"), a, b, c);
        d.Box(d.PaperFill(), p, q + V.At(0, 1));
        d.Box(d.Outline("Frame", Lgf.Ink, 3.5), V.At(0, 0), V.At(3, 3));
        string ink = d.Line("Arrow");
        d.Dimension(ink, d.Text("Big", Lgf.Ink, 32), c, d.Pt(3, 3), -0.3, "3", 0.25, 32);
        d.Dimension(ink, d.Text("Small", Lgf.Ink, 24), p, q, 0.15, "1", 0.14, 24);
        d.Scene(0, 3, 3.45, -0.15);
    }

    // A rectangle, two quarter circles as high as it centered on its top left and bottom right
    // corners, and the black line touching both, through the middle. What fraction is shaded (the two
    // triangles between the diagonal, the black line and the long sides)? Nothing to drag: the line
    // touching both on one line through the middle makes the quarter circles touch there, so the
    // rectangle is √3 times as wide as high (her drawing is that shape).
    [Puzzle("quarter-circles-in-a-rectangle")]
    public static void QuarterCirclesInARectangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double h = 4;
        double w = h * Math.Sqrt(3), shift = h / (2 * Math.Sqrt(3));
        var bl = d.Pt(0, 0);
        var br = d.Pt(w, 0);
        var tr = d.Pt(w, h);
        var tl = d.Pt(0, h);
        var m = d.Pt(w / 2, h / 2);
        var foot = d.Pt(w / 2 - shift, 0);
        var head = d.Pt(w / 2 + shift, h);
        string green = d.Fill("Green", "#FFD8FD94", null);
        d.Polygon(green, bl, foot, m);
        d.Polygon(green, tr, head, m);
        string lime = "#FFC8E22A";
        d.Polygon(d.Outline("Frame", lime, 3), bl, br, tr, tl);
        d.Segment(d.Line("Diagonal", lime, 3), bl, tr);
        string cyan = d.Line("Cyan", "#FF10D8C0", 4);
        d.Arc("CircleArc", cyan, tl, bl, d.Pt(h, h));
        d.Arc("CircleArc", cyan, br, tr, d.Pt(w - h, 0));
        d.Segment(d.Line("Ink", Lgf.Ink, 3), foot, head);
        d.Scene(0, h, w, 0);
    }

    // A quarter circle and a semicircle on each of its radii; the lens where the semicircles overlap
    // is 4. The shaded area (what the semicircles leave of the quarter circle)? Nothing to drag: it is
    // all one shape.
    [Puzzle("semi-circles-in-a-quarter-circle")]
    public static void SemiCirclesInAQuarterCircle(Lgf d)
    {
        const double R = 6;
        var o = d.Pt(0, 0);
        var east = d.Pt(R, 0);
        var north = d.Pt(0, R);
        var up = d.Pt(0, R / 2);
        var right = d.Pt(R / 2, 0);
        string quarter = d.Arc("CircleSector", null, o, east, north, visible: false);
        string upright = d.Arc("CircleSegment", null, up, o, north, visible: false);
        string lying = d.Arc("CircleSegment", null, right, east, o, visible: false);
        d.Region(d.Fill("Blue", "#FF96DAF8", null), "Difference", quarter, upright, lying);
        string magenta = "#FFC84CAD";
        string line = d.Line("Magenta", magenta, 3);
        d.Arc("CircleArc", line, up, o, north);
        d.Arc("CircleArc", line, right, east, o);
        d.Arc("CircleSector", d.Outline("Frame", "#FF2A1E9B", 3.5), o, east, north);
        d.Label("4", V.At(1.3, 1.25), d.Text("MagentaMark", magenta, 44), 44);
        d.Scene(0, R, R, -0.6);
    }

    // Three squares on a line, of areas 20, 65 and 45: the tilted one stands on a corner, its side
    // corners on the outer top corners of the other two. The shaded area (the two triangles up to its
    // top corner)? Nothing to drag: the areas fix it (20 + 45 = 65 is what lets the corners meet).
    [Puzzle("three-squares-on-a-line")]
    public static void ThreeSquaresOnALine(Lgf d)
    {
        d.Texture = "FeltTip";
        double q = Math.Sqrt(20), p = Math.Sqrt(45);
        var b = d.Pt(0, 0);
        var l = d.Pt(-p, q);
        var r = d.Pt(q, p);
        var a = d.Pt(q - p, p + q);
        string orange = d.Fill("Orange", "#FFFCDC28", "#FFF8B020");
        d.Polygon(orange, d.Pt(-p - q, q), a, l);
        d.Polygon(orange, a, r, d.Pt(q + p, p));
        string blue = "#FF3679CC", pink = "#FFF08890", cyan = "#FF22D8E0";
        d.Box(d.Outline("Blue", blue, 3.5), V.At(-p - q, 0), V.At(-p, q));
        d.Box(d.Outline("Cyan", cyan, 3.5), V.At(q, 0), V.At(q + p, p));
        d.Polygon(d.Outline("Pink", pink, 2.5), b, r, a, l);
        d.Segment(d.Line("Ground", "#FFC8E820", 3), d.Pt(-p - q - 0.6, 0), d.Pt(q + p + 0.6, 0));
        d.Label("20", V.At(-p - q / 2, q / 2), d.Text("BlueMark", blue, 36), 36);
        d.Label("65", V.At((q - p) / 2, (p + q) / 2 + 0.6), d.Text("PinkMark", "#FFE8606C", 36), 36);
        d.Label("45", V.At(q + p / 2, p / 2), d.Text("CyanMark", cyan, 36), 36);
        d.Scene(-p - q - 0.6, p + q, q + p + 0.6, 0);
    }

    // Two circles, a rectangle and a regular hexagon packed inside each other: the rectangle's corners
    // on the outer circle, the hexagon's side corners on its sides and its top and bottom along them,
    // the inner circle touching the hexagon's sides. What fraction of the outer circle is shaded?
    // Nothing to drag: it is all one shape.
    [Puzzle("nested-circles-and-polygons")]
    public static void NestedCirclesAndPolygons(Lgf d)
    {
        const double a = 2;
        double h = a * S3, R = Math.Sqrt(a * a + h * h);
        var o = d.Pt(0, 0);
        d.Circle(d.Fill("Pink", "#FFFCD9F5", "#FFF20B01", 3.5), o, d.Pt(0, h));
        d.Circle(d.Outline("Lime", "#FFB6F210", 4), o, d.Pt(R, 0));
        var hexagon = Enumerable.Range(0, 6).Select(i => d.Pt(a * Math.Cos(i * Math.PI / 3), a * Math.Sin(i * Math.PI / 3))).ToArray();
        var corners = new[] { d.Pt(-a, -h), d.Pt(a, -h), d.Pt(a, h), d.Pt(-a, h) };
        d.Inked(d.Outline("Orange", "#FFF9B30C", 3), hexagon);
        d.Inked(d.Outline("Yellow", "#FFFCD807", 3.5), corners);
        // her pen's dots on the corners
        string green = d.Fill("GreenDot", "#FF0CE804", null);
        foreach (var corner in corners)
        {
            d.Circle(green, corner, d.Pt(corner + V.At(0.09, 0)));
        }

        string red = d.Fill("RedDot", "#FFE82010", null);
        foreach (var corner in hexagon)
        {
            d.Circle(red, corner, d.Pt(corner + V.At(0.05, 0)));
        }

        d.Scene(-R, R, R, -R);
    }

    // A square, the semicircle on its bottom side, and the tangents to it from the top corners, which
    // reach the far sides at a quarter of the way up. What fraction of the square is shaded (the two
    // triangles between the tangents and the sides)? Nothing to drag: it is all one shape.
    [Puzzle("semi-circle-and-tangents-in-a-square")]
    public static void SemiCircleAndTangentsInASquare(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 4;
        var bl = d.Pt(0, 0);
        var br = d.Pt(L, 0);
        var tr = d.Pt(L, L);
        var tl = d.Pt(0, L);
        var x = d.Pt(L / 2, L * 5 / 8);
        var left = d.Pt(0, L / 4);
        var right = d.Pt(L, L / 4);
        string yellow = d.Fill("Yellow", "#FFFEF99A", null);
        d.Polygon(yellow, tl, x, left);
        d.Polygon(yellow, tr, x, right);
        string line = d.Line("YellowLine", "#FFEACC07", 3);
        d.Segment(line, tl, right);
        d.Segment(line, tr, left);
        d.Arc("CircleArc", d.Line("Magenta", "#FFD3007A", 4), d.Pt(L / 2, 0), br, bl);
        d.Polygon(d.Outline("Frame", Lgf.Ink, 3.5), bl, br, tr, tl);
        d.Scene(0, L, L, -0.4);
    }

    // A regular hexagon and an equilateral triangle with its corners on every other side, each a third
    // of the way along (a and 2a). What proportion of the hexagon is shaded? Nothing to drag: the
    // given a : 2a fixes it.
    [Puzzle("equilateral-triangle-in-a-hexagon")]
    public static void EquilateralTriangleInAHexagon(Lgf d)
    {
        d.Texture = "FeltTip";
        var h = HexagonUp(d, V.At(0, 0), "3");
        var t = d.Pt(h[4] + (1.0 / 3) * (h[3] - h[4]));
        var r = d.Pt(h[2] + (1.0 / 3) * (h[1] - h[2]));
        var l = d.Pt(h[0] + (1.0 / 3) * (h[5] - h[0]));
        d.Polygon(d.Fill("Teal", "#FF4CD6E8", "#FF12B8C8"), t, r, l);
        string navy = "#FF3A38B8";
        d.Polygon(d.Outline("Frame", navy, 3.5), h);
        string text = d.Text("NavyMark", navy, 30);
        d.Label("a", V.At(0.5, 6 * S3 + 0.4), text, 30);
        d.Label("2a", V.At(2, 6 * S3 + 0.4), text, 30);
        d.Scene(-1.5, 6 * S3 + 0.7, 4.5, 0);
    }

    // A regular hexagon, the rectangle between its top and bottom sides, and two regular hexagons
    // standing on a corner in the rectangle: the big one as wide as it on its bottom, the small one
    // between its top corner and the rectangle's top. What fraction is shaded (the two hexagons and the
    // triangles beside the rectangle)? Nothing to drag: it is all one shape (the small hexagon is half
    // the big one's size).
    [Puzzle("three-regular-hexagons")]
    public static void ThreeRegularHexagons(Lgf d)
    {
        const double s = 4;
        double h = s * S3, t = s / (2 * S3), u = t / 2;
        var outer = HexagonFromLeft(d, V.At(-s, 0), Lgf.N(s));
        V[] Hexagon(double cy, double side) => Enumerable.Range(0, 6).Select(i => d.Pt(side * Math.Cos((90 + 60 * i) * Math.PI / 180), cy + side * Math.Sin((90 + 60 * i) * Math.PI / 180))).ToArray();
        string red = d.Fill("Red", "#FFF4806A", null);
        d.Polygon(red, outer[0], outer[1], outer[5]);
        d.Polygon(red, outer[3], outer[4], outer[2]);
        string frame = d.Fill("Hexagon", "#FFF4806A", Lgf.Ink, 3.5);
        d.Polygon(frame, Hexagon(-h + t, t));
        d.Polygon(frame, Hexagon(h - u, u));
        string darkRed = d.Line("DarkRed", "#FF8A0A06", 3);
        d.Segment(darkRed, outer[1], outer[5]);
        d.Segment(darkRed, outer[2], outer[4]);
        d.Polygon(d.Outline("Frame", Lgf.Ink, 3.5), outer);
        d.Scene(-s, h, s, -h);
    }

    // A right triangle in a semicircle, on its diameter, split by its height into triangles of areas 8
    // and 32. The semicircle's area? Nothing to drag: the areas fix it (the parts of the diameter are
    // 1 to 4, the height twice the short one).
    [Puzzle("triangles-in-a-semi-circle")]
    public static void TrianglesInASemiCircle(Lgf d)
    {
        double k = Math.Sqrt(2), R = 5 * k;
        var o = d.Pt(0, 0);
        var east = d.Pt(R, 0);
        var west = d.Pt(-R, 0);
        var foot = d.Pt(-R + 2 * k, 0);
        var apex = d.Pt(-R + 2 * k, 4 * k);
        string blue = "#FF0C40F0", cyan = "#FF31CCE3";
        string blueLine = d.Line("Blue", blue, 3.5);
        string cyanLine = d.Line("Cyan", cyan, 3.5);
        d.Segment(cyanLine, foot, east);
        d.Segment(cyanLine, apex, east);
        d.Segment(blueLine, west, foot);
        d.Segment(blueLine, west, apex);
        d.Segment(blueLine, foot, apex);
        d.Angle(blueLine, foot, west, apex, 14);
        d.Angle(cyanLine, foot, apex, east, 14);
        d.Arc("CircleSegment", d.Outline("Red", "#FFF21A14", 4), o, east, west);
        d.Label("8", V.At(-R + 1.38 * k, 1.25 * k), d.Text("BlueMark", blue, 34), 34);
        d.Label("32", V.At(-R + 4 * k, 1.6 * k), d.Text("CyanMark", cyan, 38), 38);
        d.Scene(-R, R, R, 0);
    }

    // A square, an equilateral triangle (12) on its bottom side, and a small upside down one along its
    // top from its top right corner down to the big one's side, so that its left side runs along that
    // side. The small one's area? Nothing to drag: the square fixes both.
    [Puzzle("equilateral-triangles-in-a-square")]
    public static void EquilateralTrianglesInASquare(Lgf d)
    {
        const double L = 4;
        double side = L / Math.Sqrt(3);
        var t = d.Pt(L - side, L);
        var tr = d.Pt(L, L);
        var v = d.Pt(L - side / 2, L - side * S3);
        var bl = d.Pt(0, 0);
        var br = d.Pt(L, 0);
        var apex = d.Pt(L / 2, L * S3);
        string violet = "#FFDA80D8";
        d.Polygon(d.Outline("Violet", violet, 3), t, tr, v);
        d.Polygon(d.Outline("Magenta", "#FFF40580", 3.5), bl, br, apex);
        d.Polygon(d.Outline("Frame", "#FF3A2010", 3.5), bl, br, tr, d.Pt(0, L));
        d.Label("12", V.At(L / 2, L * 0.3), d.Text("MagentaMark", "#FFF40580", 36), 36);
        d.Label("?", V.At(L - side / 2, L - side * S3 / 3 + 0.05), d.Text("VioletMark", "#FF9A3AAA", 36), 36);
        d.Scene(0, L, L, -0.4);
    }

    // Three equilateral triangles side by side on a line, of areas 9, 1 and 4, and the triangle between
    // their top corners. Its area? Nothing to drag: the areas fix it.
    [Puzzle("three-aligned-triangles")]
    public static void ThreeAlignedTriangles(Lgf d)
    {
        var a0 = d.Pt(0, 0);
        var a1 = d.Pt(3, 0);
        var a2 = d.Pt(4, 0);
        var a3 = d.Pt(6, 0);
        var big = d.Pt(1.5, 3 * S3);
        var small = d.Pt(3.5, S3);
        var right = d.Pt(5, 2 * S3);
        string orange = "#FFF3830B";
        string frame = d.Outline("Orange", orange, 3);
        d.Polygon(frame, a0, a1, big);
        d.Polygon(frame, a1, a2, small);
        d.Polygon(frame, a2, a3, right);
        d.Polygon(d.Fill("Pink", "#FFFDACD0", "#FFE8207A", 3.5), big, small, right);
        string text = d.Text("OrangeMark", orange, 34);
        d.Label("9", V.At(1.5, 1.05), text, 34);
        d.Label("1", V.At(3.5, 0.3), d.Text("OrangeSmall", orange, 24), 24);
        d.Label("4", V.At(5, 0.7), text, 34);
        d.Scene(0, 3 * S3, 6, 0);
    }
}
