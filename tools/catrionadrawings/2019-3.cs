#nullable disable

static partial class Puzzles
{
    // Two congruent rectangles packed in a square: one standing along its left side, full height, the
    // other tilted with its bottom corner on the square's bottom, its left corner on the first one's
    // side, its top corners on the square's top and right sides. The angle at its bottom corner?
    // Nothing to drag: packing them so fixes it (the tilted one at 60 degrees, both 2 - sqrt(3) of the
    // square wide).
    [Puzzle("two-congruent-rectangles-in-a-square")]
    public static void TwoCongruentRectanglesInASquare(Lgf d)
    {
        const double L = 6;
        double w = L * (2 - Math.Sqrt(3));
        string blue = "#FF0A10A8";
        var b = d.Pt(L / 2, 0, "B");
        var left = d.Pt(L / 2 - w * S3, w / 2, "Lc");
        var right = d.Pt(L, L * S3, "R");
        var top = d.Pt(L / 2 - w * S3 + L / 2, w / 2 + L * S3, "T");
        var east = d.Pt(L, 0, "E");
        string fill = d.Fill("Blue", "#FF05C2E8", blue, 3.5);
        d.Box(fill, V.At(0, 0), V.At(w, L));
        d.Polygon(fill, b, right, top, left);
        // inked after coloring: the square's sides run over the standing rectangle's
        d.Box(d.Outline("Frame", Lgf.Ink, 4), V.At(0, 0), V.At(L, L));
        d.ArcMark(d.Line("BlueLine", blue, 3.5), b, east, right, 0.16 * L);
        d.Label("?", V.At(0.755 * L, 0.17 * L), d.Text("BlueMark", blue, 26));
        d.Scene(0, L, L, 0);
    }

    // A rectangle cut by its diagonal and a line from the top right corner square onto it: three similar
    // right triangles, their incircles of diameters 3, 4 and 5. The rectangle's area? Nothing to drag:
    // the diameters fix it (7.5 by 10).
    [Puzzle("three-circles-in-a-rectangle-ii")]
    public static void ThreeCirclesInARectangleII(Lgf d)
    {
        const double W = 7.5, H = 10;
        var bl = d.Pt(0, 0, "BL");
        var br = d.Pt(W, 0, "BR");
        var tr = d.Pt(W, H, "TR");
        var tl = d.Pt(0, H, "TL");
        // the foot of the line from the top right corner, 4.5 along the diagonal
        var f = d.Pt(2.7, 6.4, "F");
        string ink = d.Line("Ink");
        d.Polygon(d.Outline("Frame"), bl, br, tr, tl);
        d.Segment(ink, tl, br);
        d.Segment(ink, tr, f);
        foreach (var (x, y, r, name, color) in new[]
        {
            (3.0, 8.5, 1.5, "Yellow", "#FFF0D000"),
            (5.5, 6.0, 2.0, "Blue", "#FF1E9BE6"),
            (2.5, 2.5, 2.5, "Red", "#FFE01E1E"),
        })
        {
            var center = d.Pt(x, y);
            var left = d.Pt(x - r, y);
            var right = d.Pt(x + r, y);
            d.Circle(d.Outline(name, color), center, right);
            d.Dimension(d.Line(name + "Arrow", color, 3), d.Text(name + "Mark", color, 32), left, right, 0, Lgf.N(2 * r), 0.36);
        }

        d.Scene(0, H, W, 0);
    }

    // A circle, a square with its top corners on it and its bottom corners on the sides of the
    // equilateral triangle inscribed point down, and a second equilateral triangle on the square's top
    // side, point down. The yellow part of the big triangle under the square is 24; the overlap of the
    // two triangles? Nothing to drag: the circle fixes it all (the square's side is the radius, the small
    // triangle's point the center).
    [Puzzle("overlapping-triangles-in-a-circle")]
    public static void OverlappingTrianglesInACircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4;
        double r3 = Math.Sqrt(3), bottom = 2 * r3 - R;
        string orange = "#FFF08A24";
        var o = d.Pt(0, 0, "O");
        var rim = d.Pt(0, -R, "P");
        var a = d.Pt(-2 * r3, 2, "A");
        var b = d.Pt(2 * r3, 2, "Bt");
        var tl = d.Pt(-2, 2 * r3, "TL");
        var tr = d.Pt(2, 2 * r3, "TR");
        var bl = d.Pt(-2, bottom, "BL");
        var br = d.Pt(2, bottom, "BR");
        string yellow = d.Fill("Yellow", "#FFFFF901", null);
        d.Polygon(yellow, bl, br, rim);
        d.Polygon(yellow, d.Pt(-2 / r3, 2), d.Pt(2 / r3, 2), o);
        d.Polygon(d.Outline("Square", orange), tl, tr, br, bl);
        d.Circle(d.Outline("Circle", orange), o, rim);
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        d.Polygon(frame, a, b, rim);
        d.Polygon(frame, tl, tr, o);
        d.Label("24", V.At(0, -1.9), d.ShapeText());
        d.Label("?", V.At(0, 1.35), d.ShapeText());
        d.Scene(-R, R, R, -R);
    }

    // A square's diagonal with a point on it, a the part to the top corner, c the part to the bottom
    // corner and b the line to the other bottom corner, with a + b = c. The angle between b and c?
    // Nothing to drag: the condition fixes the point (1/2 + sqrt(3)/6 of the way up).
    [Puzzle("diagonal-in-a-square")]
    public static void DiagonalInASquare(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 6;
        double t = L * (0.5 + Math.Sqrt(3) / 6);
        string red = "#FFE0284A", cyan = "#FF28C0E8", green = "#FF9ACD32";
        var bl = d.Pt(0, 0, "BL");
        var br = d.Pt(L, 0, "BR");
        var tr = d.Pt(L, L, "TR");
        var p = d.Pt(t, t, "P");
        d.Box(d.Outline("Frame", "#FF333D4A", 4), V.At(0, 0), V.At(L, L));
        d.Segment(d.Line("Red", red, 4), p, tr);
        d.Segment(d.Line("Cyan", cyan, 4), bl, p);
        d.Segment(d.Line("Green", green, 4), p, br);
        d.DotAt(p, "Pd", size: 12);
        string ink = d.Line("Ink", Lgf.Ink, 3);
        d.ArcMark(ink, p, bl, br, 0.5);
        d.Label("?", p + V.At(-0.3, -1.05), d.Text("Mark", Lgf.Ink, 30), 30);
        d.Label("a", V.At(0.81 * L, 0.89 * L), d.Text("RedMark", red, 30), 30);
        d.Label("c", V.At(0.32 * L, 0.46 * L), d.Text("CyanMark", cyan, 30), 30);
        d.Label("b", V.At(0.81 * L, 0.38 * L), d.Text("GreenMark", green, 30), 30);
        d.Scene(0, L, L, 0);
    }

    // Seven circles, their centers marked: a big one, two blue ones on its horizontal diameter through
    // its center, a black one through the blue centers, a ring around that one of width 1, and two pink
    // ones with their centers on the ring's outer edge, touching the big circle and the blue ones. The
    // ring's area? Nothing to drag: the touching fixes the size (the big circle's radius is 6).
    [Puzzle("seven-circles")]
    public static void SevenCircles(Lgf d)
    {
        const double R = 6, inner = R / 2, outer = inner + 1;
        string navy = "#FF1A14A0", crimson = "#FFB8306A";
        var o = d.Pt(0, 0, "O");
        string innerCircle = d.Circle(null, o, d.Pt(inner, 0));
        string outerCircle = d.Circle(null, o, d.Pt(outer, 0));
        d.Region(d.Fill("Pink", "#FFFCC8F8", null), "Difference", outerCircle, innerCircle);
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        d.Circle(frame, o, d.Pt(R, 0));
        d.Circle(frame, o, d.Pt(inner, 0));
        d.Circle(frame, o, d.Pt(outer, 0));
        string blue = d.Outline("Blue", navy, 3);
        string red = d.Outline("Crimson", crimson, 3);
        var west = d.Pt(-inner, 0, "W");
        var east = d.Pt(inner, 0, "E");
        var north = d.Pt(0, outer, "N");
        var south = d.Pt(0, -outer, "S");
        d.Circle(blue, west, o);
        d.Circle(blue, east, o);
        d.Circle(red, north, d.Pt(0, R));
        d.Circle(red, south, d.Pt(0, -R));
        foreach (var center in new[] { o, west, east, north, south })
        {
            d.DotAt(center);
        }

        // the width across the ring, up and to the right, the 1 above it
        double angle = 28 * Math.PI / 180, cos = Math.Cos(angle), sin = Math.Sin(angle);
        var mid = d.Pt((inner + outer) / 2 * cos, (inner + outer) / 2 * sin);
        string arrow = d.Line("Arrow", Lgf.Ink, 2.5);
        d.Vector(arrow, mid, d.Pt(inner * cos, inner * sin));
        d.Vector(arrow, mid, d.Pt(outer * cos, outer * sin));
        d.Label("1", V.At(3.25 * cos - 0.45 * sin, 3.25 * sin + 0.45 * cos), d.Text("Mark", Lgf.Ink, 22), 22);
        d.Scene(-R, R, R, -R);
    }

    // Two regular hexagons with vertical sides, the small one's lower left corner on the big one's lower
    // right corner and its left side along the big one's right side; lines from the big one's lower left
    // corner to the small one's top and from the big one's top to the small one's lower right corner.
    // The sum of the two blue angles, each between a line and a vertical side? Dragging the small
    // hexagon's upper right corner resizes it: the angles change, their sum stays 120 degrees.
    [Puzzle("angles-in-two-hexagons")]
    public static void AnglesInTwoHexagons(Lgf d)
    {
        const double A = 4;
        var big = Enumerable.Range(0, 6).Select(k => d.Pt(A * Math.Cos((90 + 60 * k) * Math.PI / 180), A * Math.Sin((90 + 60 * k) * Math.PI / 180))).ToArray();
        // big: top, upper left, lower left, bottom, lower right, upper right
        var corner = big[4];
        var step = V.At(2 * S3, 1);
        var h = d.Slider(corner + 1.2 * step, corner + 3.7 * step, (2.6 - 1.2) / 2.5, (2 * A * S3 + 2.6 * 2 * S3 - A * S3, -2 + 2.6), "H");
        string b = d.Num($"(H.Y + {Lgf.N(A / 2)})", "b");
        var small = new[]
        {
            corner,
            d.Pt(corner + new V("0", b)),
            d.Pt(corner + new V($"{s3} * {b}", $"1.5 * {b}"), "Ts"),
            h,
            d.Pt(corner + new V($"2 * {s3} * {b}", "0"), "Rs"),
            d.Pt(corner + new V($"{s3} * {b}", $"-0.5 * {b}")),
        };
        string green = d.Outline("Green", "#FF10C828", 3.5);
        d.Polygon(green, big);
        d.Polygon(green, small);
        string ink = d.Line("Ink", Lgf.Ink, 4);
        d.Segment(ink, big[2], small[2]);
        d.Segment(ink, big[0], small[4]);
        string blue = d.Line("Blue", "#FF0204EC", 4);
        d.ArcMark(blue, big[2], small[2], big[1], 0.95);
        d.ArcMark(blue, small[4], small[3], big[0], 0.95);
        d.Scene(-A * S3, A, A * S3 + 3 * 2 * S3, -A);
    }

    // A big square; on its top side the diagonal of a square of side 6 from its top left corner, and a
    // square of side 7 with a corner at its top right corner and the opposite corner on the small
    // square's lower right side. The length from the small square's right corner to the 7-square's
    // bottom corner? Dragging that opposite corner along the side turns the 7-square and resizes the
    // big one: the length stays 7 (the right corner sees the 7-square's diagonal at 135 degrees, so it
    // is on the circle of radius 7 around that bottom corner).
    [Puzzle("three-overlapping-squares")]
    public static void ThreeOverlappingSquares(Lgf d)
    {
        d.Texture = "FeltTip";
        double r2 = Math.Sqrt(2), diagonal = 6 * r2;
        string navy = "#FF1A1F5E", blue = "#FF2448C0", cyan = "#FF30C0E8", red = "#FFE83060";
        var tl = d.Pt(0, 0, "TL");
        var p = d.Pt(diagonal, 0, "P");
        var top = d.Pt(diagonal / 2, diagonal / 2);
        var bottom = d.Pt(diagonal / 2, -diagonal / 2, "D");
        var down = V.At(-1 / r2, -1 / r2);
        var c = d.Slider(p + 3 * down, p + 5.8 * down, (4.2 - 3) / 2.8, (diagonal - 4.2 / r2, -4.2 / r2), "C");
        var t = d.Pt(new V($"C.X + sqrt(98 - C.Y ^ 2)", "0"), "T");
        string s = "T.X";
        var m = Mid(d, c, t);
        var half = t - m;
        var q = d.Pt(m + new V(half.Y, V.Neg(half.X)), "Q");
        var u = d.Pt(m - new V(half.Y, V.Neg(half.X)), "U");
        d.Polygon(d.Outline("Frame", navy, 4), tl, t, d.Pt(new V(s, $"-{s}")), d.Pt(new V("0", $"-{s}")));
        d.Polygon(d.Outline("Blue", blue, 3), tl, top, p, bottom);
        d.Polygon(d.Outline("Cyan", cyan, 3), t, u, c, q);
        d.Dimension(d.Line("RedArrow", red, 3), d.Text("RedMark", red, 28), p, q, 0, "?", 0.55);
        d.Label("6", V.At(diagonal / 4 - 0.75, diagonal / 4 + 0.75), d.Text("BlueMark", blue, 30), 30);
        d.Label("7", d.Pt(Mid(d, u, t) + 0.11 * (u - q)), d.Text("CyanMark", cyan, 30), 30);
        d.Scene(-0.3, 5.2, 15.5, -15.5);
    }

    // Four squares on a semicircle's diameter: the big yellow one, the green one in its bottom left
    // corner, the orange one from the green one to the yellow one's right side, the red one in the
    // orange one's corner; the green one's and the red one's top left corners and the orange one's top
    // right corner on the arc. The yellow one is 18, the total? With the green corner at (a, g) on the
    // circle the yellow side is g - a and the red one's is sqrt(-2 a g): dragging that corner along the
    // arc changes all four, their total stays twice the yellow.
    [Puzzle("four-squares-and-a-semi-circle")]
    public static void FourSquaresAndASemiCircle(Lgf d)
    {
        const double R = 6;
        string teal = "#FF0A667A";
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0, "E");
        var west = d.Pt(-R, 0, "W");
        var g = d.ArcSlider(o, R, 156, 174, 165, (0, 0), "G");
        string a = "G.X", h = "G.Y";
        string r = d.Num($"sqrt(-2 * {a} * {h})", "r");
        // the fills stay light in the dark theme: the outlines between them stay ink there
        d.Box(d.Fill("Yellow", "#FFFFFE66", width: 3.5, darkOutline: Lgf.Ink), new V(a, "0"), new V(h, $"{h} - {a}"));
        d.Box(d.Fill("Orange", "#FFFED65F", width: 3.5, darkOutline: Lgf.Ink), new V($"{a} + {h}", "0"), new V(h, V.Neg(a)));
        d.Box(d.Fill("Green", "#FFCCF896", width: 3.5, darkOutline: Lgf.Ink), new V(a, "0"), new V($"{a} + {h}", h));
        d.Box(d.Fill("Red", "#FFFB9090", width: 3.5, darkOutline: Lgf.Ink), new V($"{a} + {h}", "0"), new V($"{a} + {h} + {r}", r));
        string line = d.Line("Teal", teal, 3.5);
        d.Arc("CircleArc", line, o, east, west);
        d.Segment(line, west, d.Pt(new V(a, "0")));
        d.Segment(line, d.Pt(new V(h, "0")), east);
        d.Scene(-R, 7.5, R, 0);
    }

    // Two congruent regular hexagons with vertical sides in a rectangle, each touching its top and
    // bottom and one of its sides, the left one's lower right corner and the right one's upper left
    // corner on the rectangle's diagonal. What fraction of the rectangle is their overlap? Nothing to
    // drag: that fixes the rectangle (4 sqrt(3) / 3 of the hexagon's side wide for every 2 high).
    [Puzzle("two-overlapping-hexagons-and-a-rectangle")]
    public static void TwoOverlappingHexagonsAndARectangle(Lgf d)
    {
        const double s = 3, H = 2 * s, W = 8.0 / 3 * S3 * s;
        string lime = "#FF8CE84A";
        V[] Hexagon(double cx) => Enumerable.Range(0, 6)
            .Select(k => d.Pt(cx + s * Math.Cos((90 + 60 * k) * Math.PI / 180), s + s * Math.Sin((90 + 60 * k) * Math.PI / 180)))
            .ToArray();
        var left = Hexagon(S3 * s);
        var right = Hexagon(W - S3 * s);
        d.Region(d.Fill("Green", "#FFC4FCA4", null), "Intersection", d.Polygon(null, left), d.Polygon(null, right));
        string frame = d.Outline("Frame", Lgf.Ink, 4);
        d.Polygon(frame, left);
        d.Polygon(frame, right);
        var tl = d.Pt(0, H);
        var br = d.Pt(W, 0);
        d.Polygon(d.Outline("Lime", lime, 3), d.Pt(0, 0), br, d.Pt(W, H), tl);
        d.Segment(d.Line("LimeLine", lime, 3), tl, br);
        d.Scene(0, H, W, 0);
    }

    // A pink and a blue rectangle of equal area on one line, sharing their right side, and the line from
    // the pink one's bottom left corner to the blue one's top right corner, which cuts two purple
    // triangles off by the pink one's top and the blue one's sides: 12 and 3. The area? The triangles are
    // similar, twice as big: with the small one's legs u and v (u v = 6) the pink one is 4u by 3v, the
    // blue one 3u by 4v. The handle on the diagonal turns it: the shapes change, both areas stay 72.
    [Puzzle("two-rectangles-of-equal-area")]
    public static void TwoRectanglesOfEqualArea(Lgf d)
    {
        string navy = "#FF1E2A6E";
        var o = d.Pt(0, 0, "O");
        var h = d.ArcSlider(o, 2.6, 25, 38, 28.4, (0, 0), "H");
        string u = d.Num("sqrt(6 * H.X / H.Y)", "u");
        string v = d.Num("sqrt(6 * H.Y / H.X)", "v");
        V At(double i, double j) => d.Pt(new V(V.Times(Lgf.N(i), u), V.Times(Lgf.N(j), v)));
        string purple = d.Fill("Purple", "#FFB898F8", null);
        d.Polygon(purple, At(1, 3), At(3, 3), At(1, 1));
        d.Polygon(purple, At(3, 3), At(4, 3), At(4, 4));
        d.Polygon(d.Outline("Pink", "#FFFC58F2", 3), o, At(4, 0), At(4, 3), At(0, 3));
        d.Polygon(d.Outline("Blue", "#FF0848F5", 3), At(1, 0), At(4, 0), At(4, 4), At(1, 4));
        d.Segment(d.Line("Navy", navy, 3.5), o, At(4, 4));
        string text = d.ShapeText("NavyMark", navy, 28);
        d.Label("12", new V($"5 * {u} / 3", $"7 * {v} / 3"), text, 28);
        d.Label("3", new V($"11 * {u} / 3", $"10 * {v} / 3"), text, 28);
        d.Scene(0, 7.2, 13.6, 0);
    }

    // A red square on a semicircle's diameter at its left end, and two squares of area 4 side by side,
    // tilted, from its top left corner: the far corners of their top side are on the arc (the shared one
    // and the yellow one's). The red square's area? The power of the red corner says its side squared is
    // 2 times 4. Dragging the yellow corner turns the pair, the semicircle grows or shrinks to stay
    // through both corners: the red square stays 8.
    [Puzzle("three-squares-overlapping-a-semi-circle")]
    public static void ThreeSquaresOverlappingASemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double a = 2;
        double L = a * Math.Sqrt(2);
        var bl = d.Pt(0, 0, "BL");
        var a0 = d.Pt(0, L, "A");
        var b = d.ArcSlider(a0, 2 * a, 2, 22, 9.8, (0, L), "B");
        // the pair's top side, along and across (down)
        var e = new V($"B.X / {Lgf.N(2 * a)}", $"(B.Y - {Lgf.N(L)}) / {Lgf.N(2 * a)}");
        var n = new V(e.Y, V.Neg(e.X));
        string r = d.Num($"(1.5 * {Lgf.N(a)} + {Lgf.N(L)} * {e.Y}) / ({e.X})", "R");
        var o = d.Pt(new V(r, "0"), "O");
        var east = d.Pt(new V($"2 * {r}", "0"), "E");
        var m = d.Pt(a0 + a * e, "M");
        var dm = d.Pt(m + a * n, "Dm");
        var da = d.Pt(a0 + a * n, "Da");
        var db = d.Pt(b + a * n, "Db");
        var br = d.Pt(L, 0);
        var tr = d.Pt(L, L);
        d.Polygon(d.Fill("Red", "#FFFC1A04", null), bl, br, tr, a0);
        d.Polygon(d.Fill("Orange", "#FFFCA704", null), a0, m, dm, da);
        d.Polygon(d.Fill("Yellow", "#FFFFFF01", null), m, b, db, dm);
        d.Arc("CircleSegment", d.Outline("Semicircle", Lgf.Ink, 3.5), o, east, bl);
        // inked after coloring: every outline over the fills and the arc
        string frame = d.Outline("Frame", Lgf.Ink, 3.5);
        d.Inked(frame, bl, br, tr, a0);
        d.Inked(frame, a0, m, dm, da);
        d.Inked(frame, m, b, db, dm);
        d.Scene(0, 3.7, 7.1, 0);
    }

    // Two regular hexagons with flat tops, one in the other, and a square of side 2 between them: its
    // corners on the big one's slanted sides, the small one's top and bottom along its top and bottom.
    // The shaded area between the hexagons? Nothing to drag: the square fixes both (sides 1 + 1/sqrt(3)
    // and 2/sqrt(3)).
    [Puzzle("two-hexagons-and-a-square")]
    public static void TwoHexagonsAndASquare(Lgf d)
    {
        double outer = 1 + 1 / Math.Sqrt(3), inner = 2 / Math.Sqrt(3);
        string blue = "#FF1E5AE6";
        V[] Hexagon(double side) => Enumerable.Range(0, 6)
            .Select(k => d.Pt(side * Math.Cos(60 * k * Math.PI / 180), side * Math.Sin(60 * k * Math.PI / 180)))
            .ToArray();
        var big = Hexagon(outer);
        var small = Hexagon(inner);
        d.Region(d.Fill("Cyan", "#FF04C8FD", null), "Difference", d.Polygon(null, big), d.Polygon(null, small));
        string frame = d.Outline("Blue", blue, 3.5);
        d.Polygon(frame, big);
        d.Polygon(frame, small);
        d.Box(d.Outline("Square", "#FF0A1C8C", 3.5), V.At(-1, -1), V.At(1, 1));
        d.Scene(-outer, outer * S3, outer, -outer * S3);
    }

    // Four regular hexagons: a big white one; a yellow one with a corner at the big one's top left
    // corner; the orange one (12) with a long diagonal from the yellow one's second corner clockwise to
    // the big one's right corner; the pink one with a side from the yellow one's second corner the other
    // way to the big one's bottom left corner. The pink area? Both of those spans are sqrt(3) times the
    // same span from the big one's top left corner, turned: the pink side is the orange one's long
    // diagonal, the pink area four times 12. Dragging the yellow one's opposite corner turns and resizes
    // it: the orange and pink ones change, the pink stays four times the orange.
    [Puzzle("four-hexagons")]
    public static void FourHexagons(Lgf d)
    {
        d.Texture = "FeltTip";
        const double S = 4;
        var big = Enumerable.Range(0, 6).Select(k => d.Pt(S * Math.Cos(60 * k * Math.PI / 180), S * Math.Sin(60 * k * Math.PI / 180))).ToArray();
        // big: right, top right, top left, left, bottom left, bottom right
        var t = big[2];
        var r = big[0];
        var bl = big[4];
        string range = d.Circle(null, d.Pt(-0.74, -2.73), d.Pt(-0.14, -2.73));
        var opposite = d.HandleIn(range, -0.74, -2.73, "Dy");
        var yc = Mid(d, t, opposite);
        // the yellow one's corners clockwise from the big one's top left corner
        var yellow = Enumerable.Range(0, 6).Select(k => k == 0 ? t : k == 3 ? opposite : d.Pt(yc + (t - yc).Turn(-60 * k))).ToArray();
        var c = yellow[2];
        var e = yellow[4];
        var oc = Mid(d, c, r);
        var orange = Enumerable.Range(0, 6).Select(k => k == 0 ? r : k == 3 ? c : d.Pt(oc + (r - oc).Turn(60 * k))).ToArray();
        var pc = d.Pt(bl + (e - bl).Turn(-60));
        var pink = Enumerable.Range(0, 6).Select(k => k == 0 ? bl : d.Pt(pc + (bl - pc).Turn(60 * k))).ToArray();
        d.Polygon(d.Fill("Yellow", "#FFFEF270", null), yellow);
        d.Polygon(d.Fill("Pink", "#FFD8007E", null), pink);
        d.Polygon(d.Fill("Orange", "#FFFEB89A", null), orange);
        // inked after coloring: every outline over the fills
        string frame = d.Outline("Frame", "#FF2A2A6E", 3);
        d.Polygon(frame, big);
        d.Polygon(frame, yellow);
        d.Polygon(frame, pink);
        d.Polygon(frame, orange);
        d.Label("?", pc, d.ShapeText());
        d.Label("12", oc, d.ShapeText());
        d.Scene(-S, S * S3, S, -S * S3);
    }
}
