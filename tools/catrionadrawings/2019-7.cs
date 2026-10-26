#nullable disable

static partial class Puzzles
{
    // A regular hexagon (flat top and bottom), its long horizontal diagonal cut in three equal parts
    // (ticked), and two parallel blue lines through the two cut points from the top side to the bottom
    // one: she drew the right one from the top right corner, so the left one ends at the bottom left
    // corner. What fraction is shaded (the outer parts)? Dragging the right line's top end along the top
    // side tilts both lines (the left one stays parallel, through the other cut point): the white part
    // is always a parallelogram a third of the diagonal wide and the hexagon high, so the shaded
    // fraction stays 5/9.
    [Puzzle("divided-hexagon-ii")]
    public static void DividedHexagonII(Lgf d)
    {
        const double s = 4, h = 2 * S3 * s;
        var left = d.Pt(-s, 0);
        var right = d.Pt(s, 0);
        var topLeft = d.Pt(-s / 2, h / 2);
        var topRight = d.Pt(s / 2, h / 2);
        var bottomLeft = d.Pt(-s / 2, -h / 2);
        var bottomRight = d.Pt(s / 2, -h / 2);
        var p = d.Slider(V.At(s / 6, h / 2), V.At(s / 2, h / 2), 1, (s / 2, h / 2), "P");
        var q = d.Pt(new V($"{Lgf.N(2 * s / 3)} - P.X", Lgf.N(-h / 2)), "Q");
        var leftTop = d.Pt(new V($"P.X - {Lgf.N(2 * s / 3)}", Lgf.N(h / 2)), "Lt");
        var leftBottom = d.Pt(new V("-P.X", Lgf.N(-h / 2)), "Lb");
        var third = d.Pt(-s / 3, 0);
        var twoThirds = d.Pt(s / 3, 0);
        string blue = d.Fill("Blue", "#FFC8F4FB", null);
        d.Polygon(blue, left, topLeft, leftTop, leftBottom, bottomLeft);
        d.Polygon(blue, p, topRight, right, bottomRight, q);
        string cyan = d.Line("Cyan", "#FF10D4F6", 3);
        d.Segment(cyan, left, third, "TwoTicks");
        d.Segment(cyan, third, twoThirds, "TwoTicks");
        d.Segment(cyan, twoThirds, right, "TwoTicks");
        string line = d.Line("BlueLine", "#FF0505C0", 3.5);
        d.Segment(line, leftTop, leftBottom);
        d.Segment(line, p, q);
        d.Inked(d.Outline("Frame", Lgf.Ink, 4), left, topLeft, topRight, right, bottomRight, bottomLeft);
        d.Scene(-s, h / 2, s, -h / 2);
    }

    // Two quarter circles on one line, touching at a corner: the green one of radius 2, the yellow one
    // of radius 3, and a semicircle of radius 1 (its diameter level) touching both arcs. The angle at
    // the corner between the lines to the two points of contact? Nothing to drag: the radii fix it (the
    // centers make a 3-4-5 triangle, right-angled at the semicircle's center).
    [Puzzle("semi-circle-between-two-quarter-circles")]
    public static void SemiCircleBetweenTwoQuarterCircles(Lgf d)
    {
        d.Texture = "FeltTip";
        var o = d.Pt(0, 0, "O");
        var g = d.Pt(-2, 0, "G");
        var y = d.Pt(3, 0, "Y");
        // 3 from G, 4 from Y
        var c = d.Pt(-0.2, 2.4, "C");
        var t1 = d.Pt(-0.8, 1.6, "T1");
        var t2 = d.Pt(0.6, 1.8, "T2");
        var greenTop = d.Pt(-2, 2);
        var yellowTop = d.Pt(3, 3);
        d.Arc("CircleSector", d.Outline("Green", "#FF3CD420", 3.5), g, o, greenTop);
        d.Arc("CircleSector", d.Outline("Yellow", "#FFF5E60E", 3.5), y, yellowTop, o);
        d.Arc("CircleSegment", d.Outline("Purple", "#FF6644DE", 3.5), c, d.Pt(-1.2, 2.4), d.Pt(0.8, 2.4));
        string ink = d.Line("Ink", Lgf.Ink, 3.5);
        d.Segment(ink, o, t1);
        d.Segment(ink, o, t2);
        d.Angle(d.Line("MarkLine", Lgf.Ink, 3), o, t2, t1, 18);
        d.DotAt(o, "Od", 11);
        d.DotAt(t1, "T1d", 11);
        d.DotAt(t2, "T2d", 11);
        d.Label("?", V.At(-0.1, 0.65));
        string arrow = d.Line("Arrow", Lgf.Ink, 3);
        d.Dimension(arrow, d.Text(), d.Pt(-2, 0), greenTop, 0.22, "2", 0.3);
        d.Dimension(arrow, d.Text(), yellowTop, y, 0.22, "3", 0.3);
        d.Dimension(arrow, d.Text(), c, d.Pt(-0.2, 1.4), 0, "1", 0.22);
        d.Scene(-2.4, 3, 3.4, -0.1);
    }

    // A big square with a white equilateral triangle on one side, the rest yellow; a small pink square
    // in its top corner and a pink equilateral triangle with its apex at the small square's other
    // corner, its left vertex at the big square's bottom corner and its right one straight under the
    // small square's side (twice the small square's side). The small square has area 2; the yellow
    // area? Nothing to drag: the small square decides everything.
    [Puzzle("two-squares-and-two-equilateral-triangles-iii")]
    public static void TwoSquaresAndTwoEquilateralTrianglesIII(Lgf d)
    {
        d.Texture = "FeltTip";
        const double a = 2, k = 1 + 2 * S3;
        var c = d.Pt(0, 0, "C");
        var dd = d.Pt(-2 * a, -k * a, "D");
        var b = d.Pt(-k * a, 2 * a, "B");
        var corner = d.Pt(-2 * a - k * a, -k * a + 2 * a, "A");
        var w = d.Pt(-a - S3 * k * a, -k * a / 2 + S3 * 2 * a, "W");
        var p = d.Pt(-a, -a, "P");
        var r = d.Pt(0, -k * a, "R");
        string pink = d.Outline("Pink", "#FFE60580", 3.5);
        d.Box(pink, V.At(-a, -a), V.At(0, 0));
        d.Polygon(pink, dd, r, p);
        const string red = "#FFDD0A10", darkRed = "#FFFF4A4A";
        d.Polygon(d.Fill("Yellow", "#FFFFFF01", red, 3.5, darkOutline: darkRed), corner, b, c, w, dd);
        d.Segment(d.Line("Red", red, 3.5, dark: darkRed), c, dd);
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(-k * a - 2 * a, 2 * a, 0, -k * a - 0.9);
    }

    // A purple square with an equilateral triangle on top; a big square on the segment from the
    // triangle's apex to the square's bottom right corner, and a white equilateral triangle on that
    // side inside it, the rest of the big square pink. The small square has area 8; the pink area?
    // Nothing to drag: the small square decides everything.
    [Puzzle("two-squares-and-two-equilateral-triangles-ii")]
    public static void TwoSquaresAndTwoEquilateralTrianglesII(Lgf d)
    {
        d.Texture = "FeltTip";
        const double a = 4, k = (1 + S3) * a;
        var q = d.Pt(a, 0, "Q");
        var t = d.Pt(a / 2, k, "T");
        var n = V.At(k, a / 2);
        var u = d.Pt(t + n, "U");
        var v = d.Pt(q + n, "Vb");
        var w = d.Pt(V.At(3 * a / 4, k / 2) + S3 * n, "W");
        string purple = d.Outline("Purple", "#FF7F6AE8", 3);
        d.Box(purple, V.At(0, 0), V.At(a, a));
        d.Polygon(purple, d.Pt(0, a), t, d.Pt(a, a));
        const string navy = "#FF2A0494";
        d.Polygon(d.Fill("Pink", "#FFFD5ED8", navy, 3.5), t, u, v, q, w);
        d.Segment(d.Line("Navy", navy, 3.5), q, t);
        d.Scene(0, k + a / 2, a + k, 0);
    }

    // Two circles touching each other and a line. In the big one, from a point on it, lines to where
    // it touches the line and to where it touches the small circle: angle a; in the small one, from a
    // point on it, lines to the same two kinds of points: angle b. a + b? Dragging the big circle's
    // point round it, or where the small circle touches the line (which resizes the small circle, still
    // touching both): a and b change, a + b stays 90°.
    [Puzzle("angles-in-two-circles")]
    public static void AnglesInTwoCircles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 2;
        var p1 = d.Pt(0, 0, "P1");
        var o1 = d.Pt(0, R, "O1");
        var p2 = d.Slider(V.At(1.9, 0), V.At(3.1, 0), (2.75 - 1.9) / 1.2, (2.75, 0), "P2");
        string r = d.Num($"P2.X ^ 2 / {4 * R}", "r");
        var o2 = d.Pt(new V("P2.X", r), "O2");
        var t = d.Pt(new V($"{R} * P2.X / ({R} + {r})", $"{R} + {R} * ({r} - {R}) / ({R} + {r})"), "T");
        var a = d.ArcSlider(o1, R, 100, 250, 136, (0, R), "A");
        const double angleB = 65 * Math.PI / 180;
        var b = d.Pt(new V($"P2.X + {Lgf.N(Math.Cos(angleB))} * {r}", $"{r} + {Lgf.N(Math.Sin(angleB))} * {r}"), "B");
        var lineEnd = d.Pt(new V($"P2.X + {r} + 0.3", "0"));
        d.Segment(d.Line("Ground", "#FF8CCB1A", 3.5), d.Pt(-2.2, 0), lineEnd);
        string circle = d.Outline("Circle", "#FF038045", 3.5);
        d.Circle(circle, o1, p1);
        d.Circle(circle, o2, p2);
        string brown = d.Line("Brown", "#FF3F3414", 3.5);
        d.Segment(brown, a, p1);
        d.Segment(brown, a, t);
        d.Segment(brown, b, t);
        d.Segment(brown, b, p2);
        const string green = "#FF5CC000";
        string mark = d.Line("GreenMark", green, 3);
        d.ArcMark(mark, a, p1, t, 0.32);
        d.ArcMark(mark, b, t, p2, 0.32);
        string text = d.Text("GreenText", green, 28);
        string ua = d.Num("distance(A, P1)", "ua");
        string va = d.Num("distance(A, T)", "va");
        d.Label("a", a + "0.36" * new V($"(P1.X - A.X) / {ua} + (T.X - A.X) / {va}", $"(P1.Y - A.Y) / {ua} + (T.Y - A.Y) / {va}"), text, 28);
        string ub = d.Num("distance(B, T)", "ub");
        string vb = d.Num("distance(B, P2)", "vb");
        d.Label("b", b + "0.36" * new V($"(T.X - B.X) / {ub} + (P2.X - B.X) / {vb}", $"(T.Y - B.Y) / {ub} + (P2.Y - B.Y) / {vb}"), text, 28);
        d.Label("a+b=?", V.At(3.3, 3.1), text, 28);
        d.Scene(-2.2, 2 * R, 4.3, -0.4);
    }

    // A square, lines from two corners to the midpoints (dots) of the far sides; what fraction is
    // shaded? Nothing to drag: it is a square with its midpoints.
    [Puzzle("square-with-midpoints")]
    public static void SquareWithMidpoints(Lgf d)
    {
        d.Texture = "FeltTip";
        const double s = 4;
        var topLeft = d.Pt(0, s);
        var bottomLeft = d.Pt(0, 0);
        var bottomRight = d.Pt(s, 0);
        var bottom = d.Pt(s / 2, 0, "Mb");
        var right = d.Pt(s, s / 2, "Mr");
        var x = d.Pt(0.4 * s, 0.2 * s, "X");
        string orange = d.Fill("Orange", "#FFFEAC02", null);
        d.Polygon(orange, topLeft, x, bottomLeft);
        d.Polygon(orange, x, right, bottomRight, bottom);
        string ink = d.Line("Ink", Lgf.Ink, 5);
        d.Segment(ink, topLeft, bottom);
        d.Segment(ink, bottomLeft, right);
        d.Inked(d.Outline("Frame", Lgf.Ink, 5), bottomLeft, bottomRight, d.Pt(s, s), topLeft);
        d.DotAt(bottom, "MbDot", 20);
        d.DotAt(right, "MrDot", 20);
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(0, s, s, -0.4);
    }

    // Two squares of the same size, the second one on the first one's bottom right corner, turned so
    // that its opposite corner is on the first one's diagonal (red, drawn on to it). The angle there
    // between the diagonal and the second square's side? Nothing to drag: the corner on the diagonal
    // fixes the turn (the second square is turned 30°).
    [Puzzle("angle-formed-by-two-squares")]
    public static void AngleFormedByTwoSquares(Lgf d)
    {
        d.Texture = "FeltTip";
        const double s = 4, t = (1 + 2 * S3) / 2 * s;
        var o = d.Pt(0, 0, "O");
        var b = d.Pt(s, 0, "B");
        var p = d.Pt(t, t, "P");
        var right = d.Pt(s + S3 * s, s / 2, "R");
        var left = d.Pt(s / 2, S3 * s, "L");
        // a mid blue in the dark theme: its light tint would be as lavender as the purple's
        d.Box(d.Outline("Blue", "#FF0304F8", 3.5, dark: "#FF6B7BFF"), V.At(0, 0), V.At(s, s));
        d.Polygon(d.Outline("Purple", "#FF5010C4", 3.5), b, right, p, left);
        const string red = "#FFF00707";
        string redLine = d.Line("Red", red, 3.5);
        d.Segment(redLine, o, p);
        d.Angle(d.Line("RedMark", red, 3), p, o, right, 24);
        d.Label("?", p + V.At(0, -1), d.Text("RedText", red, 32), 32);
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(0, t, s + S3 * s, -0.4);
    }

    // A big circle around a "house": a square with an equilateral triangle on it, the triangle's apex
    // and the square's bottom corners on the circle; the circles inscribed in the square and in the
    // triangle are green. What fraction of the big circle is green? Nothing to drag: the house fits the
    // circle one way (the square's side is the big radius).
    [Puzzle("two-circles-inside-two-polygons-inside-a-circle")]
    public static void TwoCirclesInsideTwoPolygonsInsideACircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double s = 4, h = 2 * S3 * s / 2;
        var o = d.Pt(0, h, "O");
        var bottomLeft = d.Pt(-s / 2, 0);
        var bottomRight = d.Pt(s / 2, 0);
        var topRight = d.Pt(s / 2, s);
        var topLeft = d.Pt(-s / 2, s);
        var apex = d.Pt(0, s + h, "Apex");
        string green = d.Fill("Green", "#FFE0FE76", "#FF5EC43A", 3);
        d.Circle(green, d.Pt(0, s / 2), d.Pt(0, 0));
        double r = h / 3;
        d.Circle(green, d.Pt(0, s + r), d.Pt(0, s));
        string ink = d.Line("Ink", Lgf.Ink, 3);
        d.Segment(ink, bottomLeft, bottomRight);
        d.Segment(ink, bottomRight, topRight);
        d.Segment(ink, topRight, topLeft);
        d.Segment(ink, topLeft, bottomLeft);
        d.Segment(ink, topLeft, apex);
        d.Segment(ink, apex, topRight);
        d.Circle(d.Outline("Circle", "#FF0A500E", 3), o, apex);
        d.Scene(-s, s + h, s, h - s);
    }

    // A regular hexagon (flat top and bottom) and a point inside it joined to four corners: the blue
    // triangle on the top side is a quarter of the hexagon, which puts the point a quarter of the
    // hexagon's height above the bottom. What fraction is green (below the lines to the left and right
    // corners)? Dragging the point left and right at that height: the green part changes shape, it
    // stays a third.
    [Puzzle("quartered-hexagon")]
    public static void QuarteredHexagon(Lgf d)
    {
        const double s = 4, h = S3 * s;
        var left = d.Pt(-s, 0);
        var topLeft = d.Pt(-s / 2, h);
        var topRight = d.Pt(s / 2, h);
        var right = d.Pt(s, 0);
        var bottomRight = d.Pt(s / 2, -h);
        var bottomLeft = d.Pt(-s / 2, -h);
        var p = d.Slider(V.At(-0.6 * s, -h / 2), V.At(0.6 * s, -h / 2), 0.25, (-0.3 * s, -h / 2), "P");
        d.Polygon(d.Fill("Blue", "#FF64F4FB", null), topLeft, topRight, p);
        d.Polygon(d.Fill("Green", "#FFC8FD9E", null), left, p, right, bottomRight, bottomLeft);
        string line = d.Line("GreenLine", "#FF036B2D", 3);
        d.Segment(line, p, left);
        d.Segment(line, p, topLeft);
        d.Segment(line, p, topRight);
        d.Segment(line, p, right);
        d.Inked(d.Outline("Frame", Lgf.Ink, 3.5), left, topLeft, topRight, right, bottomRight, bottomLeft);
        d.Scene(-s, h, s, -h);
    }

    // A square of area 100, a semicircle on its bottom side, and a line from the bottom left corner
    // (dashed) cutting the semicircle and then the right side: a second semicircle stands on the part
    // between. The length of the tangent from the corner to the second semicircle (the arrow)? Dragging
    // where the line meets the right side turns it: the small semicircle grows or shrinks, the tangent
    // stays 10 (the corner's power: the two cuts' distances multiply to 100).
    [Puzzle("two-semi-circles-in-a-square")]
    public static void TwoSemiCirclesInASquare(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 10;
        var o = d.Pt(0, 0, "O");
        var b = d.Slider(V.At(L, 4.5), V.At(L, 8.7), (7.2 - 4.5) / 4.2, (L, 7.2), "B");
        string n = d.Num("sqrt(100 + B.Y ^ 2)", "n");
        string k = d.Num($"100 / {n} ^ 2", "k");
        var a = d.Pt(new V($"10 * {k}", $"B.Y * {k}"), "A");
        var m = d.Pt(new V("(A.X + B.X) / 2", "(A.Y + B.Y) / 2"), "M");
        string far = d.Num($"(100 / {n} + {n}) / 2", "m");
        string rho = d.Num($"({n} - 100 / {n}) / 2", "rho");
        // the tangent from O: 10 long, turned from the line by the angle whose sine is rho / m
        string cos = d.Num($"10 / {far}", "cosf");
        string sin = d.Num($"{rho} / {far}", "sinf");
        string u = $"(10 / {n})", v = $"(B.Y / {n})";
        var t = d.Pt(new V($"10 * ({u} * {cos} - {v} * {sin})", $"10 * ({u} * {sin} + {v} * {cos})"), "T");
        // where the tangent leaves the square: the top side, or the right one when it is turned low
        string reach = d.Num("min(10 / T.X, 10 / T.Y)", "reach");
        var e = d.Pt(new V($"{reach} * T.X", $"{reach} * T.Y"), "E");
        d.Box(d.Outline("Frame", Lgf.Ink, 4.5), V.At(0, 0), V.At(L, L));
        d.Arc("CircleArc", d.Line("Pink", "#FFD31D49", 3), d.Pt(L / 2, 0), d.Pt(L, 0), o);
        string orange = d.Line("Orange", "#FFF0901E", 3);
        d.Dashed(orange, o, a, 9);
        d.Segment(orange, a, b);
        d.Arc("CircleArc", d.Line("OrangeRed", "#FFE85A28", 3), m, b, a);
        const string purple = "#FF3202DD";
        string arrow = d.Line("Purple", purple, 3);
        var middle = d.Pt(new V("T.X / 2", "T.Y / 2"));
        d.Vector(arrow, middle, o);
        d.Vector(arrow, middle, t);
        d.Dashed(arrow, t, e, 3);
        d.Label("?", new V("0.6 * T.X - 0.1 * T.Y", "0.6 * T.Y + 0.1 * T.X"), d.Text("PurpleText", purple, 30), 30);
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(0, L, L, -0.8);
    }

    // A blue triangle of area 6, each side drawn on past a corner by its own length, the three ends
    // making the red triangle; the semicircles stand on the doubled sides, centered on the corners.
    // The red triangle's area? Dragging the blue triangle's top corner changes its shape: the red
    // triangle stays seven times the blue one.
    [Puzzle("three-semi-circles-around-a-triangle")]
    public static void ThreeSemiCirclesAroundATriangle(Lgf d)
    {
        var p1 = d.Pt(0, 0, "P1");
        var p2 = d.Pt(3.25, 0, "P2");
        // kept where the top semicircle stays under the scene's top and the triangle stays a triangle
        string range = d.Polygon(null, d.Pt(1.3, 1.2), d.Pt(3.1, 1.2), d.Pt(3.1, 1.7), d.Pt(2.3, 2.1), d.Pt(1.3, 2.35));
        var p3 = d.HandleIn(range, 2.33, 2.05, "P3");
        var a = d.Pt(2 * p1 - p2, "A");
        var b = d.Pt(2 * p3 - p1, "B");
        var c = d.Pt(2 * p2 - p3, "C");
        d.Polygon(d.Fill("Blue", "#FF95F8FC", null), p1, p2, p3);
        const string navy = "#FF000863";
        string line = d.Line("Navy", navy, 3);
        d.Arc("CircleArc", line, p1, a, p2);
        d.Arc("CircleArc", line, p3, b, p1);
        d.Arc("CircleArc", line, p2, c, p3);
        d.Segment(line, a, p2);
        d.Segment(line, p1, b);
        d.Segment(line, p3, c);
        d.Inked(d.Outline("Red", "#FFD40404", 3.5), a, c, b);
        d.DotAt(a, "Ad");
        d.DotAt(b, "Bd");
        d.DotAt(c, "Cd");
        d.Label("6", Mid(d, p1, p2, p3), d.ShapeText("ShapeMark", Lgf.Ink, 30), 30);
        d.Scene(-3.3, 5.2, 5.6, -3.3);
    }

    // Two overlapping squares: the orange one on the yellow one's bottom right corner, its next corner
    // inside the yellow one and its far side through the yellow one's top right corner; the white
    // overlap is a triangle. Half the total area is yellow; what fraction is orange? Nothing to drag:
    // with the side through the corner, the half fixes the turn (the orange square's side leans 1 in 2)
    // and the sizes (the yellow square's side is √5/2 of the orange one's).
    [Puzzle("two-overlapping-squares-ii")]
    public static void TwoOverlappingSquaresII(Lgf d)
    {
        d.Texture = "FeltTip";
        const double b = 4;
        double a = b * Math.Sqrt(5) / 2, u = b / Math.Sqrt(5);
        var q = d.Pt(a, 0, "Q");
        var l = d.Pt(a - u, 2 * u, "L");
        var top = d.Pt(a + u, 3 * u, "U");
        var r = d.Pt(a + 2 * u, u, "R");
        var x = d.Pt(a, a, "X");
        var bottomLeft = d.Pt(0, 0);
        var topLeft = d.Pt(0, a);
        d.Polygon(d.Fill("Yellow", "#FFFFEC02", null), bottomLeft, q, l, x, topLeft);
        d.Polygon(d.Fill("Orange", "#FFFE9A10", null), q, r, top, x);
        string frame = d.Outline("Frame", Lgf.Ink, 4);
        d.Inked(frame, bottomLeft, q, x, topLeft);
        d.Inked(frame, q, r, top, l);
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(0, 3 * u, a + 2 * u, -0.6);
    }

    // Two congruent rectangles and their diagonals (dashed): the second one has a corner on the first
    // one's bottom right corner and the next one at its center, and its diagonal runs up the first
    // one's right side. The pink overlap is what fraction of the total area? Nothing to drag: the
    // corner at the center makes the half diagonal the short side, so the rectangles are √3 by 1.
    [Puzzle("two-rectangles-with-diagonals")]
    public static void TwoRectanglesWithDiagonals(Lgf d)
    {
        d.Texture = "FeltTip";
        const double h = 2, w = 2 * S3 * h;
        var bottomLeft = d.Pt(0, 0);
        var q = d.Pt(w, 0, "Q");
        var topRight = d.Pt(w, h, "Tr");
        var topLeft = d.Pt(0, h);
        var m = d.Pt(w / 2, h / 2, "M");
        var r = d.Pt(w + S3 * h, 3 * h / 2, "R");
        var s = d.Pt(w, 2 * h, "S");
        var p = d.Pt(w / 2 + h * h / (2 * w), h, "P");
        d.Polygon(d.Fill("Pink", "#FFD3015A", null), m, p, topRight, q);
        string dash = d.Line("Dash", "#FF0A0544", 3);
        // one diagonal of each rectangle is one line, from the bottom left corner through the center
        // and the top right corner to the far corner: dashed all the way, over the pink, as she drew it
        d.Dashed(dash, bottomLeft, r, 14);
        d.Dashed(dash, topLeft, m, 4);
        d.Dashed(dash, topRight, s, 4);
        string frame = d.Outline("Frame", Lgf.Ink, 4);
        d.Inked(frame, bottomLeft, q, topRight, topLeft);
        d.Inked(frame, q, r, s, m);
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(0, 2 * h, w + S3 * h, -0.5);
    }
}
