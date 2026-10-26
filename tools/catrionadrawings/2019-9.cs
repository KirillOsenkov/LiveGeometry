#nullable disable

static partial class Puzzles
{
    // A triangle, its altitude 4 (the blue arrow), and a circle in each of the two right triangles it
    // makes, of diameters 3 and 2. The triangle's area? Nothing to drag: the inradii fix the legs
    // beside the altitude (7.5 and 3).
    [Puzzle("two-circles-in-a-triangle")]
    public static void TwoCirclesInATriangle(Lgf d)
    {
        var a = d.Pt(-7.5, 0);
        var b = d.Pt(3, 0);
        var t = d.Pt(0, 4);
        d.Polygon(d.Outline("Frame", Lgf.Ink, 4.5), a, b, t);
        string blue = d.Outline("Circle", "#FF1414A8", 4);
        d.Circle(blue, d.Pt(-1.5, 1.5), d.Pt(0, 1.5));
        d.Circle(blue, d.Pt(1, 1), d.Pt(2, 1));
        string red = "#FFE0281C";
        string redLine = d.Line("Red", red, 2.5);
        string redText = d.Text("RedMark", red, 30);
        Batch2019s9.Gapped(d, redLine, redText, (-1.5, 0), (-1.5, 3), 0.5, 0.42, "3", 30);
        Batch2019s9.Gapped(d, redLine, redText, (1, 0), (1, 2), 0.5, 0.4, "2", 30);
        string cyan = "#FF009DF2";
        Batch2019s9.Gapped(d, d.Line("Cyan", cyan, 2.5), d.Text("CyanMark", cyan, 30), (0, 0), (0, 4), 0.72, 0.42, "4", 30);
        d.Scene(-7.5, 4, 3, 0);
    }

    // Four equilateral triangles meeting at a point in a rectangle: the green one on the bottom side
    // between its corners, the red one hanging from the top, the orange and the yellow ones on the
    // sides with their outer corners at the top corners. What fraction do they cover? Nothing to drag:
    // that fixes the rectangle (sides sqrt 3 : 2).
    [Puzzle("inverse-star-in-a-rectangle")]
    public static void InverseStarInARectangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double h = 3;
        double w = 2 * h / Math.Sqrt(3), top = 4 * h / 3, r = (top - h) / Math.Sqrt(3);
        var p = d.Pt(w / 2, h, "P");
        var bl = d.Pt(0, 0);
        var br = d.Pt(w, 0);
        var tr = d.Pt(w, top);
        var tl = d.Pt(0, top);
        d.Polygon(d.PaperFill("Paper", Lgf.Ink, 5), bl, br, tr, tl);
        d.Polygon(d.Fill("Green", "#FF01F1BF", Lgf.Ink, 5), bl, br, p);
        d.Polygon(d.Fill("Red", "#FFE9014D", Lgf.Ink, 5), d.Pt(w / 2 - r, top), d.Pt(w / 2 + r, top), p);
        d.Polygon(d.Fill("Orange", "#FFFE8502", Lgf.Ink, 5), tl, d.Pt(0, top - 2 * h / 3), p);
        d.Polygon(d.Fill("Yellow", "#FFFEDD01", Lgf.Ink, 5), tr, d.Pt(w, top - 2 * h / 3), p);
        d.Scene(0, top, w, 0);
    }

    // Three equal circles, each one's center on the other two. The part covered twice is shaded, the
    // middle (covered three times) left white; the radius is 4 (the arrow between two centers). The
    // shaded area? Nothing to drag: the circles fix it.
    [Puzzle("three-overlapping-circles")]
    public static void ThreeOverlappingCircles(Lgf d)
    {
        const double R = 4;
        double low = -R * Math.Sqrt(3) / 2;
        var a = d.Pt(-R / 2, 0, "A");
        var b = d.Pt(R / 2, 0, "B");
        var c = d.Pt(0, low, "C");
        string circleA = d.Circle(null, a, b);
        string circleB = d.Circle(null, b, a);
        string circleC = d.Circle(null, c, a);
        string ab = d.Region(null, "Intersection", circleA, circleB);
        string bc = d.Region(null, "Intersection", circleB, circleC);
        string ac = d.Region(null, "Intersection", circleA, circleC);
        string blue = d.Fill("Blue", "#FFBAE2FC", null);
        d.Region(blue, "Difference", ab, circleC);
        d.Region(blue, "Difference", bc, circleA);
        d.Region(blue, "Difference", ac, circleB);
        string navy = d.Outline("Circle", "#FF0A0AAA", 4.5);
        d.Circle(navy, a, d.Pt(-R / 2 - R, 0));
        d.Circle(navy, b, d.Pt(R / 2 + R, 0));
        d.Circle(navy, c, d.Pt(0, low - R));
        Batch2019s9.Gapped(d, d.Line("Arrow", Lgf.Ink, 2.5), d.Text("Mark", Lgf.Ink, 30), (-R / 2, 0), (R / 2, 0), 0.5, 0.55, "4", 30);
        d.Scene(-1.5 * R, R, 1.5 * R, low - R);
    }

    // Three 2 by 1 rectangles: two upright a square apart, the third tilted with its diagonal from the
    // left one's bottom right corner to the right one's top left corner. Its overlaps with the upright
    // ones are shaded: what fraction of the design? Nothing to drag: the diagonal (1 across, 2 up, the
    // length of the rectangle's own) fixes the tilt.
    [Puzzle("three-tilted-rectangles")]
    public static void ThreeTiltedRectangles(Lgf d)
    {
        d.Texture = "FeltTip";
        var a = d.Pt(1, 0, "A");
        var b = d.Pt(2.6, 1.2, "B");
        var c = d.Pt(2, 2, "C");
        var e = d.Pt(0.4, 0.8, "D");
        string red = d.Fill("Red", "#FFD80303", null);
        d.Polygon(red, e, a, d.Pt(1, 1.25));
        d.Polygon(red, c, b, d.Pt(2, 0.75));
        string purple = d.Outline("Frame", "#FF6A389A", 4.5);
        d.Box(purple, V.At(0, 0), V.At(1, 2));
        d.Box(purple, V.At(2, 0), V.At(3, 2));
        d.Polygon(purple, a, b, c, e);
        d.Scene(0, 2, 3, 0);
    }

    // Two regular hexagons of side 10, the green one the blue one moved right, so that the blue part,
    // the overlap and the green part have the same area: the shift is three quarters of a side. The
    // perimeter of the design? Nothing to drag: the equal areas fix the shift.
    [Puzzle("overlapping-hexagons")]
    public static void OverlappingHexagons(Lgf d)
    {
        const double s = 10;
        var blue = HexagonFromLeft(d, V.At(0, 0), Lgf.N(s));
        var green = HexagonFromLeft(d, V.At(0.75 * s, 0), Lgf.N(s));
        string blueShape = d.Polygon(null, blue);
        string greenShape = d.Polygon(null, green);
        d.Region(d.Fill("Blue", "#FF66C4E6", null), "Difference", blueShape, greenShape);
        d.Region(d.Fill("Green", "#FFE6F9AD", null), "Difference", greenShape, blueShape);
        d.Region(d.Fill("Aqua", "#FFA8F5DF", null), "Intersection", blueShape, greenShape);
        d.Polygon(d.Outline("GreenLine", "#FF1B9A1B", 3.5), green);
        d.Polygon(d.Outline("BlueLine", "#FF080680", 3.5), blue);
        d.Dimension(d.Line("Arrow", "#FF2A1A10", 3), d.Text("Mark", "#FF2A1A10", 30), blue[0], blue[1], 1.8, "10", 1.6);
        d.Scene(-2.6, S3 * s + 0.6, 1.75 * s + s, -S3 * s);
    }

    // A quarter circle of radius 13 and a square in its corner, two lines from the corner to where the
    // arc crosses the square cutting it into three parts of equal area. The square's area? Nothing to
    // drag: equal parts put the crossings two thirds along the far sides, which fixes the square
    // (side 3 sqrt 13).
    [Puzzle("square-overlapping-a-quarter-circle")]
    public static void SquareOverlappingAQuarterCircle(Lgf d)
    {
        const double R = 13;
        double a = Math.Sqrt(117), x = 2 * a / 3;
        var o = d.Pt(0, 0, "O");
        var top = d.Pt(x, a);
        var side = d.Pt(a, x);
        string purple = "#FF7F2FD8";
        d.Polygon(d.Fill("Blue", "#FF0085ED", purple), o, d.Pt(0, a), top);
        d.Polygon(d.Fill("Purple", "#FF7C13D1", purple), o, top, d.Pt(a, a), side);
        d.Polygon(d.Fill("Pink", "#FFD80CD5", purple), o, side, d.Pt(a, 0));
        var east = d.Pt(R, 0, "E");
        d.Arc("CircleSector", d.Outline("Frame", Lgf.Ink, 3.5), o, east, d.Pt(0, R));
        d.Dimension(d.Line("Arrow", Lgf.Ink, 2.5), d.Text("Mark", Lgf.Ink, 30), east, o, 1.2, "13", 1.3);
        d.Scene(0, R, R, -1.8);
    }

    // Two regular hexagons of area 6 sharing a side in a rectangle, the left one against its left and
    // top sides, the right one against its right and bottom sides. The shaded area? Nothing to drag:
    // the hexagons fix the rectangle.
    [Puzzle("two-hexagons-in-a-rectangle")]
    public static void TwoHexagonsInARectangle(Lgf d)
    {
        double s = Math.Sqrt(4 / Math.Sqrt(3)), h = S3 * s;
        d.Box(d.Fill("Orange", "#FFFFBD75", "#FFE01450", 3), V.At(0, -2 * h), V.At(3.5 * s, h));
        string hexagon = d.PaperFill("Hexagon", "#FF0B0B45", 4);
        var left = HexagonFromLeft(d, V.At(0, 0), Lgf.N(s));
        var right = HexagonFromLeft(d, V.At(1.5 * s, -h), Lgf.N(s));
        d.Polygon(hexagon, left);
        d.Polygon(hexagon, right);
        string red = d.Text("RedMark", "#FFC8102A", 36);
        d.Label("6", V.At(s, 0), red, 36);
        d.Label("6", V.At(2.5 * s, -h), red, 36);
        d.Scene(0, h, 3.5 * s, -2 * h);
    }

    // A circle, a chord 1 below its top and a semicircle hanging from the chord, 6 above the bottom of
    // the circle. What fraction of the circle is shaded? Nothing to drag: that fixes it (radius 5, the
    // semicircle's 3).
    [Puzzle("semi-circle-in-a-circle")]
    public static void SemiCircleInACircle(Lgf d)
    {
        const double R = 5, r = 3;
        var o = d.Pt(0, 0, "O");
        d.Circle(d.Fill("Yellow", "#FFFEFE1D", Lgf.Ink, 3), o, d.Pt(R, 0));
        var m = d.Pt(0, R - 1, "M");
        d.Arc("CircleSegment", d.PaperFill("Paper", Lgf.Ink, 3), m, d.Pt(-r, R - 1), d.Pt(r, R - 1));
        string green = "#FF20A000";
        // on the yellow, which stays bright in the dark theme: deep there too
        string arrow = d.Line("Green", green, 2.5, dark: "#FF0E5C00");
        var bottom = d.Pt(0, R - 1 - r);
        var low = d.Pt(0, -R);
        var middle = d.Pt(0, (R - 1 - r - R) / 2);
        d.Vector(arrow, middle, bottom);
        d.Vector(arrow, middle, low);
        d.Label("6", V.At(-0.5, (R - 1 - r - R) / 2), d.ShapeText("GreenMark", green, 30), 30);
        var high = d.Pt(0, R);
        var top = d.Pt(0, R - 0.5);
        d.Vector(arrow, top, m);
        d.Vector(arrow, top, high);
        d.Label("1", V.At(-0.5, R - 0.5), d.ShapeText("SmallMark", green, 20), 20);
        d.Scene(-R, R, R, -R);
    }

    // A rectangle; an equilateral triangle standing on its bottom side at the right corner, reaching the
    // top, and one on its point, whose sides run through the rectangle's left corners to a point on the
    // first one's side. What fraction is not covered? The left triangle's top corner is dragged along its
    // side: the triangle grows and shrinks, its part in the rectangle (and the fraction) stays.
    [Puzzle("two-triangles-overlapping-a-rectangle")]
    public static void TwoTrianglesOverlappingARectangle(Lgf d)
    {
        const double h = 3;
        double w = h * Math.Sqrt(3);
        var p = d.Pt(w / 2, h / 2, "P");
        var bl = d.Pt(0, 0);
        var br = d.Pt(w, 0);
        var tr = d.Pt(w, h);
        var tl = d.Pt(0, h);
        double u0 = 0.08, u1 = 1.3, home = 0.48;
        var t = d.Slider(V.At(-S3 * u0, h + u0 / 2), V.At(-S3 * u1, h + u1 / 2), (home - u0) / (u1 - u0), (-S3 * home, h + home / 2), "T");
        var b = d.Pt(new V("T.X", $"{Lgf.N(h)} - T.Y"), "Bt");
        string cyan = "#FF04D8FD";
        d.Polygon(d.Fill("Cyan", "#FFAEFBFB", null), bl, br, tr, tl);
        string paper = d.PaperFill("Paper", null);
        var apex = d.Pt(w - w / 3, h);
        var foot = d.Pt(w / 3, 0);
        d.Polygon(paper, t, b, p);
        d.Polygon(paper, foot, br, apex);
        d.Polygon(d.Outline("CyanLine", cyan, 4.5), bl, br, tr, tl);
        string frame = d.Outline("Frame", Lgf.Ink, 4.5);
        d.Polygon(frame, t, b, p);
        d.Polygon(frame, foot, br, apex);
        d.Scene(-0.5, h + 0.25, w, -0.25);
    }

    // A regular octagon and a square with two opposite corners on its left and bottom sides; the
    // square's other two sides drawn on (dashed) to the top and the right sides, and the line between
    // where they end. The angle? Both corners on the octagon are dragged along their sides: the square
    // changes (its inner corner stays on the octagon's diagonal), the angle stays 45 degrees.
    [Puzzle("square-on-an-octagon")]
    public static void SquareOnAnOctagon(Lgf d)
    {
        const double a = 4, range = 0.35 * a;
        double radius = a / Math.Cos(Math.PI / 8);
        var octagon = Enumerable.Range(0, 8).Select(i => d.Pt(radius * Math.Cos((22.5 + 45 * i) * Math.PI / 180), radius * Math.Sin((22.5 + 45 * i) * Math.PI / 180))).ToArray();
        var corner = d.Slider(V.At(-a, -range), V.At(-a, range), (0.2 + range) / (2 * range), (-a, 0.2), "A");
        var other = d.Slider(V.At(-range, -a), V.At(range, -a), (-1.04 + range) / (2 * range), (-1.04, -a), "C");
        var k = d.Pt(new V("(A.Y + C.X) / 2", "(A.Y + C.X) / 2"), "K");
        var last = d.Pt(corner + other - k, "L");
        var x = d.Crossing(other, k, d.Pt(0, a), d.Pt(1, a), "X");
        var y = d.Crossing(corner, k, d.Pt(a, 0), d.Pt(a, 1), "Y");
        string red = "#FFC40404";
        d.Polygon(d.Outline("Octagon", red, 5), octagon);
        d.Polygon(d.Outline("Square", Lgf.Ink, 5), corner, k, other, last);
        string dashes = d.Line("Dashes", Lgf.Ink, 3.5);
        d.Dashed(dashes, k, x, 9);
        d.Dashed(dashes, k, y, 8);
        string orange = "#FFDD2A0A";
        d.Dashed(d.Line("OrangeDashes", orange, 3), x, y, 10);
        d.ArcMark(d.Line("AngleLine", orange, 3), x, k, y, 0.5);
        d.Label("?", x + V.At(0.22, -0.95), d.Text("OrangeMark", orange, 26));
        d.Scene(-4.8, a, a, -a);
    }

    // An equilateral triangle, a semicircle on the middle of its base touching both sides, and a circle
    // of area 6 touching both sides and the semicircle. The semicircle's area? Nothing to drag: the
    // triangle fixes both (the semicircle's radius is three times the circle's).
    [Puzzle("circle-and-semi-circle-in-triangle")]
    public static void CircleAndSemiCircleInTriangle(Lgf d)
    {
        const double H = 6, R = H / 2, r = H / 6;
        double half = H / Math.Sqrt(3);
        var o = d.Pt(0, 0, "O");
        d.Polygon(d.Outline("Frame", "#FF5A2AB0", 3.5), d.Pt(-half, 0), d.Pt(half, 0), d.Pt(0, H));
        d.Arc("CircleSegment", d.Fill("Blue", "#FFE3F8FB", Lgf.Ink, 3.5), o, d.Pt(R, 0), d.Pt(-R, 0));
        var c = d.Pt(0, R + r, "C");
        d.Circle(d.Fill("Yellow", "#FFF0FBA8", Lgf.Ink, 3.5), c, d.Pt(r, R + r));
        d.Label("6", c, d.ShapeText("GreenMark", "#FF78C010", 40), 40);
        d.Label("?", V.At(0, R / 2), d.ShapeText("BlueMark", "#FF079FD2", 40), 40);
        d.Scene(-half, H, half, 0);
    }

    // Two 3-4-5 triangles: the blue one with its right angle at the left end of the green one's
    // hypotenuse, its legs up and along it; the green one's right angle on top. The overlap? Nothing
    // to drag: the triangles fix it.
    [Puzzle("overlapping-triangles")]
    public static void OverlappingTriangles(Lgf d)
    {
        d.Texture = "FeltTip";
        var o = d.Pt(0, 0, "O");
        var e = d.Pt(4, 0, "E");
        var t = d.Pt(0, 3, "T");
        var f = d.Pt(5, 0, "F");
        var g = d.Pt(3.2, 2.4, "G");
        d.Polygon(d.Fill("Lime", "#FFB4F23E", null), o, e, d.Pt(2, 1.5));
        d.Polygon(d.Outline("Green", "#FF04A446", 5), o, f, g);
        d.Polygon(d.Outline("Blue", "#FF0550DC", 5), o, e, t);
        d.Scene(0, 3, 5, 0);
    }

    // Two congruent rectangles: an upright one 8 wide, a tilted one standing on its bottom right corner,
    // its side through the upright one's top right corner, which is 3 from its top corner; together
    // they are 16 across. The strip of the upright one left of the tilted one is shaded: its area?
    // Nothing to drag: the 16 makes the rectangles 8 wide, the 3 makes them 73/6 tall.
    [Puzzle("tilted-rectangles")]
    public static void TiltedRectangles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double w = 8, h = 73.0 / 6;
        double cos = w / h, sin = Math.Sqrt(1 - cos * cos);
        var b = d.Pt(0, 0, "B");
        var l = d.Pt(-w * sin, w * cos, "L");
        var r = d.Pt(h * cos, h * sin, "R");
        var t = d.Pt(-w * sin + h * cos, w * cos + h * sin, "T");
        d.Box(d.Fill("Yellow", "#FFFDE902", "#FFF5DC00", 3), V.At(-w, 0), V.At(-w * sin, h));
        string frame = d.Outline("Frame", Lgf.Ink, 4.5);
        d.Box(frame, V.At(-w, 0), V.At(0, h));
        d.Polygon(frame, b, r, t, l);
        string cyan = "#FF00A0C3";
        string arrow = d.Line("Arrow", cyan, 2.5);
        string text = d.Text("CyanMark", cyan, 30);
        // the 3 runs from the upright one's corner to the tilted one's top, beside its side
        double nx = -sin, ny = cos, tx = -w * sin + h * cos, ty = w * cos + h * sin;
        Batch2019s9.Beside(d, arrow, text, (0.8 * nx, h + 0.8 * ny), (tx + 0.8 * nx, ty + 0.8 * ny), "3", (-0.9, h + 2.4), 30);
        Batch2019s9.Beside(d, arrow, text, (-w, -1.2), (w, -1.2), "16", (0, -2.3), 30);
        d.Scene(-w, w * cos + h * sin, w, -2.9);
    }
}

file static class Batch2019s9
{
    /// <summary>A double arrow from a to b, the text in a gap around the point the fraction t of the way</summary>
    public static void Gapped(Lgf d, string line, string textStyle, (double X, double Y) a, (double X, double Y) b, double t, double gap, string text, double size)
    {
        double length = Math.Sqrt((b.X - a.X) * (b.X - a.X) + (b.Y - a.Y) * (b.Y - a.Y));
        double ux = (b.X - a.X) / length, uy = (b.Y - a.Y) / length;
        double mx = a.X + t * (b.X - a.X), my = a.Y + t * (b.Y - a.Y);
        d.Vector(line, d.Pt(mx - gap * ux, my - gap * uy), d.Pt(a.X, a.Y));
        d.Vector(line, d.Pt(mx + gap * ux, my + gap * uy), d.Pt(b.X, b.Y));
        d.Label(text, V.At(mx, my), textStyle, size);
    }

    /// <summary>An unbroken double arrow from a to b, the text beside it</summary>
    public static void Beside(Lgf d, string line, string textStyle, (double X, double Y) a, (double X, double Y) b, string text, (double X, double Y) at, double size)
    {
        var middle = d.Pt((a.X + b.X) / 2, (a.Y + b.Y) / 2);
        d.Vector(line, middle, d.Pt(a.X, a.Y));
        d.Vector(line, middle, d.Pt(b.X, b.Y));
        d.Label(text, V.At(at.X, at.Y), textStyle, size);
    }
}
