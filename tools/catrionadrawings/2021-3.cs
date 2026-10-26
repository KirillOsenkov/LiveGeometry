#nullable disable

static partial class Puzzles
{
    // Two regular hexagons: a big yellow one with flat top and bottom, and a blue one standing on a corner
    // whose upper left corner is the yellow one's top right corner and whose bottom corner is the yellow
    // one's right corner (so its short diagonal is the yellow one's side); their overlap is green. The
    // angle at the shared top corner, between the yellow one's top and the blue one's side? Nothing to
    // drag: the hexagons fix it.
    [Puzzle("two-regular-hexagons-iii")]
    public static void TwoRegularHexagonsIII(Lgf d)
    {
        d.Texture = "FeltTip";
        const double s = 2, L = 2 * S3 * s;
        var yellow = HexagonFromLeft(d, V.At(-L, 0), Lgf.N(L));
        var v = yellow[2];
        var b = yellow[3];
        var c = new V(Lgf.N(L), Lgf.N(s));
        var blue = Enumerable.Range(0, 6).Select(i => d.Pt(c + V.At(s * Math.Cos((90 + 60 * i) * Math.PI / 180), s * Math.Sin((90 + 60 * i) * Math.PI / 180)))).ToArray();
        d.Polygon(d.Fill("Yellow", "#FFFEF862"), yellow);
        d.Polygon(d.Fill("Blue", "#FF3887FC"), blue);
        d.Polygon(d.Fill("Green", "#FF73A65B"), v, blue[2], b);
        string teal = "#FF2E7068";
        d.Angle(d.Line("Teal", teal), v, yellow[1], blue[0], 32);
        d.Label("?", v + V.At(-0.42, 0.9), d.Text("TealMark", teal, 32));
        d.Scene(-L, 4.1, L + S3 * s, -1.5 * s);
    }

    // A blue equilateral triangle and a pink isosceles one on the same base, the pink one's equal sides
    // meeting at the base's right end (so its apex is on the circle around that corner through the other
    // two); their overlap is purple. A purple line from the pink apex to the blue top: the angle at the
    // pink apex between it and the pink side? Dragging the pink apex along that circle: the angle stays.
    [Puzzle("isosceles-and-equilateral-triangles")]
    public static void IsoscelesAndEquilateralTriangles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 4;
        var a = d.Pt(0, 0, "A");
        var b = d.Pt(L, 0, "B");
        var t = d.Pt(L / 2, L * S3, "T");
        var p = d.ArcSlider(b, L, 140, 172, 157.2, (L, 0), "P");
        string pinkShape = d.Polygon(null, a, b, p);
        string blueShape = d.Polygon(null, a, b, t);
        d.Region(d.Fill("Pink", "#FFFED6E8", null), "Difference", pinkShape, blueShape);
        d.Region(d.Fill("Blue", "#FF70ACFE", null), "Difference", blueShape, pinkShape);
        d.Region(d.Fill("Purple", "#FFCDB4EE", null), "Intersection", pinkShape, blueShape);
        string frame = d.Outline("Frame");
        d.Polygon(frame, a, b, p);
        d.Polygon(frame, a, b, t);
        string purple = "#FF5B2DB5";
        string line = d.Line("PurpleLine", purple);
        d.Segment(line, p, t);
        d.Angle(d.Line("PurpleArc", purple, 3), p, t, a, 30);
        d.Label("?", p + V.At(0.48, 0.22), d.Text("PurpleMark", purple, 28));
        d.Scene(-0.3, L * S3, L, 0);
    }

    // Two concentric circles and a chord of the big one cut by the small one into 10, 5 and 10. The
    // shaded ring's area? The handle is the chord's middle, dragged toward and away from the center: the
    // circles grow and shrink with the chord's lengths kept, and the ring's area stays.
    [Puzzle("two-concentric-circles")]
    public static void TwoConcentricCircles(Lgf d)
    {
        double angle = 24 * Math.PI / 180;
        var u = V.At(Math.Cos(angle), Math.Sin(angle));
        var normal = V.At(-Math.Sin(angle), Math.Cos(angle));
        var o = d.Pt(0, 0, "O");
        var m = d.Slider(1.5 * normal, 8 * normal, (5.06 - 1.5) / 6.5, (-5.06 * Math.Sin(angle), 5.06 * Math.Cos(angle)), "M");
        var innerLeft = d.Pt(m - 2.5 * u, "I1");
        var innerRight = d.Pt(m + 2.5 * u, "I2");
        var outerLeft = d.Pt(m - 12.5 * u, "E1");
        var outerRight = d.Pt(m + 12.5 * u, "E2");
        string clear = d.Outline("Clear", "#00000000", 0, dark: "#00000000");
        string outer = d.Circle(clear, o, outerLeft);
        string inner = d.Circle(clear, o, innerLeft);
        d.Region(d.Fill("Orange", "#FFFEC690", null), "Difference", outer, inner);
        string frame = d.Outline("Frame");
        d.Circle(frame, o, outerLeft);
        d.Circle(frame, o, innerLeft);
        string magenta = "#FFC8287A";
        d.Segment(d.Line("Magenta", magenta), outerLeft, outerRight);
        string text = d.ShapeText("MagentaMark", magenta, 26);
        d.Label("10", Mid(d, outerLeft, innerLeft) + 1.15 * normal, text);
        d.Label("10", Mid(d, outerRight, innerRight) + 1.15 * normal, text);
        d.Label("5", m - 1.7 * normal, d.Text("MagentaFive", magenta, 32));
        d.Scene(-13.5, 13.5, 13.5, -13.5);
    }

    // Four squares in a quarter circle of radius 5: two big ones stacked up its vertical radius (so the
    // big ones are half the radius), two small ones stacked beside them, the top right corner of the small
    // stack on the arc (so the small ones are 0.6 of the big). Their total area? Nothing to drag: the
    // quarter circle fixes them.
    [Puzzle("four-squares-and-a-quarter-circle")]
    public static void FourSquaresAndAQuarterCircle(Lgf d)
    {
        const double R = 5, a = R / 2, b = 0.6 * a;
        var o = d.Pt(0, 0, "O");
        var top = d.Pt(0, R, "T");
        var east = d.Pt(R, 0, "E");
        var foot = d.Pt(a + b, 0, "F");
        string pink = d.Fill("Pink", "#FFE08FA3", Lgf.Ink, 3);
        d.Box(pink, V.At(0, 0), V.At(a, a));
        d.Box(pink, V.At(0, a), V.At(a, 2 * a));
        d.Box(pink, V.At(a, 0), V.At(a + b, b));
        d.Box(pink, V.At(a, b), V.At(a + b, 2 * b));
        string red = d.Line("Red", "#FFC44860", 3);
        d.Arc("CircleArc", red, o, east, top);
        d.Segment(red, foot, east);
        d.Scene(0, R, R, 0);
    }

    // A regular hexagon with a corner at a square's top left corner and the opposite corner on its bottom
    // side; a purple line from the square's bottom left corner through the hexagon's next corner to the
    // top side. The angle there? Dragging the hexagon's bottom corner along the bottom side resizes and
    // turns it: the square's corner stays on the hexagon's circumcircle (its long diagonal subtends a
    // right angle there), so the line keeps its direction and the angle stays.
    [Puzzle("hexagon-and-a-square")]
    public static void HexagonAndASquare(Lgf d)
    {
        d.Texture = "FeltTip";
        const double S = 4;
        var o = d.Pt(0, 0, "O");
        var a = d.Pt(0, S, "A");
        var topRight = d.Pt(S, S, "B");
        var c = d.Slider(V.At(0.4, 0), V.At(2.4, 0), (1.13 - 0.4) / 2, (1.13, 0), "C");
        var m = Mid(d, a, c);
        var hexagon = new[] { a, d.Pt(m + (a - m).Turn(-60), "V"), d.Pt(m + (a - m).Turn(-120)), c, d.Pt(m + (a - m).Turn(-240)), d.Pt(m + (a - m).Turn(-300)) };
        var x = d.Pt(S * Math.Tan(Math.PI / 6), S, "X");
        d.Polygon(d.Fill("Pink", "#FFFEBBD6", null), hexagon);
        d.Box(d.Outline("Square", "#FFF5459F", 3), V.At(0, 0), V.At(S, S));
        d.Polygon(d.Outline("Frame", Lgf.Ink, 3), hexagon);
        string purple = "#FF4B3A9B";
        string line = d.Line("Purple", purple, 3);
        d.Segment(line, o, x);
        d.Angle(line, x, topRight, o, 24);
        d.Label("?", x + V.At(0.45, -0.62), d.Text("PurpleMark", purple, 28));
        d.Scene(-1.6, S, S, 0);
    }

    // A rectangle with a red quarter circle in its bottom left corner and a blue one in its top right
    // corner, touching. Lines from the red one's top end and the blue one's left end to where they touch;
    // the sum of the two marked angles? Dragging the touching point along the line of the centers trades
    // the radii: each angle changes, the sum stays.
    [Puzzle("two-quarter-circles-in-a-rectangle")]
    public static void TwoQuarterCirclesInARectangle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double W = 4.49, H = 4;
        double length = Math.Sqrt(W * W + H * H);
        var u = V.At(W / length, H / length);
        var o = d.Pt(0, 0, "O");
        var c = d.Pt(W, H, "C");
        var t = d.Slider(2.1 * u, 3.9 * u, (3.52 - 2.1) / 1.8, (3.52 * W / length, 3.52 * H / length), "T");
        string r = d.Num("distance(O, T)", "r");
        string s = d.Num($"{Lgf.N(length)} - {r}", "s");
        var a = d.Pt(new V("0", r), "A");
        var east = d.Pt(new V(r, "0"));
        var b = d.Pt(new V($"{Lgf.N(W)} - {s}", Lgf.N(H)), "B");
        var bottom = d.Pt(new V(Lgf.N(W), $"{Lgf.N(H)} - {s}"));
        d.Arc("CircleSector", d.Fill("Red", "#FFF0767A"), o, east, a);
        d.Arc("CircleSector", d.Fill("Blue", "#FF6AA0EC"), c, b, bottom);
        d.Inked(d.Outline("Frame"), o, d.Pt(W, 0), c, d.Pt(0, H));
        string darkRed = "#FFB04448", darkBlue = "#FF2F4E9E";
        // on the fills, mid-tones in the dark theme: deep there too
        d.Segment(d.Line("RedLine", darkRed, 2.5, dark: "#FF6E1418"), a, t);
        d.Segment(d.Line("BlueLine", darkBlue, 2.5, dark: "#FF101F4E"), b, t);
        d.Angle(d.Wedge("RedWedge", darkRed), a, o, t, 44);
        d.Angle(d.Wedge("BlueWedge", darkBlue), b, c, t, 44);
        d.Scene(0, H, W, 0);
    }

    // Four squares in a row on a semicircle's diameter, two small then two big, the row's top left and
    // top right corners on the arc, 8 apart. The semicircle's area? With the small side a and the big b,
    // the radius squared is 5/16 of the distance squared whatever a and b are: dragging the small squares'
    // corner along the arc trades their sizes, the 8 and the semicircle stay.
    [Puzzle("four-squares-in-a-semi-circle")]
    public static void FourSquaresInASemiCircle(Lgf d)
    {
        d.Texture = "FeltTip";
        double R = Math.Sqrt(20);
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(R, 0, "E");
        var west = d.Pt(-R, 0, "W");
        var p = d.ArcSlider(o, R, 148, 168, 158.1, (0, 0), "P");
        string a = "P.Y";
        string b = d.Num("(-4 * P.X - 3 * P.Y) / 5", "b");
        string x1 = $"P.X + 2 * {a}";
        string x2 = $"P.X + 2 * {a} + {b}";
        var q = d.Pt(new V($"P.X + 2 * {a} + 2 * {b}", b), "Q");
        string green = "#FF128A50";
        string small = d.Fill("DarkGreen", "#FF7EC160", green, 3);
        string big = d.Fill("LightGreen", "#FFC7EEB3", green, 3);
        d.Box(small, new V("P.X", "0"), new V($"P.X + {a}", a));
        d.Box(small, new V($"P.X + {a}", "0"), new V(x1, a));
        d.Box(big, new V(x1, "0"), new V(x2, b));
        d.Box(big, new V(x2, "0"), q);
        d.Arc("CircleSegment", d.Outline("Frame", green, 3), o, east, west);
        var m = Mid(d, p, q);
        string arrow = d.Line("Arrow");
        d.Vector(arrow, m, p);
        d.Vector(arrow, m, q);
        d.Label("8", m + V.At(0, -0.62), d.ShapeText("ShapeMark", Lgf.Ink, 32));
        d.Scene(-R, R, R, 0);
    }

    // Two squares: one of side 12, and a tilted one with a corner at its bottom right corner and the next
    // corner on its left side. The shaded triangle from the first square's bottom left corner to the tilted
    // one's top corner and the first one's top right corner: its area? Dragging the tilted square's corner
    // up and down the left side turns and resizes it: the triangle changes shape, its area stays half the
    // square.
    [Puzzle("triangle-in-two-tilted-squares")]
    public static void TriangleInTwoTiltedSquares(Lgf d)
    {
        const double L = 12;
        var b = d.Pt(0, 0, "B");
        var corner = d.Pt(L, 0, "D");
        var c = d.Pt(L, L, "C");
        var top = d.Pt(0, L, "T");
        var e = d.Slider(V.At(0, 4), V.At(0, 9), (7.9 - 4) / 5, (0, 7.9), "E");
        var f = d.Pt(new V("E.Y", $"E.Y + {Lgf.N(L)}"), "F");
        var g = d.Pt(new V($"E.Y + {Lgf.N(L)}", Lgf.N(L)), "G");
        d.Polygon(d.Fill("Purple", "#FF8A7AC2", null), b, f, c);
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        d.Polygon(frame, b, corner, c, top);
        d.Polygon(frame, corner, e, f, g);
        string purple = "#FF5A4A9A";
        d.Dimension(d.Line("Arrow", purple), d.Text("PurpleMark", purple, 28), b, top, 2.2, "12", 1.6);
        d.Scene(-3, 20, 20, 0);
    }

    // A quarter circle and a circle through its center and the top of its vertical radius. Lines from that
    // top to where the circle crosses the bottom radius, to where it crosses the arc, and to the end of the
    // arc; shaded the triangle on the vertical radius and the thin wedge between the last two lines inside
    // the circle. The sum of the two marked angles? The line to the arc crossing meets the bottom line at
    // R^2 / (the crossing on the bottom), so the angles add up to 45 whatever the circle: dragging the bottom
    // crossing along the radius moves the circle (always through those two points), and the sum stays.
    [Puzzle("quarter-circle-overlapping-a-circle")]
    public static void QuarterCircleOverlappingACircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const double R = 4;
        var o = d.Pt(0, 0, "O");
        var t = d.Pt(0, R, "T");
        var e = d.Pt(R, 0, "E");
        var x = d.Slider(V.At(1.6, 0), V.At(3.6, 0), (2.62 - 1.6) / 2, (2.62, 0), "X");
        var k = d.Pt(new V("X.X / 2", Lgf.N(R / 2)), "K");
        string q = d.Num($"{Lgf.N(R * R)} / X.X", "q");
        string along = d.Num($"{Lgf.N(2 * R * R)} / ({q} ^ 2 + {Lgf.N(R * R)})", "u");
        var y = d.Pt(new V($"{along} * {q}", $"{Lgf.N(R)} - {Lgf.N(R)} * {along}"), "Y");
        string s = d.Num($"(X.X + {Lgf.N(R)}) / 8", "s");
        var z = d.Pt(new V($"{Lgf.N(R)} * {s}", $"{Lgf.N(R)} - {Lgf.N(R)} * {s}"), "Z");
        string fill = d.Fill("Orange", "#FFFDCE90", null);
        d.Polygon(fill, t, o, x);
        string disk = d.Circle(d.Outline("Clear", "#00000000", 0, dark: "#00000000"), k, o);
        string wedge = d.Polygon(null, t, d.Pt(t + 1.5 * (y - t)), d.Pt(t + 1.5 * (z - t)));
        d.Region(fill, "Intersection", wedge, disk);
        d.Circle(d.Line("Circle", "#FFF5A818", 3), k, o);
        d.Arc("CircleSector", d.Outline("Frame", Lgf.Ink, 3), o, e, t);
        string orange = "#FFFE8434";
        // on the orange fills, mid-tones in the dark theme: deep there too
        string line = d.Line("OrangeLine", orange, 2.5, dark: "#FF8A3200");
        d.Segment(line, t, x);
        d.Segment(line, t, y);
        d.Segment(line, t, e);
        // her marks are big, in proportion to the drawing: sectors of a radius in units
        string mark = d.Wedge("OrangeWedge", orange);
        SectorMark2021(d, mark, t, o, x, 0.7);
        SectorMark2021(d, mark, t, z, y, 1.0);
        d.Scene(-0.5, 4.3, 4.1, -0.3);
    }

    // Three small equilateral triangles in a row (up, down, up), a yellow one standing next to them, and a
    // big white one on the first small one's top and the yellow one's top, of area 14. The total area of
    // the other four? The big one's side squared is three small sides squared plus the yellow side
    // squared: dragging the corner where the small ones meet the yellow one along the bottom (the row keeps
    // its length) trades their sizes, and the big one stays the four together.
    [Puzzle("five-equilateral-triangles")]
    public static void FiveEquilateralTriangles(Lgf d)
    {
        d.Texture = "FeltTip";
        const double W = 8.12;
        var p1 = d.Pt(0, 0, "P1");
        var m = d.Slider(V.At(2.5, 0), V.At(5.5, 0), (4 - 2.5) / 3, (4, 0), "M");
        var e = d.Pt(W, 0, "E");
        string a = d.Num("M.X / 2", "a");
        string b = d.Num($"{Lgf.N(W)} - M.X", "b");
        var p2 = d.Pt(new V(a, "0"), "P2");
        var top1 = d.Pt(new V($"{a} / 2", $"{s3} * {a}"), "A");
        var top2 = d.Pt(new V($"1.5 * {a}", $"{s3} * {a}"), "C");
        var apex = d.Pt(new V($"M.X + {b} / 2", $"{s3} * {b}"), "B");
        var far = d.Pt(top1 + (apex - top1).Turn(60), "F");
        string ink = Lgf.Ink;
        d.Polygon(d.Fill("DarkGreen", "#FF8BBF7F", ink, 4), p1, p2, top1);
        d.Polygon(d.Fill("Green", "#FFA5D06C", ink, 4), top1, p2, top2);
        d.Polygon(d.Fill("LightGreen", "#FFCCE674", ink, 4), p2, m, top2);
        d.Polygon(d.Fill("Yellow", "#FFFEFD6A", ink, 4), m, e, apex);
        d.Polygon(d.Outline("Frame", ink, 4), top1, apex, far);
        d.Label("14", Mid(d, top1, apex, far), d.Text("Mark", ink, 40));
        d.Scene(0, 7.05, W, 0);
    }

    // A regular hexagon and four equilateral triangles with a corner at its center and the other two on
    // its sides: one is a sixth of the hexagon, the other three are turned. What fraction is shaded? Each
    // turned triangle's far corners are a turn of 60 degrees apart, so the triangle and the white sliver
    // beyond it make a sixth: dragging a corner of each turned triangle along its side turns it, and the
    // shaded part stays a third.
    [Puzzle("four-equilateral-triangles-inside-a-hexagon")]
    public static void FourEquilateralTrianglesInsideAHexagon(Lgf d)
    {
        d.Texture = "FeltTip";
        const double L = 4;
        var o = d.Pt(0, 0, "O");
        var hexagon = HexagonFromLeft(d, V.At(-L, 0), Lgf.N(L));
        var (left, topRight, bottomLeft) = (hexagon[0], hexagon[2], hexagon[5]);
        var at = (double t, (double X, double Y) a, (double X, double Y) b) => (a.X + t * (b.X - a.X), a.Y + t * (b.Y - a.Y));
        (double, double) l = (-L, 0), tl = (-L / 2, L * S3), tr = (L / 2, L * S3), r = (L, 0), br = (L / 2, -L * S3);
        var p1 = d.Slider(V.At(-L + 0.1 * L / 2, 0.1 * L * S3), V.At(-L + 0.9 * L / 2, 0.9 * L * S3), (0.45 - 0.1) / 0.8, at(0.45, l, tl), "P");
        var q1 = d.Slider(V.At(L / 2 + 0.05 * L / 2, (1 - 0.05) * L * S3), V.At(L / 2 + 0.35 * L / 2, (1 - 0.35) * L * S3), (0.17 - 0.05) / 0.3, at(0.17, tr, r), "Q");
        var s1 = d.Slider(V.At(L - 0.45 * L / 2, -0.45 * L * S3), V.At(L - 0.95 * L / 2, -0.95 * L * S3), (0.6 - 0.45) / 0.5, at(0.6, r, br), "S");
        var p2 = d.Pt(p1.Turn(-60), "P2");
        var q2 = d.Pt(q1.Turn(-60), "Q2");
        var s2 = d.Pt(s1.Turn(-60), "S2");
        string blue = d.Fill("Blue", "#FF5E9EFB", null);
        d.Polygon(blue, o, left, p1);
        d.Polygon(blue, o, p2, topRight, q1);
        d.Polygon(blue, o, q2, s1);
        d.Polygon(blue, o, s2, bottomLeft);
        d.Polygon(d.Outline("Hexagon", "#FF1A1AA8", 4), hexagon);
        string frame = d.Outline("Frame", Lgf.Ink, 4);
        d.Polygon(frame, o, left, bottomLeft);
        d.Polygon(frame, o, p1, p2);
        d.Polygon(frame, o, q1, q2);
        d.Polygon(frame, o, s1, s2);
        d.Scene(-L, L * S3, L, -L * S3);
    }

    // Two quarter circles of areas 4 and 2: the small one on a semicircle's diameter, its right end the
    // diameter's end, the big one hanging from the small one's top, its left end on the arc. The
    // semicircle's area? The areas fix the radii, and then the semicircle (its radius is the big one's):
    // nothing to drag.
    [Puzzle("two-quarter-circles-in-a-semi-circle-ii")]
    public static void TwoQuarterCirclesInASemiCircleII(Lgf d)
    {
        d.Texture = "FeltTip";
        const double r2 = 2;
        double r1 = r2 * Math.Sqrt(2), R = r1;
        var dCenter = d.Pt(0, 0, "D");
        var east = d.Pt(r2, 0, "E");
        var c = d.Pt(0, r2, "C");
        var left = d.Pt(-r1, r2, "A");
        var foot = d.Pt(0, r2 - r1, "F");
        var m = d.Pt(r2 - R, 0, "M");
        var west = d.Pt(r2 - 2 * R, 0, "W");
        d.Arc("CircleSector", d.Fill("Pink", "#FFFFB8C6", Lgf.Ink, 3), c, left, foot);
        d.Arc("CircleSector", d.Fill("DarkPink", "#FFE56C81", Lgf.Ink, 3), dCenter, east, c);
        string purple = "#FF8C2A78";
        d.Arc("CircleSegment", d.Outline("Semicircle", purple, 3), m, east, west);
        // the diameter runs across the pink quarter circle: drawn again, over it
        d.Segment(d.Line("Diameter", purple, 3), west, east);
        string text = d.ShapeText("ShapeMark", Lgf.Ink, 36);
        d.Label("4", c + V.At(-1.05, -0.85), text);
        d.Label("2", V.At(0.75, 0.85), text);
        d.Scene(r2 - 2 * R, R, r2, r2 - r1);
    }

    /// <summary>
    /// A filled angle mark as a sector of the given radius in drawing units (it scales with the drawing,
    /// as an AngleArc's pixel radius doesn't), counterclockwise from the side toward a to the side toward b
    /// </summary>
    private static string SectorMark2021(Lgf d, string style, V vertex, V a, V b, double radius)
    {
        static string Name(V point) => point.X.Substring(0, point.X.Length - 2);
        string toA = d.Num($"distance({Name(vertex)}, {Name(a)})");
        string toB = d.Num($"distance({Name(vertex)}, {Name(b)})");
        var start = d.Pt(vertex + Lgf.N(radius) * new V($"({a.X} - {vertex.X}) / {toA}", $"({a.Y} - {vertex.Y}) / {toA}"));
        var end = d.Pt(vertex + Lgf.N(radius) * new V($"({b.X} - {vertex.X}) / {toB}", $"({b.Y} - {vertex.Y}) / {toB}"));
        return d.Arc("CircleSector", style, vertex, start, end);
    }
}
