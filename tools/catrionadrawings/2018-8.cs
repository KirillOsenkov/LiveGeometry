#nullable disable

static partial class Puzzles
{
    // Plain hunting on five bells: five columns of dots, eleven rows, and from one row to the next each
    // bell swaps with the neighbour it didn't swap with last time. Every dot is colored by the bell
    // there, the paths of three bells are drawn (yellow, orange, red) and the first and last rows are
    // joined in gray. Nothing to drag: the rule decides every row.
    [Puzzle("bell-ringing-braids")]
    public static void BellRingingBraids(Lgf d)
    {
        d.Texture = "FeltTip";
        List<int[]> rows = [[0, 1, 2, 3, 4]];
        for (int row = 1; row <= 10; row++)
        {
            var next = (int[])rows[^1].Clone();
            for (int i = row % 2 == 1 ? 0 : 1; i + 1 < 5; i += 2)
            {
                (next[i], next[i + 1]) = (next[i + 1], next[i]);
            }

            rows.Add(next);
        }

        var grid = new V[11, 5];
        for (int row = 0; row <= 10; row++)
        {
            for (int column = 0; column < 5; column++)
            {
                grid[row, column] = d.Pt(column, -row);
            }
        }

        string gray = d.Line("Rounds", "#FFA3AEAA", 3);
        d.Segment(gray, grid[0, 0], grid[0, 4]);
        d.Segment(gray, grid[10, 0], grid[10, 4]);
        (string Color, string Dark)[] paths = [("#FFF5CE10", "#FFF5D52A"), ("#FFF7A46A", "#FFFFA870"), ("#FFF0512A", "#FFFF6040")];
        for (int bell = 0; bell < paths.Length; bell++)
        {
            string style = d.Line("Path" + bell, paths[bell].Color, 4, dark: paths[bell].Dark);
            for (int row = 0; row < 10; row++)
            {
                d.Segment(style, grid[row, Array.IndexOf(rows[row], bell)], grid[row + 1, Array.IndexOf(rows[row + 1], bell)]);
            }
        }

        // her dots, colored by bell: small filled circles, as big against the grid as hers
        (string Color, string Dark)[] bells = [("#FFF9CB0D", "#FFF7D23C"), ("#FFF88035", "#FFFF9A5C"), ("#FFF34612", "#FFFF6040"), ("#FF8A3A22", "#FFC07A5A"), (Lgf.Ink, Lgf.DarkInk)];
        var dots = bells.Select((bell, i) => d.Fill("Bell" + i, bell.Color, bell.Color, 2, darkOutline: bell.Dark)).ToArray();
        for (int row = 0; row <= 10; row++)
        {
            for (int column = 0; column < 5; column++)
            {
                d.Circle(dots[rows[row][column]], grid[row, column], d.Pt(column + 0.12, -row));
            }
        }

        d.Scene(0, 0, 4, -10);
    }

    // A big equilateral triangle and an equilateral triangle inside it with a corner on each side; on
    // the two pieces of the sides from the bottom corners, two more equilateral triangles of areas 20
    // and 5. The inner one's area? Nothing to drag: an inner equilateral triangle cuts every side the
    // same way, so the purple side is twice the orange (areas 20 and 5) and that fixes the figure.
    [Puzzle("four-equilateral-triangles")]
    public static void FourEquilateralTriangles(Lgf d)
    {
        d.Texture = null;
        var b = d.Pt(0, 0, "B");
        var c = d.Pt(6, 0, "C");
        var a = d.Pt(3, 6 * S3, "A");
        var p = d.Pt(2, 4 * S3, "P");
        var q = d.Pt(5, 2 * S3, "Q");
        var r = d.Pt(2, 0, "R");
        var left = d.Pt(-2, 4 * S3, "L");
        var right = d.Pt(7, 2 * S3, "K");
        d.Polygon(d.Outline("Brown", "#FF775F5F"), a, b, c);
        d.Polygon(d.Fill("Purple", "#FFFBF0FF", "#FFBF6FE6"), left, p, b);
        d.Polygon(d.Fill("Orange", "#FFFEF5E6", "#FFF0A020"), q, right, c);
        d.Polygon(d.Fill("Red", "#FFFBEBEB", "#FFD93838"), p, q, r);
        foreach (var point in new[] { a, b, c, p, q, r, left, right })
        {
            d.DotAt(point, size: 6);
        }

        string text = d.ShapeText();
        d.Label("20", Mid(d, left, p, b), text);
        d.Label("?", Mid(d, p, q, r), text);
        d.Label("5", Mid(d, q, right, c), text);
        d.Scene(-2, 6 * S3, 7, 0);
    }

    // An equilateral triangle on the top side of a square of area 18, a pink equilateral triangle inside
    // it with a corner on each side, and squares of areas 8 and 2 on the two pieces of the sides from the
    // bottom corners. The pink area? Nothing to drag: the areas fix the figure (sides 3√2, 2√2, √2).
    [Puzzle("three-squares-and-two-equilateral-triangles")]
    public static void ThreeSquaresAndTwoEquilateralTriangles(Lgf d)
    {
        d.Texture = null;
        var b = d.Pt(0, 0, "B");
        var c = d.Pt(3, 0, "C");
        var a = d.Pt(1.5, 3 * S3, "A");
        var p = d.Pt(1, 2 * S3, "P");
        var q = d.Pt(2.5, S3, "Q");
        var r = d.Pt(1, 0, "R");
        var leftOut = d.Pt(-2 * S3, 1);
        var leftTop = d.Pt(1 - 2 * S3, 2 * S3 + 1);
        var rightTop = d.Pt(2.5 + S3, S3 + 0.5);
        var rightOut = d.Pt(3 + S3, 0.5);
        var bottomLeft = d.Pt(0, -3);
        var bottomRight = d.Pt(3, -3);
        d.Polygon(d.Outline("RedFrame", "#FFD32F2F"), a, b, c);
        string blue = d.Fill("Blue", "#FFE8F1FA", "#FF1565C0");
        d.Polygon(blue, bottomLeft, bottomRight, c, b);
        d.Polygon(blue, b, p, leftTop, leftOut);
        d.Polygon(blue, c, rightOut, rightTop, q);
        d.Polygon(d.Fill("Pink", "#FFFBEBEB", "#FFD32F2F"), p, q, r);
        foreach (var point in new[] { a, b, c, p, q, r, leftOut, leftTop, rightTop, rightOut, bottomLeft, bottomRight })
        {
            d.DotAt(point, size: 6);
        }

        string text = d.ShapeText();
        d.Label("8", Mid(d, b, leftTop), text);
        d.Label("2", Mid(d, c, rightTop), text);
        d.Label("18", V.At(1.5, -1.5), text);
        d.Scene(-2 * S3, 2 * S3 + 1, 3 + S3, -3);
    }

    // A right triangle, its long leg cut in five equal parts; from each cut a line at right angles to the
    // leg down to the hypotenuse, and from each cut a line to the foot of the next one, which makes a
    // zig-zag of yellow right triangles and orange ones. What fraction is orange? Dragging the right
    // corner changes the triangle's shape: the fraction stays 2/5.
    [Puzzle("zig-zag-triangles")]
    public static void ZigZagTriangles(Lgf d)
    {
        double angle = 55 * Math.PI / 180;
        double ux = Math.Cos(angle), uy = Math.Sin(angle), nx = Math.Sin(angle), ny = -Math.Cos(angle);
        var o = d.Pt(0, 0, "O");
        var legs = Enumerable.Range(0, 6).Select(k => k == 0 ? o : d.Pt(k * ux, k * uy, k == 5 ? "T" : null)).ToArray();
        double tx = 5 * ux, ty = 5 * uy;
        var e = d.Slider(V.At(tx + 2.3 * nx, ty + 2.3 * ny), V.At(tx + 4.4 * nx, ty + 4.4 * ny), (3.5 - 2.3) / 2.1, (tx + 3.5 * nx, ty + 3.5 * ny), "E");
        var feet = Enumerable.Range(0, 6).Select(k => k == 0 ? o : k == 5 ? e : d.Pt((k / 5.0) * e)).ToArray();
        string yellow = d.Fill("Yellow", "#FFFBF28F", null);
        string orange = d.Fill("Orange", "#FFFAC49A", null);
        for (int k = 1; k <= 5; k++)
        {
            d.Polygon(yellow, legs[k - 1], legs[k], feet[k]);
        }

        for (int k = 1; k <= 4; k++)
        {
            d.Polygon(orange, legs[k], feet[k], feet[k + 1]);
        }

        string ink = d.Line("Ink");
        d.Segment(ink, o, legs[5]);
        d.Segment(ink, o, e);
        string pink = d.Line("Pink", "#FFDC1C5E", 3);
        for (int k = 1; k <= 4; k++)
        {
            d.Segment(pink, legs[k], feet[k]);
            d.Segment(pink, legs[k], feet[k + 1]);
        }

        d.Segment(pink, legs[5], e);
        // her ticks, just outside the leg
        for (int k = 0; k < 5; k++)
        {
            double mx = (k + 0.5) * ux, my = (k + 0.5) * uy;
            d.Segment(pink, d.Pt(mx + 0.04 * nx, my + 0.04 * ny), d.Pt(mx - 0.22 * nx, my - 0.22 * ny));
        }

        d.Scene(-0.1, 4.2, 6.1, -0.1);
    }

    // Two identical squares meeting at a corner, a circle centered there and a diameter of it; the
    // triangles from the corner to the diameter's ends and the squares' corners are 5 and 12, the pink
    // one to the circle on the other square's side? Dragging the tilted square turns it about the shared
    // corner: the pink triangle stays half the side times the radius, 13.
    [Puzzle("squares-meeting-in-a-circle")]
    public static void SquaresMeetingInACircle(Lgf d)
    {
        // areas 5 and 12 make the side times the radius 26 and the diameter's slope 5/12
        const double r = 3.2, a = 26 / r;
        double tilt = Math.Atan2(5, 12);
        var m = d.Pt(0, 0, "M");
        var left = d.Pt(-a, 0);
        var down = d.Pt(0, -a);
        var corner = d.Pt(-a, -a);
        var d1 = d.Pt(-r * Math.Cos(tilt), r * Math.Sin(tilt), "D1");
        var d2 = d.Pt(r * Math.Cos(tilt), -r * Math.Sin(tilt), "D2");
        var k = d.ArcSlider(m, a, -12, 16, 7.4, (0, 0), "K");
        var up = new V(V.Neg(k.Y), k.X);
        var top = d.Pt(up);
        var far = d.Pt(k + up);
        var apex = d.Pt(new V($"-{Lgf.N(r / a)} * K.Y", $"{Lgf.N(r / a)} * K.X"), "Ap");
        d.Polygon(d.Fill("Orange", "#FFFBC995", null), m, left, d1);
        d.Polygon(d.Fill("Blue", "#FFB5DDEB", null), m, down, d2);
        d.Polygon(d.Fill("Pink", "#FFF7A8C8", null), m, k, apex);
        d.Circle(d.Outline("Pencil", "#FFA8A8A8", 1.5), m, d1);
        d.Segment(d.Line("Diameter", "#FFA0407A", 4), d1, d2);
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        d.Inked(frame, left, m, down, corner);
        d.Inked(frame, m, k, far, top);
        d.Label("5", V.At(-3.6, 0.42), d.ShapeText("OrangeMark", "#FFE0602A", 26));
        d.Label("12", V.At(1.0, -2.9), d.ShapeText("BlueMark", "#FF1E9AB0", 26));
        d.Label("?", 0.3 * (k + apex), d.ShapeText("PinkMark", "#FFC02A6A", 28), 28);
        d.Scene(-a, 10, a, -a);
    }

    // Four squares on a line: two small ones (area 5) stacked, one twice as wide beside them, and a big
    // one beside that; the blue triangle joins the small ones' left middle corner, the top one's right
    // corner and the big one's top right corner. Its area? Dragging that corner resizes the big square:
    // it runs parallel to the triangle's first side, and the area stays twice a small square's, 10.
    [Puzzle("four-aligned-squares")]
    public static void FourAlignedSquares(Lgf d)
    {
        d.Texture = "FeltTip";
        var p = d.Pt(0, 1, "P");
        var q = d.Pt(1, 2, "Q");
        var r = d.Slider(V.At(5.5, 2.5), V.At(8.8, 5.8), (4.95 - 2.5) / 3.3, (7.95, 4.95), "R");
        string navy = "#FF2840B0";
        d.Polygon(d.Fill("Cyan", "#FF7ACFE4", null), p, q, r);
        string frame = d.Outline("Frame", Lgf.Ink, 3);
        d.Box(frame, V.At(0, 0), V.At(1, 1));
        d.Box(frame, V.At(0, 1), V.At(1, 2));
        d.Box(frame, V.At(1, 0), V.At(3, 2));
        d.Box(frame, V.At(3, 0), r);
        d.Polygon(d.Outline("Navy", navy, 3), p, q, r);
        d.Label("5", V.At(0.5, 0.5), d.Text("CyanMark", "#FF1A9EC0", 26));
        d.Scene(0, 5.3, 8.3, 0);
    }

    // Three squares with sides of consecutive integers: two side by side on a line, the biggest on the
    // smallest, flush with its right side; the arrow from the biggest one's top left corner through the
    // middle one's top right corner to the line is 4√10. Drawn as the numbers say (sides 5, 6, 7: the
    // middle one's corner halves the arrow), not as her sketch (about 3, 4, 5). Nothing to drag.
    [Puzzle("three-aligned-squares")]
    public static void ThreeAlignedSquares(Lgf d)
    {
        d.Texture = "FeltTip";
        d.Box(d.Outline("Yellow", "#FFF5D835", 4), V.At(4, 5), V.At(11, 12));
        d.Box(d.Outline("Cyan", "#FF14A8C4", 4), V.At(0, 0), V.At(6, 6));
        d.Box(d.Outline("Navy", "#FF1A3A9A", 4, dark: "#FF6E86F5"), V.At(6, 0), V.At(11, 5));
        var y = d.Pt(4, 12, "Y");
        var x = d.Pt(8, 0, "X");
        d.DoubleArrow(d.Line("Arrow", Lgf.Ink, 3), d.Text(), y, x, "4√10", V.At(7.1, 7.9));
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(0, 12, 11, -1.2);
    }

    // A trapezium with right angles on its right side (3 long) and its diagonals; the green triangle under
    // their crossing is 6 more than the yellow one over it. The sloping side? Dragging the top left corner
    // changes the top side: the bottom stays 4 longer, and the sloping side 5.
    [Puzzle("lengths-in-a-crossed-trapezium")]
    public static void LengthsInACrossedTrapezium(Lgf d)
    {
        d.Texture = "FeltTip";
        var br = d.Pt(0, 0, "BR");
        var tr = d.Pt(0, 3, "TR");
        var tl = d.Slider(V.At(-2, 3), V.At(-5.5, 3), 2 / 3.5, (-4, 3), "TL");
        var bl = d.Pt(new V("TL.X - 4", "0"), "BL");
        var x = d.Crossing(tl, br, bl, tr, "X");
        d.Polygon(d.Fill("Yellow", "#FFF2D82A", null), tl, tr, x);
        d.Polygon(d.Fill("Green", "#FF7CC040", null), bl, br, x);
        d.Polygon(d.Outline("Frame"), bl, br, tr, tl);
        string teal = d.Line("Teal", "#FF12876A", 2.5);
        d.Segment(teal, tl, br);
        d.Segment(teal, bl, tr);
        d.Label("3", V.At(0.5, 1.5));
        d.Label("x", new V("TL.X - 2.45", "1.95"));
        d.Scene(-8.2, 3, 0.8, 0);
    }

    // Four squares: the blue one of area 16, the purple one (20) on the line from its bottom left corner
    // to a point of its right side, the red one (17) on the line from a point of its left side to its
    // bottom right corner, and the orange one (17) on the line between those two points. The difference
    // between the red and purple triangles? Nothing to drag: the areas fix the points (heights 2 and 1).
    [Puzzle("overlapping-squares")]
    public static void OverlappingSquares(Lgf d)
    {
        d.Texture = null;
        var o = d.Pt(0, 0, "O");
        var b = d.Pt(4, 0, "B");
        var p = d.Pt(4, 2, "P");
        var q = d.Pt(0, 1, "Q");
        var x = d.Pt(4.0 / 3, 2.0 / 3, "X");
        var blueTop = new[] { d.Pt(4, 4), d.Pt(0, 4) };
        var purple = new[] { d.Pt(2, 6), d.Pt(-2, 4) };
        var orange = new[] { d.Pt(3, 6), d.Pt(-1, 5) };
        var red = new[] { d.Pt(5, 4), d.Pt(1, 5) };
        string purpleColor = "#FF6650D0", orangeColor = "#FFE0681C", redColor = "#FFD93733", blueColor = "#FF1B68B7";
        d.Polygon(d.Fill("PurpleShade", "#FFA5A3E4", null), o, q, x);
        d.Polygon(d.Fill("RedShade", "#FFDE8F94", null), x, b, p);
        d.Polygon(d.Outline("Purple", purpleColor), o, p, purple[0], purple[1]);
        d.Polygon(d.Outline("Orange", orangeColor), q, p, orange[0], orange[1]);
        d.Polygon(d.Outline("Red", redColor), q, b, red[0], red[1]);
        d.Polygon(d.Outline("Blue", blueColor), o, b, blueTop[0], blueTop[1]);
        foreach (var point in new[] { o, b, p, q, x }.Concat(blueTop).Concat(purple).Concat(orange).Concat(red))
        {
            d.DotAt(point, size: 6);
        }

        d.Label("20", V.At(-2.05, 2), d.Text("PurpleMark", purpleColor));
        d.Label("17", V.At(3.85, 5.25), d.Text("OrangeMark", orangeColor));
        d.Label("17", V.At(5.2, 1.95), d.Text("RedMark", redColor));
        d.Label("16", V.At(2, -0.6), d.Text("BlueMark", blueColor));
        d.Scene(-2.4, 6, 5.6, -0.9);
    }

    // A tilted square with a corner on the ground, one side through the top right corner of a square of
    // area 12 standing on the ground, the other through the top left corners of a stack of squares
    // (27, 12, 3) to the top one's. Its area? Nothing to drag. The picture alone leaves the tilt open;
    // drawn with the slope 2 her picture has, where the 12 sits flush right on the 27 and the left square
    // stands right under the tilted square's corner (area 135).
    [Puzzle("toppled-square")]
    public static void ToppledSquare(Lgf d)
    {
        d.Texture = "FeltTip";
        double r3 = Math.Sqrt(3);
        var g = d.Pt(0, 0, "G");
        var corner = d.Pt(3 * r3, 6 * r3, "C");
        var left = d.Pt(-6 * r3, 3 * r3, "L");
        var top = d.Pt(-3 * r3, 9 * r3, "T");
        d.Box(d.Fill("Grass", "#FF7CC23A", null), V.At(-6.3 * r3, -0.35), V.At(4.8 * r3, 0));
        string cyan = d.Outline("Cyan", "#FF1E9CB8", 3);
        d.Box(cyan, V.At(-6 * r3, 0), V.At(-4 * r3, 2 * r3));
        d.Box(cyan, V.At(1.5 * r3, 0), V.At(4.5 * r3, 3 * r3));
        d.Box(cyan, V.At(2.5 * r3, 3 * r3), V.At(4.5 * r3, 5 * r3));
        d.Box(cyan, V.At(3 * r3, 5 * r3), V.At(4 * r3, 6 * r3));
        d.Polygon(cyan, g, corner, top, left);
        string navy = d.Text("Navy", "#FF1C2A8C", 24);
        d.Label("12", V.At(-5 * r3, r3), navy, 24);
        d.Label("27", V.At(3 * r3, 1.5 * r3), navy, 24);
        d.Label("12", V.At(3.5 * r3, 4 * r3), navy, 24);
        d.Label("3", V.At(3.5 * r3, 5.5 * r3), navy, 24);
        d.Label("?", V.At(-1.5 * r3, 4.5 * r3), d.Text("NavyQuestion", "#FF1C2A8C", 36), 36);
        // a little paper below: the page's buttons sit over the bottom right corner
        d.Scene(-6.3 * r3, 9 * r3, 4.8 * r3, -1.3);
    }

    // Twelve evenly spaced points on a circle and two congruent rectangles across it at right angles: each
    // has its long sides tangent at two opposite points and its short sides on chords through points
    // 60° either side of the next ones. Red (in the rectangles outside the circle) or orange (in the
    // circle outside both)? Nothing to drag: the points fix it.
    [Puzzle("two-rectangles-overlapping-a-circle")]
    public static void TwoRectanglesOverlappingACircle(Lgf d)
    {
        d.Texture = null;
        const double R = 4;
        static (double X, double Y) Unit(double degrees) => (Math.Cos(degrees * Math.PI / 180), Math.Sin(degrees * Math.PI / 180));
        var o = d.Pt(0, 0, "O");
        var dots = Enumerable.Range(0, 12).Select(k => d.DotAt(V.At(R * Unit(20 + 30 * k).X, R * Unit(20 + 30 * k).Y), size: 9)).ToArray();
        V[] Rectangle(double along, double across)
        {
            var (ax, ay) = Unit(along);
            var (cx, cy) = Unit(across);
            return [.. new[] { (1, 1), (-1, 1), (-1, -1), (1, -1) }.Select(s => d.Pt(R * (s.Item1 * ax + s.Item2 * cx / 2), R * (s.Item1 * ay + s.Item2 * cy / 2)))];
        }

        var first = Rectangle(140, 50);
        var second = Rectangle(50, 140);
        string disk = d.Circle(null, o, dots[0]);
        string firstShape = d.Polygon(null, first);
        string secondShape = d.Polygon(null, second);
        string red = d.Fill("Red", "#FFE7242F", null);
        d.Region(red, "Difference", firstShape, disk);
        d.Region(red, "Difference", secondShape, disk);
        d.Region(d.Fill("Orange", "#FFFFC90D", null), "Difference", disk, firstShape, secondShape);
        d.Circle(d.Outline("Rim", "#FFC0C0C0", 2), o, dots[0]);
        string frame = d.Outline("Frame", "#FF4A4A4A", 2.5);
        d.Inked(frame, first);
        d.Inked(frame, second);
        // the corners farthest out: 4.35 across and up
        d.Scene(-4.35, 4.35, 4.35, -4.35);
    }

    // Six circles, each with twelve evenly spaced points, cut by lines between the points and the
    // center: a fan from one point, sectors, a triangle and two segments, a chevron, a triangle with
    // three bands, parallel stripes. What fraction of each is shaded? Nothing to drag.
    [Puzzle("subdivided-circles")]
    public static void SubdividedCircles(Lgf d)
    {
        d.Texture = null;
        (double X, double Y)[] centers = [(0, 0), (2.6, 0), (5.2, 0), (0, -2.65), (2.6, -2.65), (5.2, -2.65)];
        string rim = d.Outline("Rim", "#FF6E6E6E", 1.5);
        string ink = d.Line("Ink", Lgf.Ink, 3.5);
        var circles = centers.Select(center => SubdividedCircleParts.Make(d, center.X, center.Y)).ToArray();

        // a fan from the point at 240°, every other wedge shaded
        var fan = circles[0];
        string yellow = d.Fill("Yellow", "#FFFFF200", null);
        foreach (var (from, to) in new[] { (120, 90), (60, 30), (0, 330), (180, 150) })
        {
            d.Region(yellow, "Intersection", fan.Disk, d.Polygon(null, fan.Dot(240), fan.Far(240, from), fan.Far(240, to)));
        }

        d.Arc("CircleSegment", yellow, fan.Center, fan.Dot(240), fan.Dot(300));

        // sectors
        var sectors = circles[1];
        string pink = d.Fill("Pink", "#FFFFAEC9", null);
        foreach (var (from, to) in new[] { (60, 90), (330, 0), (210, 300), (150, 180) })
        {
            d.Arc("CircleSector", pink, sectors.Center, sectors.Dot(from), sectors.Dot(to));
        }

        // a segment, a right triangle on a diameter and a segment
        var triangle = circles[2];
        string blue = d.Fill("LightBlue", "#FF99D9EA", null);
        d.Arc("CircleSegment", blue, triangle.Center, triangle.Dot(150), triangle.Dot(210));
        d.Polygon(blue, triangle.Dot(120), triangle.Dot(240), triangle.Dot(300));
        d.Arc("CircleSegment", blue, triangle.Center, triangle.Dot(330), triangle.Dot(90));

        // two segments and a chevron to the center
        var chevron = circles[3];
        string green = d.Fill("Green", "#FF22B14C", null);
        d.Arc("CircleSegment", green, chevron.Center, chevron.Dot(120), chevron.Dot(180));
        d.Arc("CircleSegment", green, chevron.Center, chevron.Dot(0), chevron.Dot(60));
        d.Region(green, "Intersection", chevron.Disk, d.Polygon(null,
            chevron.Dot(90), chevron.Beyond(90, 210), chevron.Beyond(-1, 240), chevron.Center, chevron.Beyond(-1, 300), chevron.Beyond(90, 330)));

        // bands between the sides of a triangle and the chords parallel to them
        var bands = circles[4];
        string slate = d.Fill("Slate", "#FF7092BE", null);
        foreach (var (a, b, c) in new[] { (90, 210, 120), (210, 330, 240), (330, 90, 0) })
        {
            d.Region(slate, "Intersection", bands.Disk, bands.Strip(a, b, c));
        }

        // parallel stripes
        var stripes = circles[5];
        string orange = d.Fill("Orange", "#FFFF7F27", null);
        foreach (var (a, b, c) in new[] { (120, 150, 90), (60, 210, 30), (0, 270, 330) })
        {
            d.Region(orange, "Intersection", stripes.Disk, stripes.Strip(a, b, c));
        }

        foreach (var circle in circles)
        {
            d.Circle(rim, circle.Center, circle.Dot(0));
        }

        foreach (var (from, to) in new[] { (240, 120), (240, 90), (240, 60), (240, 30), (240, 0), (240, 330), (240, 300), (240, 180), (240, 150) })
        {
            d.Segment(ink, fan.Dot(from), fan.Dot(to));
        }

        foreach (int to in new[] { 90, 60, 0, 330, 300, 210, 180, 150 })
        {
            d.Segment(ink, sectors.Center, sectors.Dot(to));
        }

        foreach (var (from, to) in new[] { (150, 210), (120, 240), (240, 300), (120, 300), (90, 330) })
        {
            d.Segment(ink, triangle.Dot(from), triangle.Dot(to));
        }

        foreach (var (from, to) in new[] { (120, 180), (0, 60), (90, 210), (90, 330) })
        {
            d.Segment(ink, chevron.Dot(from), chevron.Dot(to));
        }

        d.Segment(ink, chevron.Dot(240), chevron.Center);
        d.Segment(ink, chevron.Center, chevron.Dot(300));
        foreach (var (from, to) in new[] { (90, 210), (210, 330), (330, 90), (120, 180), (0, 60), (240, 300) })
        {
            d.Segment(ink, bands.Dot(from), bands.Dot(to));
        }

        foreach (var (from, to) in new[] { (120, 150), (90, 180), (60, 210), (30, 240), (0, 270), (330, 300) })
        {
            d.Segment(ink, stripes.Dot(from), stripes.Dot(to));
        }

        d.Scene(-1, 1, 6.2, -3.65);
    }
}

/// <summary>A circle of subdivided-circles: its center, its points by angle and the hidden disk the shaded parts are cut from</summary>
file record SubdividedCircleParts(Lgf Drawing, double X, double Y, V Center, V[] Dots, string Disk)
{
    /// <summary>A circle of radius 1 with its twelve points, dots at every 30°</summary>
    public static SubdividedCircleParts Make(Lgf d, double x, double y)
    {
        var center = d.Pt(x, y);
        var dots = Enumerable.Range(0, 12).Select(k => d.DotAt(V.At(x + Math.Cos(k * Math.PI / 6), y + Math.Sin(k * Math.PI / 6)), size: 7)).ToArray();
        return new SubdividedCircleParts(d, x, y, center, dots, d.Circle(null, center, dots[0]));
    }

    public V Dot(int degrees) => Dots[(degrees % 360 + 360) % 360 / 30];

    static (double X, double Y) Unit(double degrees) => (Math.Cos(degrees * Math.PI / 180), Math.Sin(degrees * Math.PI / 180));

    (double X, double Y) At(int degrees) => degrees < 0 ? (X, Y) : (X + Unit(degrees).X, Y + Unit(degrees).Y);

    /// <summary>A hidden point 2.3 radii from the point at one angle toward the one at another: a corner of a wedge reaching past the circle</summary>
    public V Far(int from, int to)
    {
        var (ax, ay) = At(from);
        var (bx, by) = At(to);
        double length = Math.Sqrt((bx - ax) * (bx - ax) + (by - ay) * (by - ay));
        return Drawing.Pt(ax + 2.3 * (bx - ax) / length, ay + 2.3 * (by - ay) / length);
    }

    /// <summary>A hidden point 1.3 times as far from the point at one angle (-1: the center) as the one at another</summary>
    public V Beyond(int from, int to)
    {
        var (ax, ay) = At(from);
        var (bx, by) = At(to);
        return Drawing.Pt(ax + 1.3 * (bx - ax), ay + 1.3 * (by - ay));
    }

    /// <summary>A hidden strip between the line through the points at a and b and its parallel through the point at c, longer than the circle is wide</summary>
    public string Strip(int a, int b, int c)
    {
        var (ax, ay) = At(a);
        var (bx, by) = At(b);
        var (cx, cy) = At(c);
        double length = Math.Sqrt((bx - ax) * (bx - ax) + (by - ay) * (by - ay));
        double ux = 3 * (bx - ax) / length, uy = 3 * (by - ay) / length;
        return Drawing.Polygon(null, Drawing.Pt(ax - ux, ay - uy), Drawing.Pt(ax + ux, ay + uy), Drawing.Pt(cx + ux, cy + uy), Drawing.Pt(cx - ux, cy - uy));
    }
}
