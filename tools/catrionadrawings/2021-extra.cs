#nullable disable

static partial class Puzzles
{
    // Six equal squares in a semicircle (four on the diameter, two on top of the middle ones), and six
    // in a quarter circle (a staircase of three, two and one in its corner); which is larger, the pink
    // or the blue that is left over? With squares of side 1 the semicircle's radius is √5 (it runs
    // through the top row's outer corners and the bottom row's), the quarter circle's √10 (through
    // the outer corners of the bottom and the top rows). Nothing to drag: the squares fix both.
    [Puzzle("two-rectangles-and-a-square-ii")]
    public static void TwoRectanglesAndASquareII(Lgf d)
    {
        d.Texture = null;
        double r = Math.Sqrt(5), q = Math.Sqrt(10), x0 = 4;
        var o = d.Pt(0, 0, "O");
        var east = d.Pt(r, 0);
        var west = d.Pt(-r, 0);
        d.Arc("CircleSector", d.Fill("Pink", "#FFF27BA3", null), o, east, west);
        var c = d.Pt(x0, 0, "C");
        var right = d.Pt(x0 + q, 0);
        var top = d.Pt(x0, q);
        d.Arc("CircleSector", d.Fill("Blue", "#FF02BBD5", null), c, right, top);
        d.Arc("CircleArc", d.Line("PinkArc", "#FFC0306A", 2), o, east, west);
        d.Arc("CircleArc", d.Line("BlueArc", "#FF138E9C", 2), c, right, top);
        // over the sectors (a polygon is drawn under every sector)
        string paper = d.PaperFill();
        foreach (var (x, y) in new[] { (-2, 0), (-1, 0), (0, 0), (1, 0), (-1, 1), (0, 1) })
        {
            d.Inked(paper, d.Pt(x, y), d.Pt(x + 1, y), d.Pt(x + 1, y + 1), d.Pt(x, y + 1));
        }

        foreach (var (x, y) in new[] { (0, 0), (1, 0), (2, 0), (0, 1), (1, 1), (0, 2) })
        {
            d.Inked(paper, d.Pt(x0 + x, y), d.Pt(x0 + x + 1, y), d.Pt(x0 + x + 1, y + 1), d.Pt(x0 + x, y + 1));
        }

        d.Scene(-r, q, x0 + q, 0);
    }

    // Two squares side by side on a diameter of a circle, together 10 wide, the circle as far beyond
    // them on the left as on the right (the ticks), the big one's top right corner on the circle. So
    // the squares are centered on the circle's center and its radius squared is 25 plus the big
    // square's area. Her picture puts the line well below the center, a chord; the title says a
    // diameter, which this follows. Dragging the corner the squares share along the diameter trades
    // width between them: the circle grows and shrinks with the big square, the 10 and the ticks stay.
    [Puzzle("two-squares-on-a-diameter-of-a-circle")]
    public static void TwoSquaresOnADiameterOfACircle(Lgf d)
    {
        d.Texture = "FeltTip";
        const string navy = "#FF0A0C88";
        var o = d.Pt(0, 0, "O");
        // her split: the big square 6.3 wide, the small one 3.7
        var m = d.Slider(V.At(-0.4, 0), V.At(-2, 0), (1.3 - 0.4) / 1.6, (-1.3, 0), "M");
        string b = d.Num("5 - M.X", "b");
        string radius = d.Num($"sqrt(25 + {b} ^ 2)", "R");
        var west = d.Pt(new V($"-{radius}", "0"), "W");
        var east = d.Pt(new V(radius, "0"), "E");
        var left = d.Pt(-5, 0, "L");
        var right = d.Pt(5, 0, "Rt");
        d.Circle(d.Outline("Circle", navy, 3), o, east);
        string blue = d.Fill("Blue", "#FF0182FE", Lgf.Ink, 3.5);
        d.Box(blue, V.At(-5, 0), new V("M.X", "10 - b"));
        d.Box(blue, new V("M.X", "0"), new V("5", b));
        string ink = d.Line("Ink", Lgf.Ink, 3.5);
        d.Segment(ink, west, left, "TwoTicks");
        d.Segment(ink, left, right);
        d.Segment(ink, right, east, "TwoTicks");
        d.Dimension(d.Line("Arrow", navy, 2.5), d.Text("NavyMark", navy, 30), left, right, -1.35, "10", 1.8, 30);
        d.Scene(-8.1, 8.1, 8.1, -8.1);
    }
}
