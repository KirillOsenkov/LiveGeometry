#:property Nullable=disable
#:property PublishAot=false

// hat - writes the three gallery drawings of the hat, the aperiodic monotile of Smith, Myers,
// Kaplan and Goodman-Strauss (2023), into a folder:
//
// - AperiodicMonotile.lgf: a hat in the middle and the copies of it around it that a tiling
//   puts there, the flipped ones in their own color. Two free points A and B are the ends of
//   the bottom edge of the middle hat; every corner is A plus i halves of AB and j halves of
//   AC (C makes an equilateral triangle on AB), so dragging A or B turns and resizes it all.
// - HatKites.lgf: a gray grid of hexagons, each cut into six kites, and one hat on it, its
//   eight kites colored by the hexagon they come from. Dragging the hat moves a hidden free
//   point H; the hat sits on the hexagon center nearest to H (Z) and is turned by the
//   multiple of 60 degrees nearest to the angle of a point R on a circle around Z (Q holds
//   that turn's cosine and sine), so it always lies on the grid.
// - HatToSpectre.lgf: a smaller patch whose sides change length as a point slides along a
//   track: a side of the hat is either a short side of a kite (along a multiple of 60
//   degrees) or a long one (30 degrees off), every corner is a sum of such sides, and with
//   the short ones a long and the long ones b long the patch is the tiling by Tile(a, b) of
//   the paper. a = sin(90 t), b = cos(90 t) for the place t of the point on the track: the
//   chevron at 0, the hat at 1/3, the spectre at 1/2, the turtle at 2/3, the comet at 1.
//
// The patch is found by a search on the kite grid: copies, turned by a multiple of 60
// degrees and flipped or not, are laid one by one onto the uncovered kite nearest the
// middle, without overlaps, until everything up to a radius is covered; only the copies
// well inside are kept, so that the ones shown are surrounded in the search too.
//
//   dotnet tools/hat.cs -- <folder> [seed]

using System.Globalization;
using System.Text;

if (args.Length < 1)
{
    Console.WriteLine("usage: hat <folder> [seed]");
    return 1;
}

var invariant = CultureInfo.InvariantCulture;
double root3 = Math.Sqrt(3);

// lattice coordinates: (i, j) is i * e1 + j * e2, e1 = (1, 0), e2 = (1/2, sqrt(3)/2); the
// hat as Kaplan's hatviz has it, with a hexagon's center at (0, 0). Hexagons have sides 2,
// their centers are the lattice of (2, 2) and (-2, 4)
var hat = new (int I, int J)[]
{
    (0, 0), (-1, -1), (0, -2), (2, -2), (2, -1), (4, -2), (5, -1),
    (4, 0), (3, 0), (2, 2), (0, 3), (0, 2), (-1, 2)
};

const double SearchRadius = 19;
const double ShowRadius = 10;
const double MorphRadius = 6.5;

// the order in which the search tries the copies; 2 gives a round patch with 3 flipped
int seed = args.Length > 1 ? int.Parse(args[1], invariant) : 2;

// the 12 placements about a hexagon's center: turns by 60 degrees, then flipped or not
var orientations = new List<(int I, int J)[]>();
for (int flip = 0; flip < 2; flip++)
{
    for (int turn = 0; turn < 6; turn++)
    {
        orientations.Add(hat.Select(p => Orient(p, turn, flip == 1)).ToArray());
    }
}

// the cells: the 30-60-90 triangles of the lattice cut by its medians; a hat covers 192
var cellsOf = orientations.Select(Cells).ToArray();
if (cellsOf[0].Count != 192)
{
    Console.WriteLine("the hat is not 8 kites: " + cellsOf[0].Count + " cells");
    return 1;
}

var main = new Placement(0, (0, 0));
var occupied = new Dictionary<(int, int), int>();
var placements = new List<Placement>();
Place(main);
var center = Centroid(main);
var targets = new List<(int, int)>();
for (int i = -40; i <= 40; i++)
{
    for (int j = -40; j <= 40; j++)
    {
        foreach (var cell in TriangleCells(i, j))
        {
            if (Distance(CellPoint(cell), center) < SearchRadius)
            {
                targets.Add(cell);
            }
        }
    }
}

targets = targets.OrderBy(cell => Distance(CellPoint(cell), center)).ToList();
var random = new Random(seed);
long steps = 0;
if (!Search())
{
    Console.WriteLine("no patch found");
    return 1;
}

Console.WriteLine($"{placements.Count} hats in the search, {steps} steps");
var shown = placements.Where(placement => Distance(Centroid(placement), center) < ShowRadius).ToList();

// neighbors: copies that share a piece of an edge
var neighbors = shown.ToDictionary(placement => placement, placement => new HashSet<Placement>());
foreach (var placement in shown)
{
    foreach (var other in shown)
    {
        if (placement != other && Touch(placement, other))
        {
            neighbors[placement].Add(other);
        }
    }
}

// a flipped copy never touches another one in a hat tiling
foreach (var placement in shown.Where(IsFlipped))
{
    if (neighbors[placement].Any(IsFlipped))
    {
        Console.WriteLine("two flipped hats touch: try another seed");
        return 1;
    }
}

// colors: the middle one, the flipped ones, and three more for the rest, no two neighbors
// alike
var others = new[] { "Hat1", "Hat2", "Hat3" };
var colorOf = new Dictionary<Placement, string> { [main] = "HatMain" };
foreach (var placement in shown.Where(IsFlipped))
{
    colorOf[placement] = "HatFlipped";
}

var rest = shown.Where(placement => !colorOf.ContainsKey(placement)).OrderBy(placement => Distance(Centroid(placement), center)).ToList();
if (!Color(0))
{
    Console.WriteLine("no coloring");
    return 1;
}

Console.WriteLine($"{shown.Count} hats shown, {shown.Count(IsFlipped)} flipped");

// the kites of the middle hat: a hexagon around each center is cut into six, kite k being
// the center, the middle of side k, the corner after it and the middle of the next side
var mainOutline = Vertices(main).Select(Cartesian).ToArray();
var mainKites = new List<((int I, int J) Hexagon, (int I, int J)[] Corners)>();
foreach (var hexagonCenter in HexagonCenters(4))
{
    foreach (var kite in KitesOf(hexagonCenter))
    {
        var points = kite.Select(Cartesian).ToArray();
        if (Inside(mainOutline, (points.Average(p => p.X), points.Average(p => p.Y))))
        {
            mainKites.Add((hexagonCenter, kite));
        }
    }
}

var counts = mainKites.GroupBy(kite => kite.Hexagon).Select(group => group.Count()).OrderByDescending(count => count).ToList();
Console.WriteLine("kites of the hat by hexagon: " + string.Join(" + ", counts));
if (mainKites.Count != 8)
{
    Console.WriteLine("the middle hat is not 8 kites");
    return 1;
}

Directory.CreateDirectory(args[0]);
WriteTiling(Path.Combine(args[0], "AperiodicMonotile.lgf"));
WriteKites(Path.Combine(args[0], "HatKites.lgf"));
if (!WriteMorph(Path.Combine(args[0], "HatToSpectre.lgf")))
{
    return 1;
}

return 0;

void WriteTiling(string path)
{
    // a grid unit is half of AB on screen: the patch's middle at (0, 0), AB 1 long
    const double Unit = 0.5;
    var origin = (I: 0, J: -2);
    var points = new StringBuilder();
    var figures = new StringBuilder();
    var names = new Dictionary<(int, int), string>
    {
        [(0, -2)] = "A",
        [(2, -2)] = "B",
        [(0, 0)] = "C",
    };

    string Corner((int I, int J) corner)
    {
        if (names.TryGetValue(corner, out var name))
        {
            return name;
        }

        int i = corner.I - origin.I;
        int j = corner.J - origin.J;
        name = "P" + Number(i) + "_" + Number(j);
        names[corner] = name;
        points.AppendLine(string.Format(invariant,
            "    <PointByCoordinates Name=\"{0}\" Visible=\"false\" X=\"A.X + ({1} * (B.X - A.X) + {2} * (C.X - A.X)) / 2\" Y=\"A.Y + ({1} * (B.Y - A.Y) + {2} * (C.Y - A.Y)) / 2\">",
            name, i, j));
        points.AppendLine("      <Dependency Name=\"A\" />");
        points.AppendLine("      <Dependency Name=\"B\" />");
        points.AppendLine("      <Dependency Name=\"C\" />");
        points.AppendLine("    </PointByCoordinates>");
        return name;
    }

    // the middle one last, so that its outline is on top
    int index = 0;
    foreach (var placement in shown.Where(placement => placement != main).Append(main))
    {
        index++;
        string name = placement == main ? "Hat" : "Hat" + (index + 1);
        figures.AppendLine("    <Polygon Name=\"" + name + "\" Style=\"" + colorOf[placement] + "\">");
        foreach (var corner in Vertices(placement))
        {
            figures.AppendLine("      <Dependency Name=\"" + Corner(corner) + "\" />");
        }

        figures.AppendLine("    </Polygon>");
    }

    // the lines between two kites of the middle hat
    index = 0;
    foreach (var edge in InnerKiteEdges())
    {
        index++;
        figures.AppendLine("    <Segment Name=\"Kite" + index + "\" Style=\"KiteLine\">");
        figures.AppendLine("      <Dependency Name=\"" + Corner(edge.Item1) + "\" />");
        figures.AppendLine("      <Dependency Name=\"" + Corner(edge.Item2) + "\" />");
        figures.AppendLine("    </Segment>");
    }

    var pointA = Cartesian(origin);
    var text = Start();
    TileStyles(text);
    text.AppendLine("    <LineStyle Name=\"KiteLine\" Color=\"#B0FFFFFF\" StrokeWidth=\"1\">");
    text.AppendLine("      <Dark Color=\"#80FFFFFF\" />");
    text.AppendLine("    </LineStyle>");
    text.AppendLine("  </Styles>");
    text.AppendLine("  <Figures>");
    text.AppendLine(string.Format(invariant, "    <FreePoint Name=\"A\" X=\"{0:0.####}\" Y=\"{1:0.####}\" />", (pointA.X - center.X) * Unit, (pointA.Y - center.Y) * Unit));
    text.AppendLine(string.Format(invariant, "    <FreePoint Name=\"B\" X=\"{0:0.####}\" Y=\"{1:0.####}\" />", (pointA.X + 2 - center.X) * Unit, (pointA.Y - center.Y) * Unit));
    text.AppendLine("    <PointByCoordinates Name=\"C\" Visible=\"false\" X=\"A.X + (B.X - A.X) / 2 - sqrt(3) / 2 * (B.Y - A.Y)\" Y=\"A.Y + (B.Y - A.Y) / 2 + sqrt(3) / 2 * (B.X - A.X)\">");
    text.AppendLine("      <Dependency Name=\"A\" />");
    text.AppendLine("      <Dependency Name=\"B\" />");
    text.AppendLine("    </PointByCoordinates>");
    text.Append(points);
    text.Append(figures);
    Finish(text, path, "Aperiodic Monotile",
        "In 2023 David Smith, a retired print technician from England who loves shapes, found the hat. Its copies cover the whole plane with no gaps and no overlaps, yet the pattern never repeats: slide it any way you like, and it never lands on itself. A single tile like that is called an aperiodic monotile. Mathematicians had looked for one for more than fifty years.\\n\\nSome copies are flipped over, like the purple ones: roughly one in eight. The lines in the orange one show the eight kites it is made of.\\n\\nDrag the yellow points to turn and resize the tiles.");
}

void WriteKites(string path)
{
    const double Unit = 0.5;
    var text = Start();
    TileStyles(text);
    text.AppendLine("    <ShapeStyle Name=\"Hexagon\" Fill=\"#00FFFFFF\" Color=\"#FF9BA2AC\" StrokeWidth=\"1.5\">");
    text.AppendLine("      <Dark Color=\"#FF6B727C\" />");
    text.AppendLine("    </ShapeStyle>");
    text.AppendLine("    <ShapeStyle Name=\"GridKite\" Fill=\"#00FFFFFF\" Color=\"#FFCDD1D7\">");
    text.AppendLine("      <Dark Color=\"#FF4A4F57\" />");
    text.AppendLine("    </ShapeStyle>");
    text.AppendLine("    <LineStyle Name=\"TurnCircle\" Color=\"#FFB8BEC6\" Dash=\"Dash\">");
    text.AppendLine("      <Dark Color=\"#FF5E6570\" />");
    text.AppendLine("    </LineStyle>");
    Gradient(text, "Kite1", "#FFFFA375", "#FFFF7448", "#FFFFD27A", "#FFC0530A", outline: "#C0FFFFFF", darkOutline: "#90FFFFFF", width: 1);
    Gradient(text, "Kite2", "#FFFFE89A", "#FFF5BE3C", "#FFFFF2A8", "#FF9C7A00", outline: "#C0FFFFFF", darkOutline: "#90FFFFFF", width: 1);
    Gradient(text, "Kite3", "#FFFFB8C0", "#FFF07A8C", "#FFFFB8D0", "#FFA02858", outline: "#C0FFFFFF", darkOutline: "#90FFFFFF", width: 1);
    text.AppendLine("    <ShapeStyle Name=\"HatOutline\" Fill=\"#00FFFFFF\" Color=\"#FF2B3038\" StrokeWidth=\"2.5\">");
    text.AppendLine("      <Dark Color=\"#FFF2F2F2\" />");
    text.AppendLine("    </ShapeStyle>");
    text.AppendLine("  </Styles>");
    text.AppendLine("  <Figures>");

    // the grid: constant points, hexagons around the hat, each with its six kite lines
    var gridPoints = new Dictionary<(int, int), string>();
    string GridPoint((int I, int J) p)
    {
        if (!gridPoints.TryGetValue(p, out var name))
        {
            name = "G" + (gridPoints.Count + 1);
            gridPoints[p] = name;
            var point = Cartesian(p);
            text.AppendLine(string.Format(invariant, "    <PointByCoordinates Name=\"{0}\" Visible=\"false\" X=\"{1:0.####}\" Y=\"{2:0.####}\" />", name, point.X * Unit, point.Y * Unit));
        }

        return name;
    }

    // a big hexagon of them, two rings around the one the hat has four kites of
    var middle = mainKites.GroupBy(kite => kite.Hexagon).OrderByDescending(group => group.Count()).First().Key;
    var hexagons = new List<(int I, int J)>();
    for (int a = -2; a <= 2; a++)
    {
        for (int b = -2; b <= 2; b++)
        {
            if (Math.Abs(a + b) <= 2)
            {
                hexagons.Add((middle.I + 2 * a - 2 * b, middle.J + 2 * a + 4 * b));
            }
        }
    }

    hexagons = hexagons.OrderBy(c => c.J).ThenBy(c => c.I).ToList();

    // the kites, then the hexagons' sides over them; polygons all, since a segment is drawn
    // over every polygon and the hat must be over the grid
    var figures = new StringBuilder();
    int index = 0;
    foreach (var c in hexagons)
    {
        foreach (var kite in KitesOf(c))
        {
            index++;
            AppendPolygon(figures, "GridKite" + index, "GridKite", kite.Select(GridPoint));
        }
    }

    var corners = Enumerable.Range(0, 6).Select(k => Orient((2, 0), k, flip: false)).ToArray();
    index = 0;
    foreach (var c in hexagons)
    {
        index++;
        AppendPolygon(figures, "Hexagon" + index, "Hexagon", corners.Select(p => GridPoint((c.I + p.I, c.J + p.J))));
    }

    text.Append(figures);

    // H: where the hat was dragged to; Z: the hexagon center nearest to it (rounded in the
    // lattice of the centers, (3, sqrt 3) and (0, 2 sqrt 3) in grid units); R turns it
    string m = "round(H.X / 1.5)";
    string n = "round((H.Y - sqrt(3) / 2 * " + m + ") / sqrt(3))";
    text.AppendLine("    <FreePoint Name=\"H\" Visible=\"false\" X=\"0\" Y=\"0\" />");
    text.AppendLine("    <PointByCoordinates Name=\"Z\" Visible=\"false\" X=\"1.5 * " + m + "\" Y=\"sqrt(3) / 2 * " + m + " + sqrt(3) * " + n + "\">");
    text.AppendLine("      <Dependency Name=\"H\" />");
    text.AppendLine("    </PointByCoordinates>");
    const double TurnRadius = 2.8;
    text.AppendLine(string.Format(invariant, "    <PointByCoordinates Name=\"Z2\" Visible=\"false\" X=\"Z.X + {0}\" Y=\"Z.Y\">", TurnRadius));
    text.AppendLine("      <Dependency Name=\"Z\" />");
    text.AppendLine("    </PointByCoordinates>");
    text.AppendLine("    <Circle Name=\"Turn\" Style=\"TurnCircle\">");
    text.AppendLine("      <Dependency Name=\"Z\" />");
    text.AppendLine("      <Dependency Name=\"Z2\" />");
    text.AppendLine("    </Circle>");

    // R starts at the top of the circle, the hat as it is
    text.AppendLine(string.Format(invariant, "    <PointOnFigure Name=\"R\" X=\"0\" Y=\"{0}\" Parameter=\"{1}\">", TurnRadius, Math.PI / 2));
    text.AppendLine("      <Dependency Name=\"Turn\" />");
    text.AppendLine("    </PointOnFigure>");
    string turn = "round((atan2(R.Y - Z.Y, R.X - Z.X) - pi / 2) / (pi / 3)) * pi / 3";
    text.AppendLine("    <PointByCoordinates Name=\"Q\" Visible=\"false\" X=\"cos(" + turn + ")\" Y=\"sin(" + turn + ")\">");
    text.AppendLine("      <Dependency Name=\"Z\" />");
    text.AppendLine("      <Dependency Name=\"R\" />");
    text.AppendLine("    </PointByCoordinates>");

    var hatPoints = new Dictionary<(int, int), string>();
    string HatPoint((int I, int J) p)
    {
        if (!hatPoints.TryGetValue(p, out var name))
        {
            name = "K" + (hatPoints.Count + 1);
            hatPoints[p] = name;
            var point = Cartesian(p);
            double x = point.X * Unit;
            double y = point.Y * Unit;
            text.AppendLine(string.Format(invariant,
                "    <PointByCoordinates Name=\"{0}\" Visible=\"false\" X=\"Z.X + {1:0.######} * Q.X - {2:0.######} * Q.Y\" Y=\"Z.Y + {1:0.######} * Q.Y + {2:0.######} * Q.X\">",
                name, x, y));
            text.AppendLine("      <Dependency Name=\"Z\" />");
            text.AppendLine("      <Dependency Name=\"Q\" />");
            text.AppendLine("    </PointByCoordinates>");
        }

        return name;
    }

    // the kites, by the hexagon they come from, the one with the most first
    var byHexagon = mainKites
        .GroupBy(kite => kite.Hexagon)
        .OrderByDescending(group => group.Count())
        .ToList();
    var kiteFigures = new StringBuilder();
    index = 0;
    for (int h = 0; h < byHexagon.Count; h++)
    {
        foreach (var kite in byHexagon[h])
        {
            index++;
            var names = kite.Corners.Select(HatPoint).ToList();
            kiteFigures.AppendLine("    <Polygon Name=\"HatKite" + index + "\" Style=\"Kite" + (h + 1) + "\">");
            foreach (var name in names)
            {
                kiteFigures.AppendLine("      <Dependency Name=\"" + name + "\" />");
            }

            kiteFigures.AppendLine("    </Polygon>");
        }
    }

    var outline = hat.Select(HatPoint).ToList();
    text.Append(kiteFigures);
    text.AppendLine("    <Polygon Name=\"Hat\" Style=\"HatOutline\">");
    foreach (var name in outline)
    {
        text.AppendLine("      <Dependency Name=\"" + name + "\" />");
    }

    text.AppendLine("    </Polygon>");
    var description = new StringBuilder();
    description.Append("Cut every hexagon of a grid into six kites, from its center to the middle of each side. Each kite has angles of 60°, 90°, 120° and 90°. ");
    description.Append("The hat is eight of these kites stuck together: " + Words(counts) + ", a color for each. ");
    description.Append("So whenever copies of the hat cover the plane, they lie on such a grid, kite on kite.\\n\\n");
    description.Append("Drag the hat around: it jumps from hexagon to hexagon. Drag the green point to turn it.");
    Finish(text, path, "Eight Kites Make a Hat", description.ToString());
}

bool WriteMorph(string path)
{
    // the patch: the copies around the middle one
    var patch = shown.Where(placement => Distance(Centroid(placement), center) < MorphRadius).ToList();

    // every corner as a sum of short sides (a, in lattice units) and long sides (b, in units
    // of w0 = (1, 1) and w1 = (-1, 2)), found by walking the outlines from the first corner
    var sums = new Dictionary<(int, int), (int AI, int AJ, int BI, int BJ)>();
    sums[Vertices(main).First()] = (0, 0, 0, 0);
    bool grew = true;
    while (grew)
    {
        grew = false;
        foreach (var placement in patch)
        {
            var corners = Vertices(placement).ToList();
            for (int k = 0; k < corners.Count * 2; k++)
            {
                var p = corners[k % corners.Count];
                var q = corners[(k + 1) % corners.Count];
                if (!sums.TryGetValue(p, out var sum))
                {
                    continue;
                }

                var step = (I: q.I - p.I, J: q.J - p.J);
                (int, int, int, int) next;
                if (IsShortSide(step))
                {
                    next = (sum.AI + step.I, sum.AJ + step.J, sum.BI, sum.BJ);
                }
                else
                {
                    int bj = (step.J - step.I) / 3;
                    int bi = step.I + bj;
                    next = (sum.AI, sum.AJ, sum.BI + bi, sum.BJ + bj);
                }

                if (sums.TryGetValue(q, out var known))
                {
                    if (known != next)
                    {
                        Console.WriteLine("the sides don't add up around " + q);
                        return false;
                    }
                }
                else
                {
                    sums[q] = next;
                    grew = true;
                }
            }
        }
    }

    // a corner at a = sin, b = cos: the short sides are unit vectors at multiples of 60
    // degrees, the long ones at 30 degrees off; the middle hat's corners average to (0, 0)
    const double Size = 1.2;
    (double X, double Y) ShortPart((int AI, int AJ, int BI, int BJ) sum) => (sum.AI + sum.AJ / 2.0, sum.AJ * root3 / 2);
    (double X, double Y) LongPart((int AI, int AJ, int BI, int BJ) sum) => (sum.BI * root3 / 2, sum.BI / 2.0 + sum.BJ);
    var mainSums = Vertices(main).Select(p => sums[p]).ToList();
    var shortMiddle = (X: mainSums.Average(s => ShortPart(s).X), Y: mainSums.Average(s => ShortPart(s).Y));
    var longMiddle = (X: mainSums.Average(s => LongPart(s).X), Y: mainSums.Average(s => LongPart(s).Y));
    const double PatchY = 0.6;

    var text = Start();
    TileStyles(text);
    text.AppendLine("    <LineStyle Name=\"Tick\" Color=\"#FF9BA2AC\" StrokeWidth=\"2\">");
    text.AppendLine("      <Dark Color=\"#FF7A818C\" />");
    text.AppendLine("    </LineStyle>");
    text.AppendLine("    <TextStyle Name=\"TickText\" FontSize=\"15\" Color=\"#FF4A515C\" FontFamily=\"Segoe UI\">");
    text.AppendLine("      <Dark Color=\"#FFB5BBC6\" />");
    text.AppendLine("    </TextStyle>");
    text.AppendLine("  </Styles>");
    text.AppendLine("  <Figures>");

    // the track and the point on it, which starts at the hat
    const double TrackLeft = -6;
    const double TrackRight = 6;
    const double TrackY = -6.8;
    const double StartPlace = 1.0 / 3;
    text.AppendLine(string.Format(invariant, "    <PointByCoordinates Name=\"S1\" Visible=\"false\" X=\"{0}\" Y=\"{1}\" />", TrackLeft, TrackY));
    text.AppendLine(string.Format(invariant, "    <PointByCoordinates Name=\"S2\" Visible=\"false\" X=\"{0}\" Y=\"{1}\" />", TrackRight, TrackY));
    text.AppendLine("    <Segment Name=\"Track\" Style=\"SliderTrack\">");
    text.AppendLine("      <Dependency Name=\"S1\" />");
    text.AppendLine("      <Dependency Name=\"S2\" />");
    text.AppendLine("    </Segment>");

    // a tick and a name at each shape of the paper
    var stops = new[] { (0.0, "chevron"), (1.0 / 3, "hat"), (0.5, "spectre"), (2.0 / 3, "turtle"), (1.0, "comet") };
    int index = 0;
    foreach (var (place, name) in stops)
    {
        index++;
        // the name under the tick, the spectre's above it, which is too near to its neighbors
        double x = TrackLeft + (TrackRight - TrackLeft) * place;
        bool above = name == "spectre";
        text.AppendLine(string.Format(invariant, "    <PointByCoordinates Name=\"T{0}a\" Visible=\"false\" X=\"{1:0.####}\" Y=\"{2}\" />", index, x, TrackY + (above ? -0.25 : 0.25)));
        text.AppendLine(string.Format(invariant, "    <PointByCoordinates Name=\"{0}\" Visible=\"false\" X=\"{1:0.####}\" Y=\"{2}\" />", name, x, TrackY + (above ? 0.25 : -0.25)));
        text.AppendLine("    <Segment Name=\"Tick" + index + "\" Style=\"Tick\">");
        text.AppendLine("      <Dependency Name=\"T" + index + "a\" />");
        text.AppendLine("      <Dependency Name=\"" + name + "\" />");
        text.AppendLine("    </Segment>");
        text.AppendLine(string.Format(invariant,
            "    <PointLabel Name=\"{0}Name\" Style=\"TickText\" OffsetX=\"{1}\" OffsetY=\"{2}\" ShowName=\"true\" ShowCoordinates=\"false\">",
            name, -name.Length * 3.6, above ? -26 : 6));
        text.AppendLine("      <Dependency Name=\"" + name + "\" />");
        text.AppendLine("    </PointLabel>");
    }

    text.AppendLine(string.Format(invariant, "    <PointOnFigure Name=\"T\" X=\"{0:0.####}\" Y=\"{1}\" Parameter=\"{2:0.##########}\">", TrackLeft + (TrackRight - TrackLeft) * StartPlace, TrackY, StartPlace));
    text.AppendLine("      <Dependency Name=\"Track\" />");
    text.AppendLine("    </PointOnFigure>");

    // L: the lengths of the two kinds of sides, sin and cos of 90 degrees times the place
    string angle = string.Format(invariant, "pi / 2 * (T.X - {0}) / {1}", TrackLeft, TrackRight - TrackLeft).Replace("- -", "+ ");
    text.AppendLine(string.Format(invariant, "    <PointByCoordinates Name=\"L\" Visible=\"false\" X=\"{0} * sin({1})\" Y=\"{0} * cos({1})\">", Size, angle));
    text.AppendLine("      <Dependency Name=\"T\" />");
    text.AppendLine("    </PointByCoordinates>");

    var names = new Dictionary<(int, int), string>();
    string Corner((int I, int J) corner)
    {
        if (!names.TryGetValue(corner, out var name))
        {
            name = "P" + (names.Count + 1);
            names[corner] = name;
            var sum = sums[corner];
            var shortPart = ShortPart(sum);
            var longPart = LongPart(sum);
            text.AppendLine(string.Format(invariant,
                "    <PointByCoordinates Name=\"{0}\" Visible=\"false\" X=\"{1:0.######} * L.X + {2:0.######} * L.Y\" Y=\"{5} + {3:0.######} * L.X + {4:0.######} * L.Y\">",
                name,
                shortPart.X - shortMiddle.X,
                longPart.X - longMiddle.X,
                shortPart.Y - shortMiddle.Y,
                longPart.Y - longMiddle.Y,
                PatchY));
            text.AppendLine("      <Dependency Name=\"L\" />");
            text.AppendLine("    </PointByCoordinates>");
        }

        return name;
    }

    var tiles = new StringBuilder();
    index = 0;
    foreach (var placement in patch.Where(placement => placement != main).Append(main))
    {
        index++;
        string name = placement == main ? "Tile" : "Tile" + (index + 1);
        var corners = Vertices(placement).Select(Corner).ToList();
        tiles.AppendLine("    <Polygon Name=\"" + name + "\" Style=\"" + colorOf[placement] + "\">");
        foreach (var corner in corners)
        {
            tiles.AppendLine("      <Dependency Name=\"" + corner + "\" />");
        }

        tiles.AppendLine("    </Polygon>");
    }

    text.Append(tiles);
    Console.WriteLine($"{patch.Count} tiles in the morph, {patch.Count(IsFlipped)} flipped");
    Finish(text, path, "From Hat to Spectre",
        "Every side of the hat is a short or a long side of one of its kites. Give all the short sides one length and all the long ones another, keep the angles, and the copies still fit together the same way.\\n\\nSlide the green point. Every shape on the way is an aperiodic monotile, except the two ends, which tile in a repeating pattern, and the spectre, whose sides are all equal: it repeats too, but only with flipped copies. Without them it never does, and with curvy sides it can't be flipped at all.");
    return true;
}

void AppendPolygon(StringBuilder figures, string name, string style, IEnumerable<string> points)
{
    figures.AppendLine("    <Polygon Name=\"" + name + "\" Style=\"" + style + "\">");
    foreach (var point in points)
    {
        figures.AppendLine("      <Dependency Name=\"" + point + "\" />");
    }

    figures.AppendLine("    </Polygon>");
}

// the styles of the tiles, shared by the drawings: gentle gradients on the white paper,
// bright-to-deep ones on the dark paper
void TileStyles(StringBuilder text)
{
    Gradient(text, "HatMain", "#FFFFA375", "#FFFF7448", "#FFFFD27A", "#FFC0530A");
    Gradient(text, "HatFlipped", "#FF9B7EE8", "#FF5B3BB0", "#FFD6B4FF", "#FF5B1FB0");
    Gradient(text, "Hat1", "#FFD3EBFC", "#FFA6D0F2", "#FF36C0FA", "#FF052C3C");
    Gradient(text, "Hat2", "#FFDDF3D6", "#FFB2DFA9", "#FF8FDFFF", "#FF0181B3");
    Gradient(text, "Hat3", "#FFFFF2C8", "#FFF5D98A", "#FF3FE0C0", "#FF04362E");
}

void Gradient(StringBuilder text, string name, string from, string to, string darkFrom, string darkTo, string outline = "#FF3A3F47", string darkOutline = "#FF101215", double width = 1.5)
{
    text.AppendLine(string.Format(invariant, "    <ShapeStyle Name=\"{0}\" Color=\"{1}\" StrokeWidth=\"{2}\">", name, outline, width));
    text.AppendLine("      <Fill>");
    GradientBrush(text, from, to, "        ");
    text.AppendLine("      </Fill>");
    text.AppendLine("      <Dark Color=\"" + darkOutline + "\">");
    text.AppendLine("        <Fill>");
    GradientBrush(text, darkFrom, darkTo, "          ");
    text.AppendLine("        </Fill>");
    text.AppendLine("      </Dark>");
    text.AppendLine("    </ShapeStyle>");
}

void GradientBrush(StringBuilder text, string from, string to, string indent)
{
    text.AppendLine(indent + "<LinearGradientBrush StartPoint=\"0,0\" EndPoint=\"1,1\">");
    text.AppendLine(indent + "  <GradientStop Color=\"" + from + "\" Offset=\"0\" />");
    text.AppendLine(indent + "  <GradientStop Color=\"" + to + "\" Offset=\"1\" />");
    text.AppendLine(indent + "</LinearGradientBrush>");
}

StringBuilder Start()
{
    var text = new StringBuilder();
    text.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
    text.AppendLine("<Drawing Version=\"1\">");
    text.AppendLine("  <Viewport Left=\"-6\" Top=\"4\" Right=\"6\" Bottom=\"-4\" />");
    text.AppendLine("  <Styles>");
    return text;
}

void Finish(StringBuilder text, string path, string title, string description)
{
    text.AppendLine("    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"" + title + "\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
    text.AppendLine("    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"" + description + "\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\" />");
    text.AppendLine("  </Figures>");
    text.Append("</Drawing>");
    File.WriteAllText(path, text.ToString().Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(false));
    Console.WriteLine("wrote " + path);
}

// 4 + 2 + 2: "four from one hexagon and two from each of two others"
string Words(List<int> kiteCounts)
{
    string[] numbers = { "zero", "one", "two", "three", "four", "five", "six" };
    var groups = kiteCounts.GroupBy(count => count).ToList();
    var parts = groups.Select(group => group.Count() == 1
        ? numbers[group.Key] + " from " + (group == groups[0] ? "one hexagon" : "another")
        : numbers[group.Key] + " from each of " + numbers[group.Count()] + " others");
    return string.Join(" and ", parts);
}

IEnumerable<((int, int), (int, int))> InnerKiteEdges()
{
    var edges = new Dictionary<((int, int), (int, int)), int>();
    foreach (var (_, kite) in mainKites)
    {
        for (int l = 0; l < 4; l++)
        {
            var p = kite[l];
            var q = kite[(l + 1) % 4];
            var edge = p.CompareTo(q) < 0 ? ((p.I, p.J), (q.I, q.J)) : ((q.I, q.J), (p.I, p.J));
            edges[edge] = edges.GetValueOrDefault(edge) + 1;
        }
    }

    return edges.Where(pair => pair.Value == 2).Select(pair => pair.Key);
}

IEnumerable<(int I, int J)> HexagonCenters(int reach)
{
    for (int a = -reach; a <= reach; a++)
    {
        for (int b = -reach; b <= reach; b++)
        {
            yield return (2 * a - 2 * b, 2 * a + 4 * b);
        }
    }
}

IEnumerable<(int I, int J)[]> KitesOf((int I, int J) hexagonCenter)
{
    var side = (I: 1, J: 1);
    var corner = (I: 0, J: 2);
    for (int k = 0; k < 6; k++)
    {
        var nextSide = Orient(side, 1, flip: false);
        yield return new[] { (0, 0), side, corner, nextSide }
            .Select(p => (hexagonCenter.I + p.Item1, hexagonCenter.J + p.Item2))
            .ToArray();
        side = nextSide;
        corner = Orient(corner, 1, flip: false);
    }
}

// a short side of a kite goes along a lattice direction (1 or 2 units), a long one 30
// degrees off it (sqrt 3)
static bool IsShortSide((int I, int J) step)
{
    return step.I == 0 || step.J == 0 || step.I == -step.J;
}

bool Color(int index)
{
    if (index == rest.Count)
    {
        return true;
    }

    var placement = rest[index];
    foreach (var color in others.OrderBy(color => colorOf.Values.Count(value => value == color)))
    {
        if (neighbors[placement].Any(other => colorOf.TryGetValue(other, out var taken) && taken == color))
        {
            continue;
        }

        colorOf[placement] = color;
        if (Color(index + 1))
        {
            return true;
        }

        colorOf.Remove(placement);
    }

    return false;
}

string Number(int value)
{
    return value < 0 ? "m" + (-value).ToString(invariant) : value.ToString(invariant);
}

bool Search()
{
    steps++;
    if (steps > 5_000_000)
    {
        return false;
    }

    var free = targets.FirstOrDefault(cell => !occupied.ContainsKey(cell), (int.MinValue, 0));
    if (free.Item1 == int.MinValue)
    {
        return true;
    }

    var candidates = new List<Placement>();
    for (int o = 0; o < orientations.Count; o++)
    {
        foreach (var cell in cellsOf[o])
        {
            var shift = (cell.Item1 - free.Item1, cell.Item2 - free.Item2);

            // the shift in lattice units, kept only when it is between centers of hexagons
            if (shift.Item1 % 18 != 0 || shift.Item2 % 18 != 0)
            {
                continue;
            }

            var translation = (I: -shift.Item1 / 18, J: -shift.Item2 / 18);
            if (!IsHexagonShift(translation))
            {
                continue;
            }

            var placement = new Placement(o, translation);
            if (Fits(placement))
            {
                candidates.Add(placement);
            }
        }
    }

    foreach (var placement in candidates.OrderBy(_ => random.Next()))
    {
        Place(placement);
        if (Search())
        {
            return true;
        }

        Remove(placement);
    }

    return false;
}

bool Fits(Placement placement)
{
    foreach (var cell in cellsOf[placement.Orientation])
    {
        if (occupied.ContainsKey(Shift(cell, placement.Translation)))
        {
            return false;
        }
    }

    return true;
}

void Place(Placement placement)
{
    placements.Add(placement);
    foreach (var cell in cellsOf[placement.Orientation])
    {
        occupied[Shift(cell, placement.Translation)] = placements.Count - 1;
    }
}

void Remove(Placement placement)
{
    placements.RemoveAt(placements.Count - 1);
    foreach (var cell in cellsOf[placement.Orientation])
    {
        occupied.Remove(Shift(cell, placement.Translation));
    }
}

bool Touch(Placement first, Placement second)
{
    // a point just outside the middle of a piece of the first one's edge is in the second
    var mine = Vertices(first).Select(Cartesian).ToArray();
    var theirs = Vertices(second).Select(Cartesian).ToArray();
    for (int k = 0; k < mine.Length; k++)
    {
        var p = mine[k];
        var q = mine[(k + 1) % mine.Length];
        var length = Distance(p, q);
        var outward = ((q.Y - p.Y) / length * 0.05, -(q.X - p.X) / length * 0.05);
        foreach (var t in new[] { 0.25, 0.5, 0.75 })
        {
            var point = (p.X + (q.X - p.X) * t, p.Y + (q.Y - p.Y) * t);
            if (Inside(theirs, (point.Item1 + outward.Item1, point.Item2 + outward.Item2))
                || Inside(theirs, (point.Item1 - outward.Item1, point.Item2 - outward.Item2)))
            {
                return true;
            }
        }
    }

    return false;
}

bool IsFlipped(Placement placement)
{
    return placement.Orientation >= 6;
}

IEnumerable<(int I, int J)> Vertices(Placement placement)
{
    return orientations[placement.Orientation].Select(p => (p.I + placement.Translation.I, p.J + placement.Translation.J));
}

(double X, double Y) Centroid(Placement placement)
{
    var points = cellsOf[placement.Orientation].Select(cell => CellPoint(Shift(cell, placement.Translation))).ToList();
    return (points.Average(p => p.X), points.Average(p => p.Y));
}

static (int, int) Shift((int, int) cell, (int I, int J) translation)
{
    return (cell.Item1 + translation.I * 18, cell.Item2 + translation.J * 18);
}

// the hexagons' centers: the lattice of (2, 2) and (-2, 4)
static bool IsHexagonShift((int I, int J) translation)
{
    // i = 2a - 2b, j = 2a + 4b
    int difference = translation.J - translation.I;
    if (difference % 6 != 0)
    {
        return false;
    }

    int b = difference / 6;
    return (translation.I + 2 * b) % 2 == 0;
}

static (int I, int J) Orient((int I, int J) p, int turn, bool flip)
{
    if (flip)
    {
        p = (p.I + p.J, -p.J);
    }

    for (int k = 0; k < turn; k++)
    {
        p = (-p.J, p.I + p.J);
    }

    return p;
}

(double X, double Y) Cartesian((int I, int J) p)
{
    return (p.I + p.J / 2.0, p.J * root3 / 2);
}

// a cell is named by its centroid in eighteenths of lattice units (always whole numbers)
(double X, double Y) CellPoint((int, int) cell)
{
    return (cell.Item1 / 18.0 + cell.Item2 / 36.0, cell.Item2 / 18.0 * root3 / 2);
}

// the six cells of each of the two lattice triangles at (i, j), by their centroids
static IEnumerable<(int, int)> TriangleCells(int i, int j)
{
    var up = new[] { (i, j), (i + 1, j), (i, j + 1) };
    var down = new[] { (i + 1, j), (i + 1, j + 1), (i, j + 1) };
    foreach (var triangle in new[] { up, down })
    {
        // in sixths: the center is the corners' sum * 2, a midpoint the two corners' sum * 3
        var (ax, ay) = (triangle[0].Item1 * 6, triangle[0].Item2 * 6);
        var (bx, by) = (triangle[1].Item1 * 6, triangle[1].Item2 * 6);
        var (cx, cy) = (triangle[2].Item1 * 6, triangle[2].Item2 * 6);
        var g = ((ax + bx + cx) / 3, (ay + by + cy) / 3);
        var corners = new[] { (ax, ay), (bx, by), (cx, cy) };
        for (int k = 0; k < 3; k++)
        {
            var p = corners[k];
            var q = corners[(k + 1) % 3];
            var m = ((p.Item1 + q.Item1) / 2, (p.Item2 + q.Item2) / 2);

            // the two small triangles (p, m, g) and (m, q, g); their centroids are in
            // eighteenths, which is how a cell is named
            yield return Sum(p, m, g);
            yield return Sum(m, q, g);
        }
    }
}

// three points in sixths added up: their centroid in eighteenths
static (int, int) Sum((int, int) p, (int, int) q, (int, int) r)
{
    return (p.Item1 + q.Item1 + r.Item1, p.Item2 + q.Item2 + r.Item2);
}

List<(int, int)> Cells((int I, int J)[] polygon)
{
    var points = polygon.Select(Cartesian).ToArray();
    int minI = polygon.Min(p => p.I) - 4, maxI = polygon.Max(p => p.I) + 4;
    int minJ = polygon.Min(p => p.J) - 4, maxJ = polygon.Max(p => p.J) + 4;
    var result = new List<(int, int)>();
    for (int i = minI; i <= maxI; i++)
    {
        for (int j = minJ; j <= maxJ; j++)
        {
            foreach (var cell in TriangleCells(i, j))
            {
                if (Inside(points, CellPoint(cell)))
                {
                    result.Add(cell);
                }
            }
        }
    }

    return result;
}

static bool Inside((double X, double Y)[] polygon, (double X, double Y) point)
{
    bool inside = false;
    for (int k = 0, l = polygon.Length - 1; k < polygon.Length; l = k++)
    {
        var p = polygon[k];
        var q = polygon[l];
        if ((p.Y > point.Y) != (q.Y > point.Y) && point.X < (q.X - p.X) * (point.Y - p.Y) / (q.Y - p.Y) + p.X)
        {
            inside = !inside;
        }
    }

    return inside;
}

static double Distance((double X, double Y) p, (double X, double Y) q)
{
    return Math.Sqrt((p.X - q.X) * (p.X - q.X) + (p.Y - q.Y) * (p.Y - q.Y));
}

record Placement(int Orientation, (int I, int J) Translation);
