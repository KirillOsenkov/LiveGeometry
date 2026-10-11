#:property Nullable=disable
#:property PublishAot=false

// tessellation - writes the "Make Your Own Tessellation" gallery drawing: one tile, a closed
// Bezier path on a square lattice, whose top edge is copied to the bottom and right edge to
// the left (the copied anchors and handles are points by coordinates, the originals free
// points), so that whatever shape the top and right edges are given, the tile still fits
// itself. Around it, translated copies of the tile cover the plane in two colors.
//
//   dotnet tools/tessellation.cs -- <out.lgf>

using System.Globalization;
using System.Text;

if (args.Length < 1)
{
    Console.WriteLine("usage: tessellation <out.lgf>");
    return 1;
}

const double Side = 4;
var invariant = CultureInfo.InvariantCulture;
var text = new StringBuilder();

void Write(string line)
{
    text.Append(line).Append("\r\n");
}

string Format(double value)
{
    return value.ToString("0.####", invariant);
}

// The tile's anchors clockwise from the top left corner: the corners are fixed, the two
// anchors on the top and on the right edge are free, the ones on the bottom and the left
// are the top's and the right's shifted by a side. Every anchor has an in handle and an out
// handle; on the copied edges they are the shifted handles of the matching anchor, swapped
// (the edge runs the other way round there).
var anchors = new List<Anchor>
{
    new Anchor("TL", (0, Side), free: false, inHandle: null, outHandle: (0.4, 4.4)),
    new Anchor("T1", (1.2, 4.5), free: true, inHandle: (0.8, 4.6), outHandle: (1.7, 4.4)),
    new Anchor("T2", (2.8, 3.6), free: true, inHandle: (2.4, 3.5), outHandle: (3.2, 3.7)),
    new Anchor("TR", (Side, Side), free: false, inHandle: (3.6, 4.1), outHandle: (4.1, 3.6)),
    new Anchor("R1", (4.4, 2.9), free: true, inHandle: (4.5, 3.3), outHandle: (4.3, 2.5)),
    new Anchor("R2", (3.6, 1.2), free: true, inHandle: (3.5, 1.6), outHandle: (3.7, 0.8)),
    new Anchor("BR", (Side, 0), free: false, inHandle: (3.9, 0.4), outHandle: null),
    new Anchor("B2", null, free: false, inHandle: null, outHandle: null) { Source = "T2", Shift = (0, -Side) },
    new Anchor("B1", null, free: false, inHandle: null, outHandle: null) { Source = "T1", Shift = (0, -Side) },
    new Anchor("BL", (0, 0), free: false, inHandle: null, outHandle: null),
    new Anchor("L2", null, free: false, inHandle: null, outHandle: null) { Source = "R2", Shift = (-Side, 0) },
    new Anchor("L1", null, free: false, inHandle: null, outHandle: null) { Source = "R1", Shift = (-Side, 0) },
};

// which handle of which anchor each copied handle is the shift of: the piece from TR to R1
// becomes the piece from L1 to TL, so L1's out is R1's in shifted and TL's in is TR's out
// shifted, and so on around
var shiftedHandles = new Dictionary<string, (string Anchor, string Handle, (double X, double Y) Shift)>
{
    ["BR.out"] = ("TR", "in", (0, -Side)),
    ["B2.in"] = ("T2", "out", (0, -Side)),
    ["B2.out"] = ("T2", "in", (0, -Side)),
    ["B1.in"] = ("T1", "out", (0, -Side)),
    ["B1.out"] = ("T1", "in", (0, -Side)),
    ["BL.in"] = ("TL", "out", (0, -Side)),
    ["BL.out"] = ("BR", "in", (-Side, 0)),
    ["L2.in"] = ("R2", "out", (-Side, 0)),
    ["L2.out"] = ("R2", "in", (-Side, 0)),
    ["L1.in"] = ("R1", "out", (-Side, 0)),
    ["L1.out"] = ("R1", "in", (-Side, 0)),
    ["TL.in"] = ("TR", "out", (-Side, 0)),
};

Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
Write("<Drawing Version=\"1\" Creator=\"LiveGeometry.App\">");
Write("  <Viewport Left=\"-5\" Top=\"9\" Right=\"13\" Bottom=\"-5\" />");
Write("  <Styles>");
Write("    <ShapeStyle Name=\"TileA\" IsFilled=\"true\" Fill=\"#FF5FB8A8\" Color=\"#FF1E6B5E\" StrokeWidth=\"1.5\">");
Write("      <Dark Fill=\"#FF2E8A78\" Color=\"#FF8FE0D0\" />");
Write("    </ShapeStyle>");
Write("    <ShapeStyle Name=\"TileB\" IsFilled=\"true\" Fill=\"#FFF2A65A\" Color=\"#FF8A4A10\" StrokeWidth=\"1.5\">");
Write("      <Dark Fill=\"#FFB8712E\" Color=\"#FFFFCC99\" />");
Write("    </ShapeStyle>");
Write("    <LineStyle Name=\"EdgeA\" Color=\"#FF1E6B5E\" StrokeWidth=\"1.5\">");
Write("      <Dark Color=\"#FF8FE0D0\" />");
Write("    </LineStyle>");
Write("    <LineStyle Name=\"EdgeB\" Color=\"#FF8A4A10\" StrokeWidth=\"1.5\">");
Write("      <Dark Color=\"#FFFFCC99\" />");
Write("    </LineStyle>");
Write("    <PointStyle Name=\"Anchor\" Size=\"11\" Fill=\"#FFFFE066\" Color=\"#FF806000\" StrokeWidth=\"1.5\" />");
Write("    <PointStyle Name=\"Arm\" Shape=\"Square\" Size=\"7\" Fill=\"#FFFFFFFF\" Color=\"#FF806000\" StrokeWidth=\"1.5\" />");
Write("    <PointStyle Name=\"Fish\" Character=\"\U0001F41F\" Size=\"34\" Fill=\"#FFFFFFFF\" />");
Write("    <PointStyle Name=\"Fish2\" Character=\"\U0001F420\" Size=\"34\" Fill=\"#FFFFFFFF\" />");
Write("  </Styles>");
Write("  <Figures>");

// the original tile's points: anchors, then the handles in the path's order (in1, out1,
// in2, out2...) so that the Path string can name them by their place in the list
var dependencies = new List<string>();
foreach (var anchor in anchors)
{
    if (anchor.Free)
    {
        Write($"    <FreePoint Name=\"{anchor.Name}\" Style=\"Anchor\" X=\"{Format(anchor.Place.Value.X)}\" Y=\"{Format(anchor.Place.Value.Y)}\" />");
    }
    else if (anchor.Source != null)
    {
        Write($"    <PointByCoordinates Name=\"{anchor.Name}\" Visible=\"false\" X=\"{anchor.Source}.X + {Format(anchor.Shift.X)}\" Y=\"{anchor.Source}.Y + {Format(anchor.Shift.Y)}\">");
        Write($"      <Dependency Name=\"{anchor.Source}\" />");
        Write("    </PointByCoordinates>");
    }
    else
    {
        Write($"    <PointByCoordinates Name=\"{anchor.Name}\" Visible=\"false\" X=\"{Format(anchor.Place.Value.X)}\" Y=\"{Format(anchor.Place.Value.Y)}\" />");
    }

    dependencies.Add(anchor.Name);
}

foreach (var anchor in anchors)
{
    foreach (var which in new[] { "in", "out" })
    {
        string name = anchor.Name + (which == "in" ? "In" : "Out");
        var place = which == "in" ? anchor.InHandle : anchor.OutHandle;
        if (place != null)
        {
            Write($"    <FreePoint Name=\"{name}\" Style=\"Arm\" X=\"{Format(place.Value.X)}\" Y=\"{Format(place.Value.Y)}\" />");
        }
        else
        {
            var (sourceAnchor, sourceHandle, shift) = shiftedHandles[anchor.Name + "." + which];
            string source = sourceAnchor + (sourceHandle == "in" ? "In" : "Out");
            Write($"    <PointByCoordinates Name=\"{name}\" Visible=\"false\" X=\"{source}.X + {Format(shift.X)}\" Y=\"{source}.Y + {Format(shift.Y)}\">");
            Write($"      <Dependency Name=\"{source}\" />");
            Write("    </PointByCoordinates>");
        }

        dependencies.Add(name);
    }
}

// the path: a cubic per anchor, from its out handle to the next anchor's in handle, by
// their places among the dependencies
int count = anchors.Count;
var pieces = new List<string>();
for (int i = 0; i < count; i++)
{
    int next = (i + 1) % count;
    pieces.Add($"C #{count + 2 * i + 1} #{count + 2 * next}");
}

string path = string.Join(" ", pieces);

void WritePath(string name, string style, string edge, IEnumerable<string> points)
{
    Write($"    <BezierPath Name=\"{name}\" Style=\"{style}\" Closed=\"true\" Filled=\"true\" Path=\"{path}\">");
    Write($"      <Sides Style=\"{edge}\" />");
    foreach (var point in points)
    {
        Write($"      <Dependency Name=\"{point}\" />");
    }

    Write("    </BezierPath>");
}

// the copies: every point shifted by whole sides, the tile painted in the checkerboard's
// other color where a row or a column is odd
var offsets = new List<(int I, int J)>();
for (int j = -1; j <= 1; j++)
{
    for (int i = -1; i <= 2; i++)
    {
        offsets.Add((i, j));
    }
}

foreach (var (i, j) in offsets)
{
    bool even = (i + j) % 2 == 0;
    string style = even ? "TileA" : "TileB";
    string edge = even ? "EdgeA" : "EdgeB";
    if (i == 0 && j == 0)
    {
        WritePath("Tile", style, edge, dependencies);
    }
    else
    {
        string suffix = $"_{i + 1}_{j + 1}";
        foreach (var point in dependencies)
        {
            Write($"    <PointByCoordinates Name=\"{point}{suffix}\" Visible=\"false\" X=\"{point}.X + {Format(i * Side)}\" Y=\"{point}.Y + {Format(j * Side)}\">");
            Write($"      <Dependency Name=\"{point}\" />");
            Write("    </PointByCoordinates>");
        }

        WritePath("Tile" + suffix, style, edge, dependencies.Select(point => point + suffix));
    }

    // a fish in the middle of every tile, swimming with the tile's shape
    string fish = $"Fish_{i + 1}_{j + 1}";
    Write($"    <PointByCoordinates Name=\"{fish}\" Style=\"{(even ? "Fish" : "Fish2")}\" X=\"(T1.X + T2.X + R1.X + R2.X) / 4 - 0.4 + {Format(i * Side)}\" Y=\"(T1.Y + T2.Y + R1.Y + R2.Y) / 4 - 0.6 + {Format(j * Side)}\">");
    Write("      <Dependency Name=\"T1\" />");
    Write("      <Dependency Name=\"T2\" />");
    Write("      <Dependency Name=\"R1\" />");
    Write("      <Dependency Name=\"R2\" />");
    Write("    </PointByCoordinates>");
}

Write("    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"Make Your Own Tessellation\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"Squares cover a floor with no gaps. So does this shape, and so does any shape you can turn it into: drag the yellow points and the white squares on the one tile with the yellow points, and all the copies change with it.\\n\\nThe trick: whatever you do to the top edge happens to the bottom edge too, and the left edge copies the right one. So every bump on one side is a matching dent on the other, and the tiles still lock together.\\n\\nThe Dutch artist M. C. Escher filled whole pages this way with lizards, birds and fish. Can you make a fish shape to go with the fish? Or a bird?\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("  </Figures>");
text.Append("</Drawing>");
File.WriteAllText(args[0], text.ToString(), new UTF8Encoding(false));
Console.WriteLine("wrote " + args[0]);
return 0;

class Anchor
{
    public Anchor(string name, (double X, double Y)? place, bool free, (double X, double Y)? inHandle, (double X, double Y)? outHandle)
    {
        Name = name;
        Place = place;
        Free = free;
        InHandle = inHandle;
        OutHandle = outHandle;
    }

    public string Name { get; }

    public (double X, double Y)? Place { get; }

    public bool Free { get; }

    public (double X, double Y)? InHandle { get; }

    public (double X, double Y)? OutHandle { get; }

    /// <summary>The anchor this one is a shifted copy of, for the copied edges</summary>
    public string Source { get; set; }

    public (double X, double Y) Shift { get; set; }
}
