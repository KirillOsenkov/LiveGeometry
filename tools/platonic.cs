#:property Nullable=disable
#:property PublishAot=false

// platonic - writes the "5 Platonic Solids" gallery drawing: the five solids projected onto
// the plane as fixed polygons (PointByCoordinates with constant coordinates, nothing to drag),
// one face per polygon with its own shaded style, painted back to front, a name under each.
//
//   dotnet tools/platonic.cs -- <out.lgf> [--alpha]
//
// --alpha keeps the back faces and makes every face translucent, so the whole solid shows
// through (the default drops the back faces and paints the front ones opaque).

using System.Globalization;
using System.Text;

if (args.Length < 1)
{
    Console.WriteLine("usage: platonic <out.lgf> [--alpha]");
    return 1;
}

bool alpha = args.Contains("--alpha");
var invariant = CultureInfo.InvariantCulture;
double phi = (1 + Math.Sqrt(5)) / 2;

// the light comes from the upper left, a little in front
var light = Normalize((-0.45, 0.8, 0.55));

var solids = new[]
{
    // soft colors, in the spirit of the app's pastels but with enough body to shade
    // seen a little from above: a positive nod
    new Solid("Tetrahedron", (242, 150, 128), Tilt(18, 20, 0), Tetrahedron()),
    new Solid("Cube", (128, 176, 240), Tilt(24, 36, 0), Signs3(1, 1, 1)),
    new Solid("Octahedron", (128, 206, 150), Tilt(22, 28, 0), new[] { V(1, 0, 0), V(-1, 0, 0), V(0, 1, 0), V(0, -1, 0), V(0, 0, 1), V(0, 0, -1) }),
    new Solid("Dodecahedron", (192, 158, 232), Tilt(20, 22, 0), Dodecahedron()),
    new Solid("Icosahedron", (246, 200, 110), Tilt(18, 26, 0), Icosahedron()),
};

// like the five on a die: four in the corners, the cube in the middle (names that share a
// row are then far enough apart on a phone), each solid about 1.9 units across
var centers = new[] { (-2.9, 1.7), (0.0, 0.0), (2.9, 1.7), (-2.9, -1.7), (2.9, -1.7) };
const double SolidSize = 1.9;
const double LabelDrop = 1.25;

var styles = new StringBuilder();
var figures = new StringBuilder();
for (int s = 0; s < solids.Length; s++)
{
    var solid = solids[s];
    var (cx, cy) = centers[s];

    // rotate, then scale so that the projection is SolidSize across
    var turned = solid.Vertices.Select(v => Rotate(v, solid.Tilt)).ToArray();
    double extent = turned.Max(v => Math.Max(Math.Abs(v.X), Math.Abs(v.Y))) * 2;
    double scale = SolidSize / extent;
    var projected = turned.Select(v => (X: v.X * scale, Y: v.Y * scale, Z: v.Z * scale)).ToArray();

    for (int i = 0; i < projected.Length; i++)
    {
        figures.AppendLine(string.Format(invariant,
            "    <PointByCoordinates Name=\"{0}{1}\" Visible=\"false\" Style=\"DependentPointStyle\" X=\"{2:0.####}\" Y=\"{3:0.####}\" />",
            solid.Name, i, cx + projected[i].X, cy + projected[i].Y));
    }

    // back to front, the back faces only when they are to show through
    var faces = Faces(projected)
        .Where(face => alpha || face.Normal.Z > 0)
        .OrderBy(face => face.Depth)
        .ToArray();
    int faceIndex = 0;
    foreach (var face in faces)
    {
        string styleName = solid.Name + "Face" + faceIndex;
        double lit = Math.Max(0, Dot(face.Normal, light));
        bool isBack = face.Normal.Z <= 0;
        if (isBack)
        {
            // lit from the same lamp, as if seen through
            lit = Math.Max(0, Dot((-face.Normal.X, -face.Normal.Y, -face.Normal.Z), light)) * 0.7;
        }

        var shade = Shade(solid.Color, lit);
        var lighter = Shade(solid.Color, Math.Min(1, lit + 0.18));
        var darker = Shade(solid.Color, Math.Max(0, lit - 0.14));
        byte fillAlpha = (byte)(alpha ? (isBack ? 110 : 200) : 255);
        var edge = Shade(solid.Color, 0);
        edge = ((byte)(edge.R * 0.75), (byte)(edge.G * 0.75), (byte)(edge.B * 0.75));
        styles.AppendLine(string.Format(invariant,
            "    <ShapeStyle IsFilled=\"true\" Color=\"{0}\" StrokeWidth=\"1\" Dash=\"Solid\" Name=\"{1}\">",
            Hex(edge, (byte)(alpha ? 160 : 255)), styleName));
        styles.AppendLine("      <Fill>");
        styles.AppendLine("        <LinearGradientBrush StartPoint=\"0,0\" EndPoint=\"1,1\">");
        styles.AppendLine("          <GradientStop Offset=\"0\" Color=\"" + Hex(lighter, fillAlpha) + "\" />");
        styles.AppendLine("          <GradientStop Offset=\"1\" Color=\"" + Hex(darker, fillAlpha) + "\" />");
        styles.AppendLine("        </LinearGradientBrush>");
        styles.AppendLine("      </Fill>");
        styles.AppendLine("    </ShapeStyle>");

        figures.AppendLine("    <Polygon Name=\"" + styleName + "\" Style=\"" + styleName + "\">");
        foreach (int vertex in face.Vertices.Concat(new[] { face.Vertices[0] }))
        {
            figures.AppendLine("      <Dependency Name=\"" + solid.Name + vertex + "\" />");
        }

        figures.AppendLine("    </Polygon>");
        faceIndex++;
    }

    // the name under the solid: a point label on a hidden point, centered by a pixel offset
    // (about 9.5 px per letter at this size)
    double textWidth = solid.Name.Length * 9.5;
    figures.AppendLine(string.Format(invariant,
        "    <PointByCoordinates Name=\"{0}\" Visible=\"false\" Style=\"DependentPointStyle\" X=\"{1:0.####}\" Y=\"{2:0.####}\" />",
        solid.Name, cx, cy - LabelDrop));
    figures.AppendLine(string.Format(invariant,
        "    <PointLabel Name=\"{0}Name\" IsHitTestVisible=\"false\" Style=\"Name\" OffsetX=\"{1:0.#}\" OffsetY=\"6\" ShowName=\"true\" ShowCoordinates=\"false\">",
        solid.Name, -textWidth / 2));
    figures.AppendLine("      <Dependency Name=\"" + solid.Name + "\" />");
    figures.AppendLine("    </PointLabel>");
}

var text = new StringBuilder();
text.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
text.AppendLine("<Drawing Version=\"1\">");
text.AppendLine("  <Viewport Left=\"-4.5\" Top=\"3.2\" Right=\"4.5\" Bottom=\"-3.4\" Grid=\"false\" />");
text.AppendLine("  <Styles>");
text.AppendLine("    <PointStyle Size=\"10\" Fill=\"#FFFFFF64\" IsFilled=\"true\" Color=\"#64000000\" StrokeWidth=\"1\" Dash=\"Solid\" Name=\"FreePoint\" />");
text.AppendLine("    <PointStyle Size=\"10\" Fill=\"#FF7CE38B\" IsFilled=\"true\" Color=\"#64000000\" StrokeWidth=\"1\" Dash=\"Solid\" Name=\"PointOnFigure\" />");
text.AppendLine("    <PointStyle Size=\"8\" Fill=\"#FF6FD3F7\" IsFilled=\"true\" Color=\"#64000000\" StrokeWidth=\"1\" Dash=\"Solid\" Name=\"IntersectionPoint\" />");
text.AppendLine("    <PointStyle Size=\"8\" Fill=\"#FFFFB45A\" IsFilled=\"true\" Color=\"#64000000\" StrokeWidth=\"1\" Dash=\"Solid\" Name=\"Midpoint\" />");
text.AppendLine("    <PointStyle Size=\"8\" Fill=\"#FFD0D0D0\" IsFilled=\"true\" Color=\"#64000000\" StrokeWidth=\"1\" Dash=\"Solid\" Name=\"DependentPointStyle\" />");
text.AppendLine("    <TextStyle FontSize=\"17\" Color=\"#FF2B3038\" FontFamily=\"Segoe UI\" Bold=\"false\" Italic=\"false\" Underline=\"false\" Name=\"Name\" />");
text.Append(styles);
text.AppendLine("    <TextStyle FontSize=\"30\" Color=\"#FF1F4E8C\" FontFamily=\"Segoe UI\" Bold=\"true\" Italic=\"false\" Underline=\"false\" Name=\"GalleryTitle\" />");
text.AppendLine("    <TextStyle FontSize=\"16\" Color=\"#FF2B3038\" FontFamily=\"Segoe UI\" Bold=\"false\" Italic=\"false\" Underline=\"false\" Name=\"GalleryText\" />");
text.AppendLine("    <LineStyle Color=\"#FFE0362B\" StrokeWidth=\"2.5\" Name=\"GalleryLocus\" />");
text.AppendLine("  </Styles>");
text.AppendLine("  <Figures>");
text.Append(figures);
text.AppendLine("    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"5 Platonic Solids\" DecimalsToShow=\"2\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"78.75\" WrapWidth=\"400\" Backdrop=\"true\" />");
text.AppendLine("    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"A Platonic solid is a shape whose faces are all the same regular polygon, with the same number of them meeting at every corner. There are exactly five - the ancient Greeks proved that no sixth one is possible.\\n\\nTetrahedron: 4 triangles. Cube: 6 squares. Octahedron: 8 triangles. Dodecahedron: 12 pentagons. Icosahedron: 20 triangles.\\n\\nRole-playing dice come in exactly these five shapes.\" DecimalsToShow=\"2\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"126.75\" WrapWidth=\"400\" Backdrop=\"true\" />");
text.AppendLine("  </Figures>");
text.Append("</Drawing>");
File.WriteAllText(args[0], text.ToString().Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(false));
Console.WriteLine("wrote " + args[0] + (alpha ? " (translucent)" : ""));
return 0;

static (double X, double Y, double Z) V(double x, double y, double z) => (x, y, z);

// standing on a face, a corner up
static (double X, double Y, double Z)[] Tetrahedron()
{
    double radius = 2 * Math.Sqrt(2) / 3;
    var result = new List<(double, double, double)> { (0, 1, 0) };
    foreach (double degrees in new[] { 90, 210, 330 })
    {
        double angle = degrees * Math.PI / 180;
        result.Add((radius * Math.Cos(angle), -1.0 / 3, radius * Math.Sin(angle)));
    }

    return result.ToArray();
}

static (double X, double Y, double Z)[] Signs3(double x, double y, double z)
{
    var result = new List<(double, double, double)>();
    foreach (int sx in new[] { -1, 1 })
    {
        foreach (int sy in new[] { -1, 1 })
        {
            foreach (int sz in new[] { -1, 1 })
            {
                result.Add((sx * x, sy * y, sz * z));
            }
        }
    }

    return result.ToArray();
}

static (double X, double Y, double Z)[] Dodecahedron()
{
    double phi = (1 + Math.Sqrt(5)) / 2;
    var result = new List<(double, double, double)>(Signs3(1, 1, 1));
    foreach (int a in new[] { -1, 1 })
    {
        foreach (int b in new[] { -1, 1 })
        {
            result.Add((0, a / phi, b * phi));
            result.Add((a / phi, b * phi, 0));
            result.Add((a * phi, 0, b / phi));
        }
    }

    return result.ToArray();
}

static (double X, double Y, double Z)[] Icosahedron()
{
    double phi = (1 + Math.Sqrt(5)) / 2;
    var result = new List<(double, double, double)>();
    foreach (int a in new[] { -1, 1 })
    {
        foreach (int b in new[] { -1, 1 })
        {
            result.Add((0, a, b * phi));
            result.Add((a, b * phi, 0));
            result.Add((a * phi, 0, b));
        }
    }

    return result.ToArray();
}

// degrees about x (nod), y (turn), z (roll)
static (double X, double Y, double Z) Tilt(double x, double y, double z) => (x, y, z);

static (double X, double Y, double Z) Rotate((double X, double Y, double Z) v, (double X, double Y, double Z) tilt)
{
    double ay = tilt.Y * Math.PI / 180;
    var turned = (X: v.X * Math.Cos(ay) + v.Z * Math.Sin(ay), Y: v.Y, Z: -v.X * Math.Sin(ay) + v.Z * Math.Cos(ay));
    double ax = tilt.X * Math.PI / 180;
    var nodded = (X: turned.X, Y: turned.Y * Math.Cos(ax) - turned.Z * Math.Sin(ax), Z: turned.Y * Math.Sin(ax) + turned.Z * Math.Cos(ax));
    double az = tilt.Z * Math.PI / 180;
    return (nodded.X * Math.Cos(az) - nodded.Y * Math.Sin(az), nodded.X * Math.Sin(az) + nodded.Y * Math.Cos(az), nodded.Z);
}

static double Dot((double X, double Y, double Z) a, (double X, double Y, double Z) b) => a.X * b.X + a.Y * b.Y + a.Z * b.Z;

static (double X, double Y, double Z) Cross((double X, double Y, double Z) a, (double X, double Y, double Z) b)
    => (a.Y * b.Z - a.Z * b.Y, a.Z * b.X - a.X * b.Z, a.X * b.Y - a.Y * b.X);

static (double X, double Y, double Z) Normalize((double X, double Y, double Z) v)
{
    double length = Math.Sqrt(Dot(v, v));
    return (v.X / length, v.Y / length, v.Z / length);
}

static (double X, double Y, double Z) Minus((double X, double Y, double Z) a, (double X, double Y, double Z) b) => (a.X - b.X, a.Y - b.Y, a.Z - b.Z);

/// The faces of a convex solid: every plane through three vertices that has all the others
/// on one side, with the vertices on it in order around it and the normal pointing out
static List<Face> Faces((double X, double Y, double Z)[] vertices)
{
    var faces = new List<Face>();
    var seen = new HashSet<string>();
    int n = vertices.Length;
    for (int i = 0; i < n; i++)
    {
        for (int j = i + 1; j < n; j++)
        {
            for (int k = j + 1; k < n; k++)
            {
                var normal = Cross(Minus(vertices[j], vertices[i]), Minus(vertices[k], vertices[i]));
                if (Dot(normal, normal) < 1e-9)
                {
                    continue;
                }

                normal = Normalize(normal);
                double offset = Dot(normal, vertices[i]);
                var distances = vertices.Select(v => Dot(normal, v) - offset).ToArray();
                if (distances.All(d => d < 1e-6) || distances.All(d => d > -1e-6))
                {
                    var onPlane = Enumerable.Range(0, n).Where(index => Math.Abs(distances[index]) < 1e-6).ToArray();
                    string key = string.Join(",", onPlane);
                    if (!seen.Add(key))
                    {
                        continue;
                    }

                    // outward: away from the middle of the solid, which is the origin
                    if (offset < 0)
                    {
                        normal = (-normal.X, -normal.Y, -normal.Z);
                    }

                    // around the face: by angle in the face's own plane
                    var centroid = (X: onPlane.Average(index => vertices[index].X), Y: onPlane.Average(index => vertices[index].Y), Z: onPlane.Average(index => vertices[index].Z));
                    var axisU = Normalize(Minus(vertices[onPlane[0]], centroid));
                    var axisV = Cross(normal, axisU);
                    var ordered = onPlane
                        .OrderBy(index => Math.Atan2(Dot(Minus(vertices[index], centroid), axisV), Dot(Minus(vertices[index], centroid), axisU)))
                        .ToArray();
                    faces.Add(new Face(ordered, normal, centroid.Z));
                }
            }
        }
    }

    return faces;
}

// the base color lit from 0 (shade, but a gentle one: these are pastels) to 1 (full light)
static (byte R, byte G, byte B) Shade((int R, int G, int B) color, double lit)
{
    double keep = 0.68 + 0.32 * lit;
    double toWhite = 0.4 * Math.Max(0, lit - 0.55);
    return (Channel(color.R, keep, toWhite), Channel(color.G, keep, toWhite), Channel(color.B, keep, toWhite));
}

static byte Channel(int value, double keep, double toWhite)
{
    double result = value * keep;
    result += (255 - result) * toWhite;
    return (byte)Math.Clamp(Math.Round(result), 0, 255);
}

static string Hex((byte R, byte G, byte B) color, byte alpha) => "#" + alpha.ToString("X2") + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");

record Solid(string Name, (int R, int G, int B) Color, (double X, double Y, double Z) Tilt, (double X, double Y, double Z)[] Vertices);

record Face(int[] Vertices, (double X, double Y, double Z) Normal, double Depth);
