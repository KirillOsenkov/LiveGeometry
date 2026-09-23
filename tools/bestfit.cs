#:property Nullable=disable
#:property PublishAot=false

// bestfit - writes the "Line of Best Fit" gallery drawing: twelve draggable points with an
// upward trend, the least-squares line through them as a live expression of the points, and
// the vertical gap from every point to the line.
//
//   dotnet tools/bestfit.cs -- <out.lgf>
//
// The arithmetic lives in three hidden points whose coordinates are numbers, not places:
// Mean = (mean x, mean y), Sums = (sum of x·y, sum of x²), Fit = (slope, intercept).

using System.Globalization;
using System.Text;

if (args.Length < 1)
{
    Console.WriteLine("usage: bestfit <out.lgf>");
    return 1;
}

var invariant = CultureInfo.InvariantCulture;

// scattered about y = 0.75 x + 1 like real measurements: x anywhere in the range, the
// error in y roughly normal (a sum of uniforms), so that a few points sit right on the
// line and a few well off it. A fixed seed, so that the drawing is the same every time.
// The x's are jittered rather than uniform, so that they spread over the range without
// two of them landing on top of each other.
var random = new Random(3);
double Noise() => (random.NextDouble() + random.NextDouble() + random.NextDouble() - 1.5) * 1.7;
var points = Enumerable.Range(0, 12)
    .Select(i => { double x = Math.Round(0.7 + i * 0.75 + (random.NextDouble() - 0.5) * 0.6, 1); return (X: x, Y: Math.Round(0.75 * x + 1 + Noise(), 1)); })
    .ToArray();
var names = Enumerable.Range(0, points.Length).Select(i => ((char)('A' + i)).ToString()).ToArray();
int n = points.Length;

string Sum(Func<string, string> term) => string.Join(" + ", names.Select(term));
string Dependencies() => string.Concat(names.Select(name => "      <Dependency Name=\"" + name + "\" />\r\n"));

var figures = new StringBuilder();
for (int i = 0; i < n; i++)
{
    figures.AppendLine(string.Format(invariant, "    <FreePoint Name=\"{0}\" Style=\"FreePoint\" X=\"{1}\" Y=\"{2}\" />", names[i], points[i].X, points[i].Y));
}

figures.AppendLine("    <PointByCoordinates Name=\"Mean\" Visible=\"false\" Style=\"DependentPointStyle\" X=\"(" + Sum(p => p + ".X") + ") / " + n + "\" Y=\"(" + Sum(p => p + ".Y") + ") / " + n + "\">");
figures.Append(Dependencies());
figures.AppendLine("    </PointByCoordinates>");
figures.AppendLine("    <PointByCoordinates Name=\"Sums\" Visible=\"false\" Style=\"DependentPointStyle\" X=\"" + Sum(p => p + ".X * " + p + ".Y") + "\" Y=\"" + Sum(p => p + ".X * " + p + ".X") + "\">");
figures.Append(Dependencies());
figures.AppendLine("    </PointByCoordinates>");

// slope = (sum xy - n mean x mean y) / (sum x² - n mean x²), intercept = mean y - slope mean x
figures.AppendLine("    <PointByCoordinates Name=\"Fit\" Visible=\"false\" Style=\"DependentPointStyle\" X=\"(Sums.X - " + n + " * Mean.X * Mean.Y) / (Sums.Y - " + n + " * Mean.X * Mean.X)\" Y=\"Mean.Y - (Sums.X - " + n + " * Mean.X * Mean.Y) / (Sums.Y - " + n + " * Mean.X * Mean.X) * Mean.X\">");
figures.AppendLine("      <Dependency Name=\"Mean\" />");
figures.AppendLine("      <Dependency Name=\"Sums\" />");
figures.AppendLine("    </PointByCoordinates>");

// the line: a whole line (it has no bounds, so zoom to fit doesn't count it) through its
// points at x = 0 and x = 10
foreach (var (name, x) in new[] { ("LineLeft", 0), ("LineRight", 10) })
{
    figures.AppendLine("    <PointByCoordinates Name=\"" + name + "\" Visible=\"false\" Style=\"DependentPointStyle\" X=\"" + x + "\" Y=\"Fit.X * " + x + " + Fit.Y\">");
    figures.AppendLine("      <Dependency Name=\"Fit\" />");
    figures.AppendLine("    </PointByCoordinates>");
}

figures.AppendLine("    <LineTwoPoints Name=\"Line\" Style=\"Line\">");
figures.AppendLine("      <Dependency Name=\"LineLeft\" />");
figures.AppendLine("      <Dependency Name=\"LineRight\" />");
figures.AppendLine("    </LineTwoPoints>");

// the gap from every point straight down (or up) to the line, and the square on it - the
// thing whose area the line keeps small; it goes to the right of the gap when the point is
// above the line and to the left when below, so it is a square either way
foreach (var name in names)
{
    string foot = "Fit.X * " + name + ".X + Fit.Y";
    string gap = "(" + name + ".Y - (" + foot + "))";
    figures.AppendLine("    <PointByCoordinates Name=\"" + name + "Foot\" Visible=\"false\" Style=\"DependentPointStyle\" X=\"" + name + ".X\" Y=\"" + foot + "\">");
    figures.AppendLine("      <Dependency Name=\"" + name + "\" />");
    figures.AppendLine("      <Dependency Name=\"Fit\" />");
    figures.AppendLine("    </PointByCoordinates>");
    figures.AppendLine("    <PointByCoordinates Name=\"" + name + "FootAcross\" Visible=\"false\" Style=\"DependentPointStyle\" X=\"" + name + ".X + " + gap + "\" Y=\"" + foot + "\">");
    figures.AppendLine("      <Dependency Name=\"" + name + "\" />");
    figures.AppendLine("      <Dependency Name=\"Fit\" />");
    figures.AppendLine("    </PointByCoordinates>");
    figures.AppendLine("    <PointByCoordinates Name=\"" + name + "Across\" Visible=\"false\" Style=\"DependentPointStyle\" X=\"" + name + ".X + " + gap + "\" Y=\"" + name + ".Y\">");
    figures.AppendLine("      <Dependency Name=\"" + name + "\" />");
    figures.AppendLine("      <Dependency Name=\"Fit\" />");
    figures.AppendLine("    </PointByCoordinates>");
    figures.AppendLine("    <Polygon Name=\"" + name + "Square\" Style=\"Square\">");
    figures.AppendLine("      <Dependency Name=\"" + name + "\" />");
    figures.AppendLine("      <Dependency Name=\"" + name + "Foot\" />");
    figures.AppendLine("      <Dependency Name=\"" + name + "FootAcross\" />");
    figures.AppendLine("      <Dependency Name=\"" + name + "Across\" />");
    figures.AppendLine("      <Dependency Name=\"" + name + "\" />");
    figures.AppendLine("    </Polygon>");
    figures.AppendLine("    <Segment Name=\"" + name + "Gap\" Style=\"Gap\">");
    figures.AppendLine("      <Dependency Name=\"" + name + "\" />");
    figures.AppendLine("      <Dependency Name=\"" + name + "Foot\" />");
    figures.AppendLine("    </Segment>");
}

string squares = Sum(p => "(" + p + ".Y - Fit.X * " + p + ".X - Fit.Y)^2");

var text = new StringBuilder();
text.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
text.AppendLine("<Drawing Version=\"1\">");
text.AppendLine("  <Viewport Left=\"-1\" Top=\"9\" Right=\"10.5\" Bottom=\"-1\" Grid=\"true\" Axes=\"true\" />");
text.AppendLine("  <Styles>");
text.AppendLine("    <PointStyle Size=\"10\" Fill=\"#FFFFFF64\" IsFilled=\"true\" Color=\"#64000000\" StrokeWidth=\"1\" Dash=\"Solid\" Name=\"FreePoint\" />");
text.AppendLine("    <PointStyle Size=\"10\" Fill=\"#FF7CE38B\" IsFilled=\"true\" Color=\"#64000000\" StrokeWidth=\"1\" Dash=\"Solid\" Name=\"PointOnFigure\" />");
text.AppendLine("    <PointStyle Size=\"8\" Fill=\"#FF6FD3F7\" IsFilled=\"true\" Color=\"#64000000\" StrokeWidth=\"1\" Dash=\"Solid\" Name=\"IntersectionPoint\" />");
text.AppendLine("    <PointStyle Size=\"8\" Fill=\"#FFFFB45A\" IsFilled=\"true\" Color=\"#64000000\" StrokeWidth=\"1\" Dash=\"Solid\" Name=\"Midpoint\" />");
text.AppendLine("    <PointStyle Size=\"8\" Fill=\"#FFD0D0D0\" IsFilled=\"true\" Color=\"#64000000\" StrokeWidth=\"1\" Dash=\"Solid\" Name=\"DependentPointStyle\" />");
text.AppendLine("    <LineStyle Color=\"#FF2F7BD6\" StrokeWidth=\"2.5\" Dash=\"Solid\" Name=\"Line\" />");
text.AppendLine("    <LineStyle Color=\"#FFD83B3B\" StrokeWidth=\"1.25\" Dash=\"Dash\" Name=\"Gap\" />");
text.AppendLine("    <ShapeStyle Fill=\"#38D83B3B\" IsFilled=\"true\" Color=\"#70D83B3B\" StrokeWidth=\"1\" Dash=\"Solid\" Name=\"Square\" />");
text.AppendLine("    <TextStyle FontSize=\"30\" Color=\"#FF1F4E8C\" FontFamily=\"Segoe UI\" Bold=\"true\" Italic=\"false\" Underline=\"false\" Name=\"GalleryTitle\" />");
text.AppendLine("    <TextStyle FontSize=\"16\" Color=\"#FF2B3038\" FontFamily=\"Segoe UI\" Bold=\"false\" Italic=\"false\" Underline=\"false\" Name=\"GalleryText\" />");
text.AppendLine("    <LineStyle Color=\"#FFE0362B\" StrokeWidth=\"2.5\" Name=\"GalleryLocus\" />");
text.AppendLine("  </Styles>");
text.AppendLine("  <Figures>");
text.Append(figures);
text.AppendLine("    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"Line of Best Fit\" DecimalsToShow=\"2\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"57.5\" WrapWidth=\"400\" Backdrop=\"true\" />");
text.AppendLine("    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"Twelve points that lie roughly along a line, like measurements from an experiment. The blue line fits them best: of all the lines you could draw, it makes the red squares, one on each point's gap to the line, as small as possible in total area. The gap is measured straight up or down - how far off the line's guess for y is - not the shortest distance to the line. Statisticians call it the least-squares line, or linear regression, and use it to read a trend out of scattered data and to predict what comes next.\\n\\ny = [Fit.X] · x + [Fit.Y]\\nTotal area of the squares: [" + squares + "]\\n\\nDrag any point and watch the line follow. Move one far away: how much does the line care?\" DecimalsToShow=\"2\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"105.5\" WrapWidth=\"400\" Backdrop=\"true\">");
text.Append(Dependencies());
text.AppendLine("      <Dependency Name=\"Fit\" />");
text.AppendLine("    </Label>");
text.AppendLine("  </Figures>");
text.Append("</Drawing>");
File.WriteAllText(args[0], text.ToString().Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(false));
Console.WriteLine("wrote " + args[0]);
return 0;
