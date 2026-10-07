#:property Nullable=disable
#:property PublishAot=false

// moonjelly - writes the "Moon Jelly" gallery drawing: a jellyfish of Bezier paths that swims
// as the pearl is dragged around its ring. A hidden point Beat holds what the pearl's angle
// says: how far the bell is squeezed (0 to 1) and the phase of the stroke. Every anchor is a
// point by coordinates over Beat, so the whole stroke is in their expressions:
//
// - the bell: one closed path through 18 anchors (the dome, then the lobes and notches of the
//   hem), its handles automatic, its tension tied to a hidden label that grows with the
//   squeeze, so the curve pulls tighter;
// - five tentacles: open paths hanging from the lobe tips, each anchor swaying like the one
//   above it a moment later;
// - four rings on the bell: each a filled path with a smaller path for a hole, the hole's
//   anchors dilated toward the ring's middle by a hidden label;
// - a sheen: some of the dome's anchors dilated toward the bell's middle.
//
// The light beams, the specks of marine snow, the bubbles and the dial around the pearl are
// decorations: IsHitTestVisible="false", so a press on one pans the view.
//
//   dotnet tools/moonjelly.cs -- <out.lgf>
//
// The gallery's file is what LiveGeometry.Desktop.exe --rewrite makes of the output.

using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

if (args.Length < 1)
{
    Console.WriteLine("usage: moonjelly <out.lgf>");
    return 1;
}

// the phase of the stroke the drawing opens at; the pearl sits at the angle pi - phase on its
// ring, so that turning it clockwise moves the phase forward
const double StartPhase = 1.75;

// where the jellyfish hangs: a hidden free point the bell and the tentacles are built around
// (the bell bobs and sways a little about it)
const double HomeX = 0;
const double HomeY = 0.45;

// the dial at the lower right: the ring the pearl goes around
const double HubX = 2.3;
const double HubY = -1.6;
const double RingRadius = 0.7;

// the glint on the dial's upper left and the arrow over its upper right, in degrees: the arc of
// the arrow goes clockwise from ArrowFrom to its head at ArrowTo
const double GlintFrom = 118;
const double GlintTo = 158;
const double ArrowFrom = 66;
const double ArrowTo = 22;

// the bell's tension (the scale of its automatic handles): RelaxedTension when the bell is
// open, growing by SqueezeTension at the full squeeze
const double RelaxedTension = 0.9;
const double SqueezeTension = 0.35;

// how much later in the stroke the crown of the dome squeezes than its margin
const double CrownLag = 0.6;

// a soft light around the bell: the outline of its inside, drawn under the rim
const string GlowColor = "#4080E8FF";
const double GlowWidth = 12;

// the tentacles: the anchors under the lobe tip each hangs from, and the width of the strokes;
// the middle one is half a pixel wider, and its anchors show with the bell's when Dots is on
const int TentacleAnchors = 5;
const double TentacleWidth = 3;
const int MiddleTentacle = 3;

// the scene the drawing is fitted to, and the viewport the file opens with
const double SceneLeft = -1.9;
const double SceneTop = 2.1;
const double SceneRight = 3.2;
const double SceneBottom = -2.65;

// decorations are paper: a press on one goes through to the view, which it pans
const string PassThrough = " IsHitTestVisible=\"false\"";

// gradient directions, by the brush's start and end in the box of the shape
const string Down = "StartPoint=\"0,0\" EndPoint=\"0,1\"";
const string Diagonal = "StartPoint=\"0,0\" EndPoint=\"1,1\"";

// the outline of a shape that only has a fill
const string NoOutline = "Color=\"#00000000\"";

var styles = new StringBuilder();
var figures = new StringBuilder();

// every figure written so far, to find the ones an expression names
var names = new HashSet<string>();

// Styles (their order is the style picker's). The caption is light on the dark sea under
// either theme
styles.AppendLine("    <TextStyle Name=\"GalleryTitle\" FontSize=\"30\" Color=\"#FFFFE6F7\" FontFamily=\"Segoe UI\" Bold=\"true\">");
styles.AppendLine("      <Dark Color=\"#FFFFE6F7\" />");
styles.AppendLine("    </TextStyle>");
styles.AppendLine("    <TextStyle Name=\"GalleryText\" FontSize=\"16\" Color=\"#FFDDEBFF\" FontFamily=\"Segoe UI\">");
styles.AppendLine("      <Dark Color=\"#FFDDEBFF\" />");
styles.AppendLine("    </TextStyle>");
styles.AppendLine("    <TextStyle Name=\"SpinText\" FontSize=\"14\" Color=\"#FFFFC6EA\" FontFamily=\"Segoe UI\" />");
ShapeStyle("Beam", NoOutline, Gradient(Down, ("#14BFE4FF", 0), ("#09BFE4FF", 0.5), ("#00BFE4FF", 1)));

// the bell's inside, its outline the glow: the same curve as the rim over it, so the glow
// follows every change of the curve
ShapeStyle(
    "BellFill",
    $"Color=\"{GlowColor}\" StrokeWidth=\"{Format(GlowWidth)}\"",
    Gradient("StartPoint=\"0.3,0\" EndPoint=\"0.7,1\"", ("#E8FFD6F1", 0), ("#C8F59BDD", 0.5), ("#A88E68F0", 1)));
LineStyle("BellRim", "#F0FFE6F8", width: 2.5);
ShapeStyle("GonadFill", NoOutline, Gradient(Diagonal, ("#C0FF8BD8", 0), ("#C0B060F0", 1)));
LineStyle("GonadEdge", "#70FFD4F2", width: 1.2);
LineStyle("GonadInside", "#60FFD4F2", width: 1);
ShapeStyle("SheenFill", NoOutline, Gradient(Diagonal, ("#A0FFFFFF", 0), ("#10FFFFFF", 1)));
LineStyle("SheenEdge", "#00FFFFFF", width: 1);

// pink, peach, lilac, sky, rose
string[] tentacleColors = { "#FFFFB3E4", "#FFFFC2A8", "#FFD8B8FF", "#FF9EE7FF", "#FFFFA8DA" };
for (int tentacle = 1; tentacle <= tentacleColors.Length; tentacle++)
{
    double width = tentacle == MiddleTentacle ? TentacleWidth + 0.5 : TentacleWidth;
    LineStyle($"Tentacle{tentacle}", tentacleColors[tentacle - 1], width);
}

ShapeStyle("Dial", "Color=\"#A0FFB3E4\" StrokeWidth=\"2\" Dash=\"Dot\"", Gradient(Diagonal, ("#1CBFE4FF", 0), ("#05BFE4FF", 1)));
LineStyle("GlintLine", "#60FFFFFF", width: 3);
LineStyle("DialArrow", "#C0FFB3E4", width: 2);
ShapeStyle("DialArrowHead", NoOutline, Gradient(Diagonal, ("#E0FFB3E4", 0), ("#E0FFB3E4", 1)));
styles.AppendLine("    <PointStyle Name=\"Pearl\" Size=\"20\" Fill=\"#FFFFF4FB\" Color=\"#FFFF8FD1\" StrokeWidth=\"3\" />");
styles.AppendLine("    <PointStyle Name=\"Anchor\" Size=\"8\" Fill=\"#FFFFFFFF\" Color=\"#FF7A3FC8\" StrokeWidth=\"1.5\" />");
styles.AppendLine("    <PointStyle Name=\"JellyHandle\" Size=\"8\" Fill=\"#FFFFF4FB\" Color=\"#FFFF8FD1\" StrokeWidth=\"2\" />");
styles.AppendLine("    <PointStyle Name=\"Bubble\" Size=\"9\" Fill=\"#18BFE4FF\" Color=\"#80BFE4FF\" StrokeWidth=\"1.2\" />");
styles.AppendLine("    <PointStyle Name=\"BubbleSmall\" Size=\"6\" Fill=\"#18BFE4FF\" Color=\"#70BFE4FF\" StrokeWidth=\"1\" />");
styles.AppendLine("    <PointStyle Name=\"BubbleBig\" Size=\"13\" Fill=\"#18BFE4FF\" Color=\"#90BFE4FF\" StrokeWidth=\"1.4\" />");
styles.AppendLine("    <PointStyle Name=\"Plankton\" Size=\"4\" Fill=\"#80FFFFFF\" Color=\"#00FFFFFF\" />");
styles.AppendLine("    <PointStyle Name=\"PlanktonSmall\" Size=\"3\" Fill=\"#55FFFFFF\" Color=\"#00FFFFFF\" />");

// Light from the surface: slanting beams that fade out downward, each a wide faint one with a
// narrow one over it for a soft edge. A beam goes from its width at the top, above the scene,
// to a wider one at BottomY
var beams = new (double TopX, double TopWidth, double BottomX, double BottomWidth, double BottomY)[]
{
    (-1.3, 0.5, 0.2, 1.1, -1.3),
    (-0.45, 0.22, 1.0, 0.5, -0.6),
    (0.85, 0.6, 2.45, 1.2, -1.9),
};

for (int beam = 0; beam < beams.Length; beam++)
{
    var (topX, topWidth, bottomX, bottomWidth, bottomY) = beams[beam];
    foreach (var (layer, scale) in new[] { ("", 1.0), ("Core", 0.5) })
    {
        string name = $"Beam{beam + 1}{layer}";
        string[] corners = { name + "A", name + "B", name + "C", name + "D" };
        Point(corners[0], Format(topX - topWidth * scale / 2), "2.75");
        Point(corners[1], Format(topX + topWidth * scale / 2), "2.75");
        Point(corners[2], Format(bottomX + bottomWidth * scale / 2), Format(bottomY));
        Point(corners[3], Format(bottomX - bottomWidth * scale / 2), Format(bottomY));
        DecorationFigure("Polygon", name, "Beam", corners);
    }
}

// Marine snow: specks that stay put
var specks = new (double X, double Y, bool Big)[]
{
    (-1.7, 1.85, true), (-0.95, 2.1, false), (0.65, 2.15, true), (1.55, 1.75, false), (2.8, 2.0, true),
    (2.45, 0.35, false), (-1.88, 1.15, false), (-1.5, -1.05, true), (-0.95, -2.6, false), (1.05, -2.65, true),
    (2.9, -0.5, true), (1.9, 0.95, true), (-1.85, -1.95, false), (0.45, -2.8, false), (2.1, 2.2, false),
};

for (int index = 0; index < specks.Length; index++)
{
    var speck = specks[index];
    DecorationPoint($"Speck{index + 1}", Format(speck.X), Format(speck.Y), speck.Big ? "Plankton" : "PlanktonSmall");
}

// The dial: a bubble with a dotted rim (the ring the pearl goes around), a glint on its upper
// left, and an arrow saying which way swims forward
Point("Hub", Format(HubX), Format(HubY));
Point("RingEdge", Format(HubX + RingRadius), Format(HubY));
DecorationFigure("Circle", "Ring", "Dial", new[] { "Hub", "RingEdge" });
double glintRadius = RingRadius - 0.11;
Point("GlintStart", Format(HubX + glintRadius * Math.Cos(Radians(GlintFrom))), Format(HubY + glintRadius * Math.Sin(Radians(GlintFrom))));
Point("GlintEnd", Format(HubX + glintRadius * Math.Cos(Radians(GlintTo))), Format(HubY + glintRadius * Math.Sin(Radians(GlintTo))));
DecorationFigure("CircleArc", "Glint", "GlintLine", new[] { "Hub", "GlintStart", "GlintEnd" });

// an arc goes counterclockwise from its start to its end; it stops a little short of the head,
// so that its stroke doesn't show past the tip
double arrowRadius = RingRadius + 0.17;
Point("ArrowStart", Format(HubX + arrowRadius * Math.Cos(Radians(ArrowTo + 4))), Format(HubY + arrowRadius * Math.Sin(Radians(ArrowTo + 4))));
Point("ArrowEnd", Format(HubX + arrowRadius * Math.Cos(Radians(ArrowFrom))), Format(HubY + arrowRadius * Math.Sin(Radians(ArrowFrom))));
DecorationFigure("CircleArc", "Arrow", "DialArrow", new[] { "Hub", "ArrowStart", "ArrowEnd" });

// the head at the clockwise end, pointing along the ring
const double HeadLength = 0.11;
const double HeadHalfWidth = 0.05;
double tipAngle = Radians(ArrowTo);
double tipX = HubX + arrowRadius * Math.Cos(tipAngle);
double tipY = HubY + arrowRadius * Math.Sin(tipAngle);
double alongX = Math.Sin(tipAngle);
double alongY = -Math.Cos(tipAngle);
double outwardX = Math.Cos(tipAngle);
double outwardY = Math.Sin(tipAngle);
Point("ArrowTip", Format(tipX + 0.02 * alongX), Format(tipY + 0.02 * alongY));
Point("ArrowOut", Format(tipX - HeadLength * alongX + HeadHalfWidth * outwardX), Format(tipY - HeadLength * alongY + HeadHalfWidth * outwardY));
Point("ArrowIn", Format(tipX - HeadLength * alongX - HeadHalfWidth * outwardX), Format(tipY - HeadLength * alongY - HeadHalfWidth * outwardY));
DecorationFigure("Polygon", "ArrowHead", "DialArrowHead", new[] { "ArrowTip", "ArrowOut", "ArrowIn" });

// the pearl, which drives it all, and its hint under the ring
double pearlAngle = Math.PI - StartPhase;
string pearlX = Format(HubX + RingRadius * Math.Cos(pearlAngle));
string pearlY = Format(HubY + RingRadius * Math.Sin(pearlAngle));
figures.AppendLine($"    <PointOnFigure Name=\"Pearl\" Style=\"Pearl\" X=\"{pearlX}\" Y=\"{pearlY}\" Parameter=\"{Format(pearlAngle)}\">");
figures.AppendLine("      <Dependency Name=\"Ring\" />");
figures.AppendLine("    </PointOnFigure>");
names.Add("Pearl");
figures.AppendLine($"    <Label Name=\"SpinMe\" Style=\"SpinText\"{PassThrough} Text=\"spin me!\" X=\"{Format(HubX - 0.19)}\" Y=\"{Format(HubY - RingRadius - 0.09)}\" />");
figures.AppendLine($"    <FreePoint Name=\"Home\" Visible=\"false\" X=\"{Format(HomeX)}\" Y=\"{Format(HomeY)}\" />");
names.Add("Home");

// The variables. Beat = (the squeeze from 0 to 1, the phase); only sin and cos of the phase
// are used, so the jump of the pearl's angle at a full turn doesn't show. Body = the bell's
// middle, bobbing and swaying a little; Size = half the width of the dome at its margin, and
// its height; Skirt = half the width of the hem, and where it is in y
Point("Beat", "(1 + cos(xang(Hub, Pearl))) / 2", "pi - xang(Hub, Pearl)");
Point("Body", "Home.X + 0.04 * sin(Beat.Y + 0.6)", "Home.Y + 0.3 * sin(Beat.Y - 1.2)");
Point("Size", "1.35 - 0.3 * Beat.X", "0.95 + 0.25 * Beat.X");
Point("Skirt", "Size.X * (1 - 0.3 * Beat.X)", "Body.Y - 0.1 - 0.12 * Beat.X");

// Bubbles that jiggle as the jelly swims, each a moment after the one before
var bubbles = new (double X, double Y, string Style)[]
{
    (2.78, 0.05, "BubbleSmall"), (2.9, 0.5, "Bubble"), (2.74, 1.0, "BubbleSmall"), (2.86, 1.55, "BubbleBig"),
    (-1.62, -0.55, "BubbleSmall"), (-1.72, -0.15, "Bubble"), (-1.6, 0.45, "BubbleBig"),
};

for (int index = 0; index < bubbles.Length; index++)
{
    var bubble = bubbles[index];
    string beat = index == 0 ? "Beat.Y" : $"Beat.Y + {Format(1.3 * index)}";
    string x = $"{Format(bubble.X)} + 0.035 * sin({beat})";
    string y = $"{Format(bubble.Y)} + 0.025 * cos({beat})";
    DecorationPoint($"Bubble{index + 1}", x, y, bubble.Style);
}

// The dome, counterclockwise from the right corner of the margin to the left one. The squeeze
// runs up it from the margin to the crown: each anchor has its own moment of the stroke, a
// little later the higher it is
var bellAnchors = new List<string>();
for (int step = 0; step <= 8; step++)
{
    double angle = step * Math.PI / 8;
    double cosine = Math.Cos(angle);
    double sine = Math.Sin(angle);
    string x;
    string y;
    if (Math.Abs(sine) < 1e-9)
    {
        x = $"Body.X + {Format(cosine)} * Size.X";
        y = "Body.Y";
    }
    else
    {
        string beat = $"cos(Beat.Y - {Format(CrownLag * sine)})";
        x = Math.Abs(cosine) < 1e-9 ? "Body.X" : $"Body.X + {Format(cosine)} * (1.2 + 0.15 * {beat})";
        y = $"Body.Y + {Format(sine)} * (1.075 - 0.125 * {beat})";
    }

    Point($"Dome{step}", x, y, "Anchor");
    bellAnchors.Add($"Dome{step}");
}

// The hem, left to right: lobe tips (even) and the notches between them (odd), as deep as a
// third of their spacing, which shrinks as the hem pulls in, so the lobes stay round
double[] hemPlaces = { -0.8, -0.6, -0.4, -0.2, 0, 0.2, 0.4, 0.6, 0.8 };
for (int index = 0; index < hemPlaces.Length; index++)
{
    double across = hemPlaces[index];
    string x = across == 0 ? "Body.X" : $"Body.X + {Format(across)} * Skirt.X";
    string y = index % 2 == 0 ? "Skirt.Y - 0.045 * Skirt.X" : "Skirt.Y + 0.03 * Skirt.X";
    Point($"Rim{index}", x, y, "Anchor");
    bellAnchors.Add($"Rim{index}");
}

// The tentacles, each hanging from a lobe tip, its anchors its reach apart going down. Each
// anchor sways like the one above it, 0.9 later in the stroke and a little wider, and they
// fan out; each tentacle sways 0.6 ahead of the one on its left. They bob with the bell, a
// moment later the lower the anchor, hang lower in the squeeze and stretch with the stroke
int[] tentacleTips = { 0, 2, 4, 6, 8 };
double[] reach = { 0.34, 0.42, 0.46, 0.43, 0.36 };
for (int tentacle = 1; tentacle <= tentacleTips.Length; tentacle++)
{
    double across = hemPlaces[tentacleTips[tentacle - 1]];
    var anchors = new List<string> { $"Rim{tentacleTips[tentacle - 1]}" };
    for (int anchor = 1; anchor <= TentacleAnchors; anchor++)
    {
        string top = across == 0 ? "Home.X" : $"Home.X + {Format(across)} * Skirt.X";
        string fan = across == 0 ? "" : $" + {Format(0.05 * anchor * across)}";
        double lag = 0.9 * anchor - 0.6 * tentacle;
        string wave = Math.Abs(lag) < 1e-9 ? "sin(Beat.Y)" : $"sin(Beat.Y - {Format(lag)})";
        string x = $"{top}{fan} + {Format(0.055 * anchor)} * {wave}";
        string bob = $"0.3 * sin(Beat.Y - {Format(1.2 + 0.5 * anchor)})";
        string stretch = $"(1 + 0.12 * cos(Beat.Y - {Format(0.9 * anchor)}))";
        string y = $"Home.Y + {bob} - 0.15 - 0.15 * Beat.X - {Format(reach[tentacle - 1] * anchor)} * {stretch}";
        string name = $"Tent{tentacle}_{anchor}";
        Point(name, x, y, "Anchor");
        anchors.Add(name);
    }

    BezierPath(
        $"Tentacle{tentacle}",
        anchors,
        closed: false,
        fill: null,
        sides: $"Tentacle{tentacle}");
}

// The bell, its tension a hidden label that grows with the squeeze
HiddenLabel("Squeeze", $"{Format(RelaxedTension)} + {Format(SqueezeTension)} * Beat.X");
BezierPath(
    "Bell",
    bellAnchors,
    closed: true,
    fill: "BellFill",
    sides: "BellRim",
    tensionFrom: "Squeeze");

// Four rings on the bell, in a clover seen from the side: the back one higher and smaller,
// the front one lower and bigger. Gonad = half the width and height of a ring, which breathes;
// Hollow = how big its hole is, the ring's points dilated toward its middle, smaller as the
// bell squeezes
Point("Gonad", "0.15 + 0.03 * Beat.X", "0.085 + 0.055 * Beat.X");
HiddenLabel("Hollow", "0.56 - 0.28 * Beat.X");
var spots = new (string Name, double X, double Y, double Scale)[]
{
    ("Back", 0, 0.62, 0.8),
    ("Left", -0.44, 0.42, 1),
    ("Right", 0.44, 0.42, 1),
    ("Front", 0, 0.22, 1.2),
};

foreach (var spot in spots)
{
    string middle = $"Gonad{spot.Name}";
    Point(middle, spot.X == 0 ? "Body.X" : $"Body.X + {Format(spot.X)} * Size.X", $"Body.Y + {Format(spot.Y)} * Size.Y");
    Point($"{middle}E", $"{middle}.X + {Format(spot.Scale)} * Gonad.X", $"{middle}.Y");
    Point($"{middle}N", $"{middle}.X", $"{middle}.Y + {Format(spot.Scale)} * Gonad.Y");
    Point($"{middle}W", $"{middle}.X - {Format(spot.Scale)} * Gonad.X", $"{middle}.Y");
    Point($"{middle}S", $"{middle}.X", $"{middle}.Y - {Format(spot.Scale)} * Gonad.Y");
    foreach (var side in "ENWS")
    {
        Dilated($"{middle}In{side}", $"{middle}{side}", middle, "Hollow");
    }

    BezierPath(
        $"{middle}In",
        new[] { $"{middle}InE", $"{middle}InN", $"{middle}InW", $"{middle}InS" },
        closed: true,
        fill: null,
        sides: "GonadInside");
    BezierPath(
        $"{middle}Ring",
        new[] { $"{middle}E", $"{middle}N", $"{middle}W", $"{middle}S" },
        closed: true,
        fill: "GonadFill",
        sides: "GonadEdge",
        holes: new[] { $"{middle}In" });
}

// A glassy highlight on the upper left of the dome: a crescent of the dome's own anchors from
// the crown to the left, dilated toward the bell's middle, by SheenOuter along its outer edge,
// SheenInner along its inner one and SheenTip at its ends
foreach (var (name, value) in new[] { ("SheenOuter", 0.92), ("SheenTip", 0.87), ("SheenInner", 0.8) })
{
    figures.AppendLine($"    <Number Name=\"{name}\" Value=\"{Format(value)}\" />");
    names.Add(name);
}

Dilated("Sheen0", "Dome4", "Body", "SheenTip");
Dilated("Sheen1", "Dome5", "Body", "SheenOuter");
Dilated("Sheen2", "Dome6", "Body", "SheenOuter");
Dilated("Sheen3", "Dome7", "Body", "SheenTip");
Dilated("Sheen4", "Dome6", "Body", "SheenInner");
Dilated("Sheen5", "Dome5", "Body", "SheenInner");
BezierPath(
    "Sheen",
    new[] { "Sheen0", "Sheen1", "Sheen2", "Sheen3", "Sheen4", "Sheen5" },
    closed: true,
    fill: "SheenFill",
    sides: "SheenEdge",
    tension: 1.6);

// The Dots box: shows the anchors of the bell and of the middle tentacle, hidden at first
figures.AppendLine("    <ShowHideControl Name=\"Dots\" Style=\"GalleryText\" Show=\"false\" Text=\"Dots\" X=\"-1.75\" Y=\"-2.3\">");
WriteDependencies(bellAnchors.Concat(Enumerable.Range(start: 1, count: TentacleAnchors).Select(anchor => $"Tent{MiddleTentacle}_{anchor}")));
figures.AppendLine("    </ShowHideControl>");

// The caption; the file says a line break as \n, and the quotes as the attribute needs them
const string Title = "Moon Jelly";
const string Description = "Drag the pearl around its ring, and the jellyfish swims. Waves run down its tentacles (drag backward, and they run up): each dot swings like the one above it, a moment later.\\n\\nThe bell is one smooth curve through 18 dots. As it squeezes, the curve pulls tighter and the holes in its rings shrink. Check \"Dots,\" click a dot, then drag the pearl: its two pink handles, which steer the curve, turn by themselves.";
string description = Description.Replace("\"", "&quot;");
figures.AppendLine($"    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"{Title}\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
figures.AppendLine($"    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"{description}\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\" />");

var drawing = new StringBuilder();
drawing.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
drawing.AppendLine("<Drawing Version=\"1\">");
string scene = $"Left=\"{Format(SceneLeft)}\" Top=\"{Format(SceneTop)}\" Right=\"{Format(SceneRight)}\" Bottom=\"{Format(SceneBottom)}\"";
drawing.AppendLine($"  <Viewport {scene}>");

// the sea, darker the deeper
drawing.AppendLine("    <Background>");
foreach (var line in Gradient(Down, ("#FF12467F", 0), ("#FF0A2754", 0.55), ("#FF050F24", 1)))
{
    drawing.AppendLine("      " + line);
}

drawing.AppendLine("    </Background>");
drawing.AppendLine("  </Viewport>");
drawing.AppendLine($"  <Scene {scene} />");
drawing.AppendLine("  <Styles>");
drawing.Append(styles);
drawing.AppendLine("  </Styles>");
drawing.AppendLine("  <Figures>");
drawing.Append(figures);
drawing.AppendLine("  </Figures>");
drawing.Append("</Drawing>");
File.WriteAllText(args[0], drawing.ToString().Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
Console.WriteLine("wrote " + args[0]);
return 0;

// numbers as the file has them: invariant, at most six decimals
static string Format(double value) => Math.Round(value, digits: 6).ToString("0.######", CultureInfo.InvariantCulture);

static double Radians(double degrees) => degrees * Math.PI / 180;

// "Body.X + -0.5 * Size.X" reads "Body.X - 0.5 * Size.X"
static string Tidy(string expression) => expression.Replace("+ -", "- ").Replace("- -", "+ ");

static string StyleAttribute(string style) => style != null ? $" Style=\"{style}\"" : "";

// a linear gradient brush, as the lines of its element
static List<string> Gradient(string direction, params (string Color, double Offset)[] stops)
{
    var lines = new List<string> { $"<LinearGradientBrush {direction}>" };
    foreach (var stop in stops)
    {
        lines.Add($"  <GradientStop Color=\"{stop.Color}\" Offset=\"{Format(stop.Offset)}\" />");
    }

    lines.Add("</LinearGradientBrush>");
    return lines;
}

// a shape style: its outline (the attributes as the file has them) and a gradient for a fill
void ShapeStyle(string name, string outline, List<string> gradient)
{
    styles.AppendLine($"    <ShapeStyle Name=\"{name}\" {outline}>");
    styles.AppendLine("      <Fill>");
    foreach (var line in gradient)
    {
        styles.AppendLine("        " + line);
    }

    styles.AppendLine("      </Fill>");
    styles.AppendLine("    </ShapeStyle>");
}

void LineStyle(string name, string color, double width)
{
    styles.AppendLine($"    <LineStyle Name=\"{name}\" Color=\"{color}\" StrokeWidth=\"{Format(width)}\" />");
}

// the figures written so far that the expressions name, in the order they first name them
List<string> NamesIn(params string[] expressions)
{
    var found = new List<string>();
    foreach (var expression in expressions)
    {
        foreach (Match match in Regex.Matches(expression, @"[A-Za-z_][A-Za-z0-9_]*"))
        {
            if (names.Contains(match.Value) && !found.Contains(match.Value))
            {
                found.Add(match.Value);
            }
        }
    }

    return found;
}

void WriteDependencies(IEnumerable<string> dependencies)
{
    foreach (var dependency in dependencies)
    {
        figures.AppendLine($"      <Dependency Name=\"{dependency}\" />");
    }
}

// a point by coordinates, which depends on the figures its expressions name
void PointByCoordinates(string name, string attributes, string x, string y)
{
    x = Tidy(x);
    y = Tidy(y);
    var dependencies = NamesIn(x, y);
    string element = $"    <PointByCoordinates Name=\"{name}\"{attributes} X=\"{x}\" Y=\"{y}\"";
    if (dependencies.Count == 0)
    {
        figures.AppendLine(element + " />");
    }
    else
    {
        figures.AppendLine(element + ">");
        WriteDependencies(dependencies);
        figures.AppendLine("    </PointByCoordinates>");
    }

    names.Add(name);
}

// a hidden point by coordinates: a corner, a variable or an anchor (which the Dots box shows)
void Point(string name, string x, string y, string style = null)
{
    PointByCoordinates(name, " Visible=\"false\"" + StyleAttribute(style), x, y);
}

// a point by coordinates that shows, as a decoration
void DecorationPoint(string name, string x, string y, string style)
{
    PointByCoordinates(name, StyleAttribute(style) + PassThrough, x, y);
}

// a figure of the scenery, built on the points given
void DecorationFigure(string kind, string name, string style, string[] dependencies)
{
    figures.AppendLine($"    <{kind} Name=\"{name}\" Style=\"{style}\"{PassThrough}>");
    WriteDependencies(dependencies);
    figures.AppendLine($"    </{kind}>");
    names.Add(name);
}

// a hidden point: the source dilated about the center by the factor (a figure that says a number)
void Dilated(string name, string source, string center, string factor)
{
    figures.AppendLine($"    <DilatedPoint Name=\"{name}\" Visible=\"false\">");
    WriteDependencies(new[] { source, center, factor });
    figures.AppendLine("    </DilatedPoint>");
    names.Add(name);
}

// a hidden label that says the value of an expression, to tie a path's tension or a dilation to
void HiddenLabel(string name, string expression)
{
    figures.AppendLine($"    <Label Name=\"{name}\" Visible=\"false\" Text=\"[{expression}]\" X=\"-1.9\" Y=\"2.2\">");
    WriteDependencies(NamesIn(expression));
    figures.AppendLine("    </Label>");
    names.Add(name);
}

// a Bezier path through the anchors, every handle automatic ("a"), its sides in their own style
// and its inside in the fill style, or not filled. The holes and a tension tied to a figure
// follow the anchors among the dependencies, the tension naming its place there
void BezierPath(
    string name,
    IList<string> anchors,
    bool closed,
    string fill,
    string sides,
    IList<string> holes = null,
    double? tension = null,
    string tensionFrom = null)
{
    var dependencies = new List<string>(anchors);
    if (holes != null)
    {
        dependencies.AddRange(holes);
    }

    string tensionAttribute = tension != null ? $" Tension=\"{Format(tension.Value)}\"" : "";
    if (tensionFrom != null)
    {
        dependencies.Add(tensionFrom);
        tensionAttribute = $" Tension=\"#{dependencies.Count - 1}\"";
    }

    // a piece per anchor, the closing one written whether the path is closed or not
    string path = string.Join(" ", Enumerable.Repeat("C a a", anchors.Count));
    string closedAttribute = closed ? " Closed=\"true\"" : "";
    string filledAttribute = fill != null ? " Filled=\"true\"" : "";
    figures.AppendLine($"    <BezierPath Name=\"{name}\"{StyleAttribute(fill)}{closedAttribute}{filledAttribute}{tensionAttribute} Path=\"{path}\">");
    figures.AppendLine($"      <Sides Style=\"{sides}\" />");
    figures.AppendLine("      <Handles Style=\"JellyHandle\" />");
    WriteDependencies(dependencies);
    figures.AppendLine("    </BezierPath>");
    names.Add(name);
}
