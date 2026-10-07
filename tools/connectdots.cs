#:property Nullable=disable
#:property PublishAot=false

// connectdots - writes the "Connect the Dots" gallery drawing: three translucent blobs, closed
// Bezier paths with automatic handles through the same seven free points (the dots), each
// smoothed by its own rule - pink by Catmull-Rom, orange by a natural spline, blue by Hobby's -
// and each with a hole through four more free points (the diamonds), smoothed by the same
// rule. A dashed polygon joins the dots by straight lines. A bead on a wire under them sets the
// tension of all six paths, puffy at the left end and tight at the right. A check box per color
// hides its blob, and the Hint box shows a dashed circle and the text of the hint.
//
//   dotnet tools/connectdots.cs -- <out.lgf>
//
// The file is written the way the app saves it (the order of the attributes, X="0" Y="0" on
// the hidden label that works out the tension, the bead's X from its rounded parameter), so a
// --rewrite of it changes nothing.

using System.Globalization;
using System.Text;

if (args.Length < 1)
{
    Console.WriteLine("usage: connectdots <out.lgf>");
    return 1;
}

// the dots the blobs go through, counterclockwise from the right
var dots = new (double X, double Y)[]
{
    (2.6, 0.9), (1.0, 2.0), (-0.9, 1.9), (-2.2, 0.4), (-1.5, -1.6), (0.4, -1.0), (2.1, -1.6),
};

// the diamonds the holes go through (points named Pip), counterclockwise from the right
var pips = new (double X, double Y)[]
{
    (1.1, 0.2), (0.35, 0.85), (-0.45, 0.45), (0.2, -0.35),
};

// the wire the bead slides on, under the blobs. The tension of every path runs from
// TensionLow at the left end (Hobby's rule takes no less) to TensionHigh at the right, evenly
// in its logarithm, so that each step of the bead multiplies it by the same factor
const double WireLeft = -1.75;
const double WireRight = 2.25;
const double WireY = -2.3;
const double TensionLow = 0.75;
const double TensionHigh = 3;

// half the length of the tick across the wire where the tension is 1, which is where the bead
// starts
const double MarkReach = 0.06;

// the words "puffy" and "tight" at the ends of the wire. A label is placed by its upper left
// corner and sized in pixels, so "puffy" starts PuffyWidth (its width at a desktop window's
// zoom) and WordGap to the left of the wire, and "tight" WordGap to the right of it. Both tops
// are WordRise above the wire
const double WordGap = 0.15;
const double PuffyWidth = 0.27;
const double WordRise = 0.1;

// the dashed circle of the hint: spread evenly around it, the dots give the three outer edges
// the same shape
const double HintX = 0.3;
const double HintY = 0.25;
const double HintRadius = 2.2;

// the check boxes at the upper left, one under the other: the first BoxMargin pixels from the
// window's corner, each next one BoxSpacing pixels lower
const int BoxMargin = 16;
const int BoxSpacing = 24;

// the view, and the scene the gallery fits beside or above the caption: the dots, the wire and
// its words, with room for the blobs to bulge out
const double SceneLeft = -2.9;
const double SceneTop = 2.8;
const double SceneRight = 3.65;
const double SceneBottom = -2.55;

// the numbers of the file are written with at most this many decimals
const int Decimals = 6;

// the caption and the hint, as they read (LabelText writes them as the file has them)
const string Title = "Connect the Dots";
const string Description = "Three see-through blobs share seven dots, but each bends between them by its own rule. Drag a dot and each blob shows its handles, the hidden arms that steer its curve. On both sides, pink stays near the dashed line, orange swings out, blue bulges widest. Uncheck a color to hide its blob.\n\nThe diamonds cut a hole. Slide the bead to make every curve puffy or tight. With the bead on its mark, can you make the three outer edges agree?";
const string Hint = "Pink is a Catmull-Rom spline: each dot's handles come from its two neighbors alone. Orange is a natural spline: all its handles are worked out together, so one dot nudges the whole loop. Blue is John Hobby's rule, made for drawing letters: it keeps the bend as even as it can. Spread the dots evenly around the dashed circle and the outer edges agree. The hole never quite does: four diamonds are too few, and the more dots, the closer the rules come. Dragged an arm? That blob stops following its rule there, until you press Undo.";

// the three blobs. Each fill is its ink at one side of the blob's box, fading out across it,
// each from another side: pink from the left, orange from the bottom, blue from the right and
// again from the left, where its lobes bulge out; where all three overlap the inks are thin,
// or they would stack up to gray
var hugger = new Blob(
    Name: "Hugger",
    Word: "Pink",
    Smoothing: "CatmullRom",
    FillStart: "0,0.3",
    FillEnd: "1,0.7",
    FillStops: [("#98FF4FC8", 0), ("#48FF4FC8", 0.33), ("#14FF4FC8", 0.6), ("#00FF4FC8", 0.78)],
    RimColor: "#FFE8338F",
    HandleColor: "#FFFF5FD2",
    WordColor: "#FFC2185B");
var bouncer = new Blob(
    Name: "Bouncer",
    Word: "Orange",
    Smoothing: "NaturalSpline",
    FillStart: "0.5,1",
    FillEnd: "0.5,0",
    FillStops: [("#90FFD84D", 0), ("#48FFD84D", 0.3), ("#10FFD84D", 0.55), ("#00FFD84D", 0.7)],
    RimColor: "#FFF29100",
    HandleColor: "#FFFFB300",
    WordColor: "#FFB85F00");
var roundy = new Blob(
    Name: "Roundy",
    Word: "Blue",
    Smoothing: "Hobby",
    FillStart: "1,0.5",
    FillEnd: "0,0.5",
    FillStops: [("#8000BFFF", 0), ("#3C00BFFF", 0.22), ("#1000BFFF", 0.45), ("#1000BFFF", 0.8), ("#6800BFFF", 1)],
    RimColor: "#FF0A9BE0",
    HandleColor: "#FF35D6FF",
    WordColor: "#FF0072B0");

// the order of the styles and the check boxes
var blobs = new[] { hugger, bouncer, roundy };

// the order the blobs are painted in, the last on top
var paintOrder = new[] { bouncer, roundy, hugger };

double wireLength = WireRight - WireLeft;

// how much the logarithm of the tension grows per unit of wire: the Squeeze label works the
// tension out as TensionLow * exp(tensionGrowthPerUnit * (Bead.X - WireL.X))
double tensionGrowthPerUnit = Math.Log(TensionHigh / TensionLow) / wireLength;

// the tick where the tension is 1, and the bead on it. The bead's parameter is written rounded
// to Decimals, and a point on a figure is where its parameter says: its X is worked out from
// the rounded parameter, as the app does when it reads the file
double restParameter = Math.Log(1 / TensionLow) / tensionGrowthPerUnit / wireLength;
double markX = WireLeft + wireLength * restParameter;
double beadParameter = Math.Round(restParameter, Decimals);
double beadX = WireLeft + wireLength * beadParameter;

var text = new StringBuilder();
Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
Write("<Drawing Version=\"1\" Creator=\"LiveGeometry.App\">");

// the view and the scene, with the same bounds, and a paper of the drawing's own, light under
// every theme
string sceneBounds = $"Left=\"{Format(SceneLeft)}\" Top=\"{Format(SceneTop)}\" Right=\"{Format(SceneRight)}\" Bottom=\"{Format(SceneBottom)}\"";
Write($"  <Viewport {sceneBounds}>");
Write("    <Background>");
Write("      <LinearGradientBrush StartPoint=\"0,0\" EndPoint=\"0,1\">");
Write("        <GradientStop Color=\"#FFFFFFFF\" Offset=\"0\" />");
Write("        <GradientStop Color=\"#FFEFF4FF\" Offset=\"1\" />");
Write("      </LinearGradientBrush>");
Write("    </Background>");
Write("  </Viewport>");
Write($"  <Scene {sceneBounds} />");

// the styles, written out as the file has them: colors are #AARRGGBB, sizes and stroke widths
// pixels
Write("  <Styles>");
TextStyle("GalleryTitle", size: 30, color: "#FF1B2A4A", bold: true);
TextStyle("GalleryText", size: 16, color: "#FF2D3557", bold: false);
TextStyle("WireText", size: 15, color: "#FF5A6380", bold: false);

// a blob's fill has no outline: its sides are parts of the path, in the blob's Rim style. The
// gradient goes between two points of the path's box, (0, 0) its upper left corner and (1, 1)
// the lower right
foreach (var blob in blobs)
{
    Write($"    <ShapeStyle Name=\"{blob.Name}\" Color=\"#00000000\">");
    Write("      <Fill>");
    Write($"        <LinearGradientBrush StartPoint=\"{blob.FillStart}\" EndPoint=\"{blob.FillEnd}\">");
    foreach (var (color, offset) in blob.FillStops)
    {
        Write($"          <GradientStop Color=\"{color}\" Offset=\"{Format(offset)}\" />");
    }

    Write("        </LinearGradientBrush>");
    Write("      </Fill>");
    Write("    </ShapeStyle>");
}

// the sides of a blob and of its hole, and the handles that show next to a dot or a diamond
// that is selected or dragged
foreach (var blob in blobs)
{
    Write($"    <LineStyle Name=\"{blob.Name}Rim\" Color=\"{blob.RimColor}\" StrokeWidth=\"2\" />");
    Write($"    <PointStyle Name=\"{blob.Name}Handle\" Size=\"8\" Fill=\"{blob.HandleColor}\" Color=\"#FFFFFFFF\" StrokeWidth=\"1.5\" />");
}

Write("    <ShapeStyle Name=\"Band\" IsFilled=\"false\" Color=\"#80303A5A\" StrokeWidth=\"1.5\" Dash=\"Dash\" />");
Write("    <PointStyle Name=\"Dot\" Size=\"14\" Fill=\"#FF26215C\" Color=\"#FFFFFFFF\" StrokeWidth=\"2\" />");
Write("    <PointStyle Name=\"HoleDot\" Shape=\"Diamond\" Size=\"13\" Fill=\"#FF26215C\" Color=\"#FFFFFFFF\" StrokeWidth=\"2\" />");
Write("    <LineStyle Name=\"Wire\" Color=\"#FF9AA0B8\" StrokeWidth=\"3\" />");
Write("    <LineStyle Name=\"WireMark\" Color=\"#FF9AA0B8\" StrokeWidth=\"2\" />");
Write("    <PointStyle Name=\"Bead\" Size=\"18\" Fill=\"#FFFFFFFF\" Color=\"#FF26215C\" StrokeWidth=\"3\" />");
Write("    <LineStyle Name=\"HintCircle\" Color=\"#8026215C\" StrokeWidth=\"1.5\" Dash=\"Dash\" />");

// the words of the check boxes, each in its blob's color
foreach (var blob in blobs)
{
    TextStyle(blob.Name + "Text", size: 16, color: blob.WordColor, bold: true);
}

Write("  </Styles>");
Write("  <Figures>");

// the dots and the diamonds: dragging one reshapes all three blobs, or all three holes
for (int i = 0; i < dots.Length; i++)
{
    Write($"    <FreePoint Name=\"Dot{i + 1}\" Style=\"Dot\" X=\"{Format(dots[i].X)}\" Y=\"{Format(dots[i].Y)}\" />");
}

for (int i = 0; i < pips.Length; i++)
{
    Write($"    <FreePoint Name=\"Pip{i + 1}\" Style=\"HoleDot\" X=\"{Format(pips[i].X)}\" Y=\"{Format(pips[i].Y)}\" />");
}

// the wire, between points that can't be dragged, so that the wire stays where it is
Write($"    <PointByCoordinates Name=\"WireL\" Visible=\"false\" X=\"{Format(WireLeft)}\" Y=\"{Format(WireY)}\" />");
Write($"    <PointByCoordinates Name=\"WireR\" Visible=\"false\" X=\"{Format(WireRight)}\" Y=\"{Format(WireY)}\" />");
Write("    <Segment Name=\"Wire\" Style=\"Wire\">");
Write("      <Dependency Name=\"WireL\" />");
Write("      <Dependency Name=\"WireR\" />");
Write("    </Segment>");

// the tick across the wire where the tension is 1, and the bead, which starts on it
Write($"    <PointByCoordinates Name=\"MarkTop\" Visible=\"false\" X=\"{Format(markX)}\" Y=\"{Format(WireY + MarkReach)}\" />");
Write($"    <PointByCoordinates Name=\"MarkBottom\" Visible=\"false\" X=\"{Format(markX)}\" Y=\"{Format(WireY - MarkReach)}\" />");
Write("    <Segment Name=\"BeadMark\" Style=\"WireMark\">");
Write("      <Dependency Name=\"MarkTop\" />");
Write("      <Dependency Name=\"MarkBottom\" />");
Write("    </Segment>");
Write($"    <PointOnFigure Name=\"Bead\" Style=\"Bead\" X=\"{Format(beadX)}\" Y=\"{Format(WireY)}\" Parameter=\"{Format(beadParameter)}\">");
Write("      <Dependency Name=\"Wire\" />");
Write("    </PointOnFigure>");

// the tension, worked out from the bead's place: every path takes its tension from this hidden
// label
Write($"    <Label Name=\"Squeeze\" Visible=\"false\" Text=\"[{Format(TensionLow)} * exp({Format(tensionGrowthPerUnit)} * (Bead.X - WireL.X))]\" X=\"0\" Y=\"0\">");
Write("      <Dependency Name=\"Bead\" />");
Write("      <Dependency Name=\"WireL\" />");
Write("    </Label>");

// the words at the ends of the wire, a little above it
Write($"    <Label Name=\"Puffy\" Style=\"WireText\" Text=\"puffy\" X=\"{Format(WireLeft - PuffyWidth - WordGap)}\" Y=\"{Format(WireY + WordRise)}\" />");
Write($"    <Label Name=\"Tight\" Style=\"WireText\" Text=\"tight\" X=\"{Format(WireRight + WordGap)}\" Y=\"{Format(WireY + WordRise)}\" />");

// the dots joined by straight lines, dashed
Write("    <Polygon Name=\"Band\" Style=\"Band\">");
for (int i = 0; i < dots.Length; i++)
{
    Write($"      <Dependency Name=\"Dot{i + 1}\" />");
}

Write("    </Polygon>");

// the holes, before the blobs: a blob is built on its hole, and a figure comes after what it is
// built on. Path="C a a ..." is a piece per anchor with both handles automatic; Tension="#k"
// takes the tension from the k-th dependency (from 0), Squeeze after the anchors
foreach (var blob in paintOrder)
{
    Write($"    <BezierPath Name=\"Hole{blob.Name}\" Closed=\"true\"{SmoothingAttribute(blob)} Tension=\"#{pips.Length}\" Path=\"{AutomaticPath(pips.Length)}\">");
    Write($"      <Sides Style=\"{blob.Name}Rim\" />");
    Write($"      <Handles Style=\"{blob.Name}Handle\" />");
    for (int i = 0; i < pips.Length; i++)
    {
        Write($"      <Dependency Name=\"Pip{i + 1}\" />");
    }

    Write("      <Dependency Name=\"Squeeze\" />");
    Write("    </BezierPath>");
}

// the blobs: the dots, then the hole (a dependency after the anchors that neither the Path nor
// the Tension names is a hole), then Squeeze
foreach (var blob in paintOrder)
{
    Write($"    <BezierPath Name=\"{blob.Name}\" Style=\"{blob.Name}\" Closed=\"true\" Filled=\"true\"{SmoothingAttribute(blob)} Tension=\"#{dots.Length + 1}\" Path=\"{AutomaticPath(dots.Length)}\">");
    Write($"      <Sides Style=\"{blob.Name}Rim\" />");
    Write($"      <Handles Style=\"{blob.Name}Handle\" />");
    for (int i = 0; i < dots.Length; i++)
    {
        Write($"      <Dependency Name=\"Dot{i + 1}\" />");
    }

    Write($"      <Dependency Name=\"Hole{blob.Name}\" />");
    Write("      <Dependency Name=\"Squeeze\" />");
    Write("    </BezierPath>");
}

// the hint, hidden until its box is checked: the dashed circle, and the text, pinned to the
// window, which the gallery lays out under the caption (as it does the caption itself, last)
Write($"    <PointByCoordinates Name=\"HintCenter\" Visible=\"false\" X=\"{Format(HintX)}\" Y=\"{Format(HintY)}\" />");
Write($"    <PointByCoordinates Name=\"HintRim\" Visible=\"false\" X=\"{Format(HintX + HintRadius)}\" Y=\"{Format(HintY)}\" />");
Write("    <Circle Name=\"HintCircle\" Visible=\"false\" Style=\"HintCircle\">");
Write("      <Dependency Name=\"HintCenter\" />");
Write("      <Dependency Name=\"HintRim\" />");
Write("    </Circle>");
Write($"    <Label Name=\"Hint\" Visible=\"false\" Style=\"GalleryText\" Text=\"{LabelText(Hint)}\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"300\" WrapWidth=\"400\" Backdrop=\"true\" />");

// a check box per color, each hiding its blob and the sides of its hole (a path of its own),
// and the Hint box under them
int boxY = BoxMargin;
foreach (var blob in blobs)
{
    Write($"    <ShowHideControl Name=\"{blob.Word}Box\" Style=\"{blob.Name}Text\" Show=\"true\" Text=\"{blob.Word}\" Pin=\"TopLeft\" OffsetX=\"{BoxMargin}\" OffsetY=\"{boxY}\">");
    Write($"      <Dependency Name=\"{blob.Name}\" />");
    Write($"      <Dependency Name=\"Hole{blob.Name}\" />");
    Write("    </ShowHideControl>");
    boxY += BoxSpacing;
}

Write($"    <ShowHideControl Name=\"HintBox\" Style=\"GalleryText\" Show=\"false\" Text=\"Hint\" Pin=\"TopLeft\" OffsetX=\"{BoxMargin}\" OffsetY=\"{boxY}\">");
Write("      <Dependency Name=\"HintCircle\" />");
Write("      <Dependency Name=\"Hint\" />");
Write("    </ShowHideControl>");

// the caption, pinned to the window. When the drawing opens, the gallery lays out the title,
// the description and the hint anew for the window at hand (pin, offsets and wrap width), so
// the places written here for the three are only where they start
Write($"    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"{LabelText(Title)}\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write($"    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"{LabelText(Description)}\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("  </Figures>");

// the app writes no line break after the root's closing tag
text.Append("</Drawing>");

File.WriteAllText(args[0], text.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
Console.WriteLine("wrote " + args[0]);
return 0;

void Write(string line) => text.Append(line).Append("\r\n");

// GalleryTitle and GalleryText are default styles, which follow the theme: on this paper, light
// under every theme, every text keeps its Light color under Dark too
void TextStyle(string name, int size, string color, bool bold)
{
    Write($"    <TextStyle Name=\"{name}\" FontSize=\"{size}\" Color=\"{color}\" FontFamily=\"Segoe UI\"{(bold ? " Bold=\"true\"" : "")}>");
    Write($"      <Dark Color=\"{color}\" />");
    Write("    </TextStyle>");
}

// the file's numbers: at most Decimals decimals
static string Format(double value) => Math.Round(value, Decimals).ToString("0." + new string('#', Decimals), CultureInfo.InvariantCulture);

// a label's text as the file has it: a backslash doubled, a line break as the two characters
// \n, and what XML can't have in an attribute escaped
static string LabelText(string text)
{
    return text
        .Replace("\\", "\\\\")
        .Replace("\n", "\\n")
        .Replace("&", "&amp;")
        .Replace("<", "&lt;")
        .Replace(">", "&gt;")
        .Replace("\"", "&quot;");
}

// Hobby's rule is a path's default, which the file doesn't name
static string SmoothingAttribute(Blob blob) => blob.Smoothing == "Hobby" ? "" : $" Smoothing=\"{blob.Smoothing}\"";

// a piece per anchor, both of its handles automatic
static string AutomaticPath(int anchorCount) => string.Join(" ", Enumerable.Repeat("C a a", anchorCount));

// a blob: the name of its path, its hole (Hole + name) and its styles, the word on its check
// box, its smoothing rule, its fill (a gradient between two points of the path's box, and its
// stops) and the colors of its sides, its handles and its word
record Blob(
    string Name,
    string Word,
    string Smoothing,
    string FillStart,
    string FillEnd,
    (string Color, double Offset)[] FillStops,
    string RimColor,
    string HandleColor,
    string WordColor);
