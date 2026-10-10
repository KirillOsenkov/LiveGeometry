#:property Nullable=disable
#:property PublishAot=false

// designafont - writes the "Design a Font" gallery drawing: a toy font editor. Three big
// letters, t, h and e, are closed Bezier paths over free points, drawn between the guide lines
// of a font (ascender, x-height, baseline, descender) in a color each. Under them the sentence
// "the quick brown fox jumps over the lazy dog" is set in the same font, and every t, h and e
// in it is a dilated image of the big letter (its anchors and handles dilated towards a hidden
// center by the Shrink number, 1/8), so that dragging a point of the big letter reshapes the
// small ones with it. The other letters are static: closed paths over hidden points by
// coordinates, drawn by stroking a skeleton of lines and circular arcs with a round pen.
//
//   dotnet tools/designafont.cs -- <out.lgf>
//
// Then --rewrite the folder: the app lays the file out as it saves it.

using System.Globalization;
using System.Text;
using static Geo;

if (args.Length < 1)
{
    Console.WriteLine("usage: designafont <out.lgf>");
    return 1;
}

// ---- the font, in font units: 1000 to the em ----

// every letter sits in a box this wide (a monospaced font), the x-height is half an em
const double Advance = 600;
const double XHeight = 500;
const double Ascender = 740;
const double Descender = -220;

// the letters are drawn with a round pen this wide
const double PenWidth = 110;
const double Half = PenWidth / 2;

// where the pen's path runs so that the ink just reaches a line
const double LowBottom = Half;
const double LowTop = XHeight - Half;
const double HighTop = Ascender - Half;
const double DeepBottom = Descender + Half;

// the bowls of a, b, d, g, o, p, q: a circle in the middle of the box, as high as the x-height
const double BowlX = Advance / 2;
const double BowlY = XHeight / 2;
const double BowlRadius = XHeight / 2 - Half;

// the stems of h, n, u and the legs of the arches between them
const double LeftStem = 150;
const double RightStem = 450;
const double ArchRadius = (RightStem - LeftStem) / 2;

// the dot of i and j
const double DotY = 650;
const double DotRadius = 70;

// ---- the drawing, in plane units ----

// the sentence's em is one unit; the big letters are Magnify times that, and their images in
// the sentence are dilated by one over it
const double TextUnit = 0.001;
const double Magnify = 8;
const double BigUnit = TextUnit * Magnify;

const string BigLetters = "the";
string[] sentence = { "the quick brown fox", "jumps over the lazy dog" };

// the big letters stand on y = 0 with their boxes centered on x = 0; the sentence's lines are
// centered too, the first one's ascenders SentenceGap under the big descender line
const double SentenceGap = 0.9;
const double LineHeight = 1.3;

// the names of the guide lines: two at the right end of the lines ("ascender", "x-height"),
// two at the left ("baseline", "descender"), each starting LabelLift above its line, which
// puts the text's baseline on the line on a phone. A label is placed by its upper left corner
// and sized in pixels, so the ones at the right start RightLabelInset inside the big letters'
// right edge: just right of the e's top point, and about their width on a phone
const double LabelLift = 0.45;
const double RightLabelInset = 1.9;

const string Title = "Design a Font";
const string Description = "Every letter you read on a screen is a shape: an outline made of a few points and curves, filled with color. The big t, h and e are three such outlines. Drag their points to reshape them, and watch every t, h and e in the sentence change along: each small one is the big one, shrunk to one eighth and set into its place in the line.\n\nThat is how fonts work. A type designer draws each letter once, between guide lines like these, and your computer scales the outlines to any size wherever you type. The sentence has every letter of the alphabet, which is why font makers use it to try out a new design.";
const string Hint = "Click a point to see its two arms: drag an arm and the curve bends while the point stays put. Round points sit where the outline bends smoothly, square ones where it turns a corner. The small letters copy the arms as well.\n\nCan you shorten the h into an n? Give the t a longer foot? Make the e smile wider?";

// the big letters' hues: the outline, the points and the small copies of each
var hues = new Dictionary<char, Hue>
{
    ['t'] = new Hue("T", Light: "#FFEF5B3C", Dark: "#FFFF7D62"),
    ['h'] = new Hue("H", Light: "#FF2F8FE0", Dark: "#FF62AEFF"),
    ['e'] = new Hue("E", Light: "#FF9256E0", Dark: "#FFB38CFF"),
};

var alphabet = Alphabet();
var bigGlyphs = new Dictionary<char, Glyph>
{
    ['t'] = GlyphT(),
    ['h'] = GlyphH(),
    ['e'] = GlyphE(),
};

double bigLeft = -BigLetters.Length * Advance * BigUnit / 2;
double bigRight = -bigLeft;
double sentenceTop = Descender * BigUnit - SentenceGap;

var text = new StringBuilder();
Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
Write("<Drawing Version=\"1\" Creator=\"LiveGeometry.App\">");

// the view: the big letters with the names of their lines, and the sentence
double top = Ascender * BigUnit + LabelLift + 0.3;
double bottom = sentenceTop - Ascender * TextUnit - (sentence.Length - 1) * LineHeight + Descender * TextUnit - 0.3;
Write($"  <Viewport Left=\"{Format(bigLeft - 0.3)}\" Top=\"{Format(top)}\" Right=\"{Format(bigRight + 0.3)}\" Bottom=\"{Format(bottom)}\" />");

Write("  <Styles>");

// the sentence's ink, in the theme's text colors: a fill without an outline
Write("    <ShapeStyle Name=\"Ink\" Fill=\"#FF2A3140\" Color=\"#00000000\">");
Write("      <Dark Fill=\"#FFE4E8F0\" />");
Write("    </ShapeStyle>");
Write("    <LineStyle Name=\"NoLine\" Color=\"#00FFFFFF\" StrokeWidth=\"0.5\" />");

// per big letter: its see-through fill, its outline (the sides of its path), its points - round
// where the outline is smooth (at the default size, 10), square at a corner - its handles, and
// the solid ink of its copies
foreach (var hue in hues.Values)
{
    Write($"    <ShapeStyle Name=\"{hue.Name}Fill\" Fill=\"{Translucent(hue.Light)}\" Color=\"#00000000\">");
    Write($"      <Dark Fill=\"{Translucent(hue.Dark)}\" />");
    Write("    </ShapeStyle>");
    Write($"    <LineStyle Name=\"{hue.Name}Rim\" Color=\"{hue.Light}\" StrokeWidth=\"2\">");
    Write($"      <Dark Color=\"{hue.Dark}\" />");
    Write("    </LineStyle>");
    Write($"    <PointStyle Name=\"{hue.Name}Node\" Fill=\"#FFFFFFFF\" Color=\"{hue.Light}\" StrokeWidth=\"2.5\">");
    Write($"      <Dark Color=\"{hue.Dark}\" />");
    Write("    </PointStyle>");
    Write($"    <PointStyle Name=\"{hue.Name}Corner\" Shape=\"Square\" Size=\"9\" Fill=\"#FFFFFFFF\" Color=\"{hue.Light}\" StrokeWidth=\"2.5\">");
    Write($"      <Dark Color=\"{hue.Dark}\" />");
    Write("    </PointStyle>");
    Write($"    <PointStyle Name=\"{hue.Name}Handle\" Size=\"7\" Fill=\"{hue.Light}\" Color=\"#FFFFFFFF\" StrokeWidth=\"1.5\">");
    Write($"      <Dark Fill=\"{hue.Dark}\" />");
    Write("    </PointStyle>");
    Write($"    <ShapeStyle Name=\"{hue.Name}Ink\" Fill=\"{hue.Light}\" Color=\"#00000000\">");
    Write($"      <Dark Fill=\"{hue.Dark}\" />");
    Write("    </ShapeStyle>");
}

// the guide lines and their names
Write("    <LineStyle Name=\"Guide\" Color=\"#FFB4BAC8\" Dash=\"Dash\">");
Write("      <Dark Color=\"#FF5A6275\" />");
Write("    </LineStyle>");
Write("    <TextStyle Name=\"GuideText\" FontSize=\"12\" Color=\"#FF8A93A6\" FontFamily=\"Segoe UI\">");
Write("      <Dark Color=\"#FF8A93A6\" />");
Write("    </TextStyle>");
Write("  </Styles>");

Write("  <Figures>");

// the factor every copy is shrunk by
Write($"    <Number Name=\"Shrink\" Value=\"{Format(1 / Magnify)}\" />");

// the guide lines, through points that can't be dragged
foreach (var (name, height) in new[] { ("Ascender", Ascender), ("XHeight", XHeight), ("Baseline", 0.0), ("Descender", Descender) })
{
    double y = height * BigUnit;
    Write($"    <PointByCoordinates Name=\"{name}L\" Visible=\"false\" X=\"-1\" Y=\"{Format(y)}\" />");
    Write($"    <PointByCoordinates Name=\"{name}R\" Visible=\"false\" X=\"1\" Y=\"{Format(y)}\" />");
    Write($"    <LineTwoPoints Name=\"{name}Line\" Style=\"Guide\">");
    Write($"      <Dependency Name=\"{name}L\" />");
    Write($"      <Dependency Name=\"{name}R\" />");
    Write("    </LineTwoPoints>");
}

// the big letters: their points, free; the eye of the e, a path of its own that the e leaves
// out; and the paths
var bigOrigins = new Dictionary<char, P>();
var bigNames = new Dictionary<char, GlyphNames>();
for (int i = 0; i < BigLetters.Length; i++)
{
    char letter = BigLetters[i];
    var glyph = bigGlyphs[letter];
    var hue = hues[letter];
    var origin = new P(bigLeft + i * Advance * BigUnit, 0);
    bigOrigins[letter] = origin;

    var names = new GlyphNames(
        Path: "Glyph" + hue.Name,
        Anchors: Enumerable.Range(1, glyph.Outline.Anchors.Count).Select(j => hue.Name + j).ToList(),
        Hole: glyph.Hole == null ? null : "Eye" + hue.Name,
        HoleAnchors: glyph.Hole == null ? null : Enumerable.Range(1, glyph.Hole.Anchors.Count).Select(j => "Eye" + j).ToList());
    bigNames[letter] = names;

    WriteFreePoints(glyph.Outline, names.Anchors, origin, hue);
    if (glyph.Hole != null)
    {
        WriteFreePoints(glyph.Hole, names.HoleAnchors, origin, hue);
        WritePath(names.Hole, style: null, glyph.Hole, BigUnit, names.HoleAnchors, sides: hue.Name + "Rim", handles: hue.Name + "Handle", filled: false, hole: null);
    }

    WritePath(names.Path, hue.Name + "Fill", glyph.Outline, BigUnit, names.Anchors, sides: hue.Name + "Rim", handles: hue.Name + "Handle", filled: true, hole: names.Hole);
}

// the sentence: a static letter is paths over hidden points, a t, h or e the image of the big
// one under a dilation whose center puts the box of the big letter onto the letter's box
int letterIndex = 0;
var copyCounts = new Dictionary<char, int>();
for (int line = 0; line < sentence.Length; line++)
{
    string words = sentence[line];
    double left = -words.Length * Advance * TextUnit / 2;
    double baseline = sentenceTop - Ascender * TextUnit - line * LineHeight;
    for (int column = 0; column < words.Length; column++)
    {
        char letter = words[column];
        if (letter == ' ')
        {
            continue;
        }

        letterIndex++;
        var origin = new P(left + column * Advance * TextUnit, baseline);
        if (bigGlyphs.TryGetValue(letter, out var glyph))
        {
            copyCounts[letter] = copyCounts.GetValueOrDefault(letter) + 1;
            WriteCopy(hues[letter].Name + "Copy" + copyCounts[letter], glyph, bigNames[letter], bigOrigins[letter], origin, hues[letter]);
        }
        else
        {
            WriteStaticLetter("W" + letterIndex, alphabet[letter], origin);
        }
    }
}

// the names of the guide lines
WriteLabel("AscenderName", "ascender", bigRight - RightLabelInset, Ascender * BigUnit + LabelLift);
WriteLabel("XHeightName", "x-height", bigRight - RightLabelInset, XHeight * BigUnit + LabelLift);
WriteLabel("BaselineName", "baseline", bigLeft, LabelLift);
WriteLabel("DescenderName", "descender", bigLeft, Descender * BigUnit + LabelLift);

// the hint, hidden until its box is checked, and the caption, pinned to the window: the gallery
// lays out the title, the description and the hint anew for the window at hand
Write($"    <Label Name=\"Hint\" Visible=\"false\" Style=\"GalleryText\" Text=\"{LabelText(Hint)}\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"300\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("    <ShowHideControl Name=\"HintBox\" Style=\"GalleryText\" Show=\"false\" Text=\"Hint\" Pin=\"TopLeft\" OffsetX=\"16\" OffsetY=\"16\">");
Write("      <Dependency Name=\"Hint\" />");
Write("    </ShowHideControl>");
Write($"    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"{LabelText(Title)}\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write($"    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"{LabelText(Description)}\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("  </Figures>");

// the app writes no line break after the root's closing tag
text.Append("</Drawing>");

File.WriteAllText(args[0], text.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
Console.WriteLine("wrote " + args[0]);
return 0;

void Write(string line) => text.Append(line).Append("\r\n");

// the anchors of a big letter's contour, as free points in the letter's hue: round where the
// outline goes on smoothly, square at a corner
void WriteFreePoints(Contour contour, IList<string> names, P origin, Hue hue)
{
    for (int j = 0; j < contour.Anchors.Count; j++)
    {
        var at = origin + contour.Anchors[j].At * BigUnit;
        string style = hue.Name + (contour.IsSmooth(j) ? "Node" : "Corner");
        Write($"    <FreePoint Name=\"{names[j]}\" Style=\"{style}\" X=\"{Format(at.X)}\" Y=\"{Format(at.Y)}\" />");
    }
}

void WriteHiddenPoints(Contour contour, IList<string> names, P origin, double scale)
{
    for (int j = 0; j < contour.Anchors.Count; j++)
    {
        var at = origin + contour.Anchors[j].At * scale;
        Write($"    <PointByCoordinates Name=\"{names[j]}\" Visible=\"false\" X=\"{Format(at.X)}\" Y=\"{Format(at.Y)}\" />");
    }
}

// a closed path over the anchors named, its handles the contour's offsets scaled to the plane
void WritePath(
    string name,
    string style,
    Contour contour,
    double scale,
    IList<string> anchors,
    string sides,
    string handles,
    bool filled,
    string hole)
{
    WritePathElement(name, style, contour.PathText(scale), anchors, sides, handles, filled, hole);
}

void WritePathElement(
    string name,
    string style,
    string path,
    IEnumerable<string> dependencies,
    string sides,
    string handles,
    bool filled,
    string hole)
{
    var element = new StringBuilder();
    element.Append($"    <BezierPath Name=\"{name}\"");
    if (style != null)
    {
        element.Append($" Style=\"{style}\"");
    }

    element.Append(" Closed=\"true\"");
    if (filled)
    {
        element.Append(" Filled=\"true\"");
    }

    element.Append($" Path=\"{path}\">");
    Write(element.ToString());
    Write($"      <Sides Style=\"{sides}\" />");
    if (handles != null)
    {
        Write($"      <Handles Style=\"{handles}\" />");
    }

    foreach (var dependency in dependencies)
    {
        Write($"      <Dependency Name=\"{dependency}\" />");
    }

    if (hole != null)
    {
        Write($"      <Dependency Name=\"{hole}\" />");
    }

    Write("    </BezierPath>");
}

// a letter of the sentence that is not one of the big three: each piece a path over hidden
// points, in the sentence's ink, a bowl with its hole before it
void WriteStaticLetter(string name, Glyph[] pieces, P origin)
{
    for (int p = 0; p < pieces.Length; p++)
    {
        var piece = pieces[p];
        string pieceName = name + (char)('a' + p);
        string holeName = null;
        if (piece.Hole != null)
        {
            holeName = pieceName + "Hole";
            var holeAnchors = Enumerable.Range(1, piece.Hole.Anchors.Count).Select(j => holeName + "P" + j).ToList();
            WriteHiddenPoints(piece.Hole, holeAnchors, origin, TextUnit);
            WritePath(holeName, style: null, piece.Hole, TextUnit, holeAnchors, sides: "NoLine", handles: null, filled: false, hole: null);
        }

        var anchors = Enumerable.Range(1, piece.Outline.Anchors.Count).Select(j => pieceName + "P" + j).ToList();
        WriteHiddenPoints(piece.Outline, anchors, origin, TextUnit);
        WritePath(pieceName, "Ink", piece.Outline, TextUnit, anchors, sides: "NoLine", handles: null, filled: true, hole: holeName);
    }
}

// a small copy of a big letter: the image of its path under the dilation from a hidden center
// by Shrink. The center is where the big letter's box corner goes onto the small one's:
// center + Shrink * (big - center) = small. The images of the anchors and of the handles are
// dilated points (hidden, auxiliary: they go with the copy), and the copy's handles are those
// points, so that an arm dragged on the big letter moves on the small one too
void WriteCopy(string name, Glyph glyph, GlyphNames source, P bigOrigin, P origin, Hue hue)
{
    double shrink = 1 / Magnify;
    var center = (origin - bigOrigin * shrink) * (1 / (1 - shrink));
    Write($"    <PointByCoordinates Name=\"{name}Center\" Visible=\"false\" X=\"{Format(center.X)}\" Y=\"{Format(center.Y)}\" />");

    string holeName = null;
    if (glyph.Hole != null)
    {
        holeName = name + "Eye";
        WriteImage(holeName, glyph.Hole, source.Hole, source.HoleAnchors, name + "Center", style: null, filled: false, hole: null);
    }

    WriteImage(name, glyph.Outline, source.Path, source.Anchors, name + "Center", hue.Name + "Ink", filled: true, hole: holeName);
}

void WriteImage(
    string name,
    Contour contour,
    string sourcePath,
    IList<string> sourceAnchors,
    string center,
    string style,
    bool filled,
    string hole)
{
    int count = contour.Anchors.Count;
    var dependencies = new List<string>();
    for (int j = 0; j < count; j++)
    {
        string anchor = name + "A" + (j + 1);
        WriteDilatedPoint(anchor, $"<Dependency Name=\"{sourceAnchors[j]}\" />", center);
        dependencies.Add(anchor);
    }

    // the handles' points, in1, out1, in2, out2..., as the app lists them
    for (int j = 0; j < count; j++)
    {
        string inName = name + "I" + (j + 1);
        string outName = name + "O" + (j + 1);
        WriteDilatedPoint(inName, $"<Dependency Name=\"{sourcePath}\" Part=\"In{j + 1}\" />", center);
        WriteDilatedPoint(outName, $"<Dependency Name=\"{sourcePath}\" Part=\"Out{j + 1}\" />", center);
        dependencies.Add(inName);
        dependencies.Add(outName);
    }

    // a piece per anchor, from its out handle to the next anchor's in handle, by their places
    // among the dependencies
    var pieces = new List<string>();
    for (int j = 0; j < count; j++)
    {
        int outIndex = count + 2 * j + 1;
        int inIndex = count + 2 * ((j + 1) % count);
        pieces.Add($"C #{outIndex} #{inIndex}");
    }

    WritePathElement(name, style, string.Join(" ", pieces), dependencies, sides: "NoLine", handles: null, filled, hole);
}

void WriteDilatedPoint(string name, string source, string center)
{
    Write($"    <DilatedPoint Name=\"{name}\" Visible=\"false\" Auxiliary=\"true\">");
    Write("      " + source);
    Write($"      <Dependency Name=\"{center}\" />");
    Write("      <Dependency Name=\"Shrink\" />");
    Write("    </DilatedPoint>");
}

void WriteLabel(string name, string caption, double x, double y)
{
    Write($"    <Label Name=\"{name}\" Style=\"GuideText\" Text=\"{caption}\" X=\"{Format(x)}\" Y=\"{Format(y)}\" />");
}

// a hue at a fifth of its opacity, for the fill of a big letter
static string Translucent(string color) => "#38" + color.Substring(3);

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

// ---- the big letters, traced by hand as one outline each: the union of the pen's strokes ----

// t: a stem from 650 down into a foot that turns right at the bottom, and a bar across at the
// x-height; counterclockwise from where the stem's left edge meets the bar's lower edge
static Glyph GlyphT()
{
    const double StemX = 300;
    const double StemTop = 650 - Half;
    const double FootRadius = 140;
    const double FootY = LowBottom + FootRadius;
    const double FootEnd = 480;
    const double BarY = LowTop;
    const double BarLeft = 150;
    const double BarRight = 450;

    var c = new Contour();
    var foot = new P(StemX + FootRadius, FootY);
    c.Start(new P(StemX - Half, BarY - Half));
    c.LineTo(new P(StemX - Half, FootY));
    c.ArcTo(foot, FootRadius + Half, 180, 270);
    c.LineTo(new P(FootEnd, LowBottom - Half));
    c.ArcTo(new P(FootEnd, LowBottom), Half, 270, 450, wholeCap: true);
    c.LineTo(new P(foot.X, LowBottom + Half));
    c.ArcTo(foot, FootRadius - Half, 270, 180);
    c.LineTo(new P(StemX + Half, BarY - Half));
    c.LineTo(new P(BarRight, BarY - Half));
    c.ArcTo(new P(BarRight, BarY), Half, 270, 450, wholeCap: true);
    c.LineTo(new P(StemX + Half, BarY + Half));
    c.LineTo(new P(StemX + Half, StemTop));
    c.ArcTo(new P(StemX, StemTop), Half, 0, 180, wholeCap: true);
    c.LineTo(new P(StemX - Half, BarY + Half));
    c.LineTo(new P(BarLeft, BarY + Half));
    c.ArcTo(new P(BarLeft, BarY), Half, 90, 270, wholeCap: true);
    c.LineTo(new P(StemX - Half, BarY - Half));
    c.Close();
    return new Glyph(c, Hole: null);
}

// h: an ascending stem on the left and an arch from it over to a leg on the right;
// counterclockwise from the bottom of the stem
static Glyph GlyphH()
{
    double archY = LowTop - ArchRadius;
    var arch = new P((LeftStem + RightStem) / 2, archY);

    // where the arch's outer edge meets the stem's right edge
    double meet = Math.Sqrt((ArchRadius + Half) * (ArchRadius + Half) - (ArchRadius - Half) * (ArchRadius - Half));
    double meetDegrees = Degrees(Math.Atan2(meet, -(ArchRadius - Half)));

    var c = new Contour();
    c.Start(new P(LeftStem - Half, LowBottom));
    c.ArcTo(new P(LeftStem, LowBottom), Half, 180, 360, wholeCap: true);
    c.LineTo(new P(LeftStem + Half, archY));
    c.ArcTo(arch, ArchRadius - Half, 180, 0);
    c.LineTo(new P(RightStem - Half, LowBottom));
    c.ArcTo(new P(RightStem, LowBottom), Half, 180, 360, wholeCap: true);
    c.LineTo(new P(RightStem + Half, archY));
    c.ArcTo(arch, ArchRadius + Half, 0, meetDegrees);
    c.LineTo(new P(LeftStem + Half, HighTop));
    c.ArcTo(new P(LeftStem, HighTop), Half, 0, 180, wholeCap: true);
    c.LineTo(new P(LeftStem - Half, LowBottom));
    c.Close();
    return new Glyph(c, Hole: null);
}

// e: a bar across the middle of the bowl, and the bowl from the bar's right end
// counterclockwise round to a terminal at the lower right; the eye above the bar is a hole.
// Counterclockwise from the bowl's right, on the bar's level
static Glyph GlyphE()
{
    const double Terminal = 320;
    var bowl = new P(BowlX, BowlY);
    double outer = BowlRadius + Half;
    double inner = BowlRadius - Half;
    double barRight = BowlX + BowlRadius;

    // where the bar's edges meet the inside of the bowl
    double barDegrees = Degrees(Math.Asin(Half / inner));

    var c = new Contour();
    c.Start(new P(barRight + Half, BowlY));
    c.ArcTo(bowl, outer, 0, Terminal);
    c.ArcTo(Polar(bowl, BowlRadius, Terminal), Half, Terminal, Terminal + 180, wholeCap: true);
    c.ArcTo(bowl, inner, Terminal, 180 + barDegrees);
    c.LineTo(new P(barRight, BowlY - Half));
    c.ArcTo(new P(barRight, BowlY), Half, 270, 360, wholeCap: true);
    c.Close();

    var eye = new Contour();
    eye.Start(Polar(bowl, inner, barDegrees));
    eye.ArcTo(bowl, inner, barDegrees, 180 - barDegrees);
    eye.LineTo(Polar(bowl, inner, barDegrees));
    eye.Close();
    return new Glyph(c, eye);
}

// ---- the other letters: skeletons of lines and arcs, stroked with the round pen ----

static Dictionary<char, Glyph[]> Alphabet()
{
    const double Dot = Advance / 2;
    double archY = LowTop - ArchRadius;
    double cupY = LowBottom + ArchRadius;
    double bowlLeft = BowlX - BowlRadius;
    double bowlRight = BowlX + BowlRadius;

    // s: two half circles stacked, touching in the middle
    double sRadius = XHeight / 4 - Half / 2;
    double sTop = BowlY + sRadius;
    double sBottom = BowlY - sRadius;

    // the tails of g and j turn left under the baseline
    const double TailRadius = 140;
    double tailY = DeepBottom + TailRadius;

    return new Dictionary<char, Glyph[]>
    {
        ['a'] = [Ring(), Stroke(Line(bowlRight, LowBottom, bowlRight, LowTop))],
        ['b'] = [Ring(), Stroke(Line(bowlLeft, LowBottom, bowlLeft, HighTop))],
        ['c'] = [Stroke(Arc(BowlX, BowlY, BowlRadius, 45, 315))],
        ['d'] = [Ring(), Stroke(Line(bowlRight, LowBottom, bowlRight, HighTop))],
        ['f'] = [Stroke(Line(300, LowBottom, 300, HighTop - TailRadius), Arc(300 + TailRadius, HighTop - TailRadius, TailRadius, 180, 90), Line(300 + TailRadius, HighTop, 470, HighTop)), Stroke(Line(150, LowTop, 450, LowTop))],
        ['g'] = [Ring(), Stroke(Line(bowlRight, LowTop, bowlRight, tailY), Arc(bowlRight - TailRadius, tailY, TailRadius, 0, -100))],
        ['i'] = [Stroke(Line(Dot, LowBottom, Dot, LowTop)), Disc(Dot, DotY, DotRadius)],
        ['j'] = [Stroke(Line(Dot, LowTop, Dot, tailY), Arc(Dot - TailRadius, tailY, TailRadius, 0, -90)), Disc(Dot, DotY, DotRadius)],
        ['k'] = [Stroke(Line(150, LowBottom, 150, HighTop)), Stroke(Line(150, 230, 480, LowTop)), Stroke(Line(280, 314.7, 490, LowBottom))],
        ['l'] = [Stroke(Line(270, HighTop, 270, LowBottom + TailRadius), Arc(270 + TailRadius, LowBottom + TailRadius, TailRadius, 180, 270), Line(270 + TailRadius, LowBottom, 460, LowBottom))],
        ['m'] = [Stroke(Line(110, LowBottom, 110, LowTop)), Stroke(Arc(205, LowTop - 95, 95, 180, 0), Line(300, LowTop - 95, 300, LowBottom)), Stroke(Arc(395, LowTop - 95, 95, 180, 0), Line(490, LowTop - 95, 490, LowBottom))],
        ['n'] = [Stroke(Line(LeftStem, LowBottom, LeftStem, LowTop)), Stroke(Arc(BowlX, archY, ArchRadius, 180, 0), Line(RightStem, archY, RightStem, LowBottom))],
        ['o'] = [Ring()],
        ['p'] = [Ring(), Stroke(Line(bowlLeft, DeepBottom, bowlLeft, LowTop))],
        ['q'] = [Ring(), Stroke(Line(bowlRight, DeepBottom, bowlRight, LowTop))],
        ['r'] = [Stroke(Line(200, LowBottom, 200, LowTop)), Stroke(Arc(325, LowTop - 125, 125, 180, 90), Line(325, LowTop, 430, LowTop))],
        ['s'] = [Stroke(Arc(BowlX, sTop, sRadius, 20, 270), Arc(BowlX, sBottom, sRadius, 90, -160))],
        ['u'] = [Stroke(Line(LeftStem, LowTop, LeftStem, cupY), Arc(BowlX, cupY, ArchRadius, 180, 360), Line(RightStem, cupY, RightStem, LowTop))],
        ['v'] = [Stroke(Line(105, LowTop, 300, LowBottom)), Stroke(Line(300, LowBottom, 495, LowTop))],
        ['w'] = [Stroke(Line(95, LowTop, 197, LowBottom)), Stroke(Line(197, LowBottom, 300, LowTop)), Stroke(Line(300, LowTop, 403, LowBottom)), Stroke(Line(403, LowBottom, 505, LowTop))],
        ['x'] = [Stroke(Line(105, LowTop, 495, LowBottom)), Stroke(Line(495, LowTop, 105, LowBottom))],
        ['y'] = [Stroke(Line(105, LowTop, 300, LowBottom)), Stroke(Line(495, LowTop, 190, DeepBottom))],
        ['z'] = [Stroke(Line(105, LowTop, 495, LowTop)), Stroke(Line(495, LowTop, 105, LowBottom)), Stroke(Line(105, LowBottom, 495, LowBottom))],
    };
}

static Seg Line(double x1, double y1, double x2, double y2) => new LineSeg(new P(x1, y1), new P(x2, y2));

static Seg Arc(double centerX, double centerY, double radius, double fromDegrees, double toDegrees) => new ArcSeg(new P(centerX, centerY), radius, fromDegrees, toDegrees);

// a bowl: the pen round a circle, a ring with a hole
static Glyph Ring()
{
    var center = new P(BowlX, BowlY);
    return new Glyph(Circle(center, BowlRadius + Half), Circle(center, BowlRadius - Half));
}

// a dot: a disc
static Glyph Disc(double x, double y, double radius) => new Glyph(Circle(new P(x, y), radius), Hole: null);

static Contour Circle(P center, double radius)
{
    var c = new Contour();
    c.Start(Polar(center, radius, 0));
    c.ArcTo(center, radius, 0, 360);
    c.Close();
    return c;
}

// the outline of the pen's stroke along a chain of segments, each starting where the last
// ended and heading the same way: the right side forward, a round cap, the left side back,
// a round cap
static Glyph Stroke(params Seg[] chain)
{
    for (int i = 1; i < chain.Length; i++)
    {
        if (!Near(chain[i].Start, chain[i - 1].End) || !Near(chain[i].StartDirection, chain[i - 1].EndDirection))
        {
            throw new InvalidOperationException($"segment {i} doesn't continue segment {i - 1}: {chain[i - 1]} then {chain[i]}");
        }
    }

    var c = new Contour();
    var first = chain[0];
    var last = chain[chain.Length - 1];
    c.Start(first.Start + Rotate(first.StartDirection, -90) * Half);
    foreach (var segment in chain)
    {
        segment.RightSide(c, Half);
    }

    double endDegrees = Degrees(Math.Atan2(last.EndDirection.Y, last.EndDirection.X));
    c.ArcTo(last.End, Half, endDegrees - 90, endDegrees + 90, wholeCap: true);
    for (int i = chain.Length - 1; i >= 0; i--)
    {
        chain[i].LeftSideBack(c, Half);
    }

    double startDegrees = Degrees(Math.Atan2(first.StartDirection.Y, first.StartDirection.X));
    c.ArcTo(first.Start, Half, startDegrees + 90, startDegrees + 270, wholeCap: true);
    c.Close();
    return new Glyph(c, Hole: null);
}

static class Geo
{
    public static P Polar(P center, double radius, double degrees)
    {
        double radians = degrees * Math.PI / 180;
        return new P(center.X + radius * Math.Cos(radians), center.Y + radius * Math.Sin(radians));
    }

    public static P Rotate(P p, double degrees)
    {
        double radians = degrees * Math.PI / 180;
        double cos = Math.Cos(radians);
        double sin = Math.Sin(radians);
        return new P(p.X * cos - p.Y * sin, p.X * sin + p.Y * cos);
    }

    public static double Degrees(double radians) => radians * 180 / Math.PI;

    public static bool Near(P a, P b) => Math.Abs(a.X - b.X) < 1e-6 && Math.Abs(a.Y - b.Y) < 1e-6;

    /// <summary>The file's numbers: at most six decimals, never a negative zero</summary>
    public static string Format(double value) => (Math.Round(value, 6) + 0.0).ToString("0.######", CultureInfo.InvariantCulture);
}

record struct P(double X, double Y)
{
    public static P operator +(P a, P b) => new P(a.X + b.X, a.Y + b.Y);

    public static P operator -(P a, P b) => new P(a.X - b.X, a.Y - b.Y);

    public static P operator *(P a, double factor) => new P(a.X * factor, a.Y * factor);

    public double Length => Math.Sqrt(X * X + Y * Y);

    public P Normalized => this * (1 / Length);

    public bool IsZero => Math.Abs(X) < 1e-9 && Math.Abs(Y) < 1e-9;
}

/// <summary>An anchor of a contour with its two handles as offsets from it (zero: on the anchor)</summary>
class Anchor
{
    public P At;
    public P In;
    public P Out;
}

/// <summary>
/// A closed outline built from lines and circular arcs: anchors with cubic handles, an arc of
/// up to 90 degrees per piece (a cap of 180: wholeCap), the handles 4/3 tan(sweep/4) of the
/// radius long
/// </summary>
class Contour
{
    public List<Anchor> Anchors { get; } = new();

    P Current => Anchors[Anchors.Count - 1].At;

    public void Start(P p)
    {
        Anchors.Add(new Anchor { At = p });
    }

    public void LineTo(P p)
    {
        Anchors.Add(new Anchor { At = p });
    }

    public void ArcTo(P center, double radius, double fromDegrees, double toDegrees, bool wholeCap = false)
    {
        var expected = Polar(center, radius, fromDegrees);
        if (!Near(Current, expected))
        {
            throw new InvalidOperationException($"the arc starts at {expected}, not at {Current}");
        }

        double sweep = toDegrees - fromDegrees;
        var cuts = new List<double> { fromDegrees };
        if (!wholeCap)
        {
            // at the multiples of 90 degrees strictly inside the arc, so that a quarter turn is
            // one piece and the top of an arch is an anchor
            double step = Math.Sign(sweep) * 90;
            double cut = sweep > 0 ? Math.Floor(fromDegrees / 90 + 1) * 90 : Math.Ceiling(fromDegrees / 90 - 1) * 90;
            while (sweep > 0 ? cut < toDegrees - 1e-9 : cut > toDegrees + 1e-9)
            {
                cuts.Add(cut);
                cut += step;
            }
        }

        cuts.Add(toDegrees);
        for (int i = 1; i < cuts.Count; i++)
        {
            AddArcPiece(center, radius, cuts[i - 1], cuts[i]);
        }
    }

    void AddArcPiece(P center, double radius, double fromDegrees, double toDegrees)
    {
        double sweep = toDegrees - fromDegrees;
        double k = 4.0 / 3 * Math.Tan(Math.Abs(sweep) * Math.PI / 720) * radius;
        Anchors[Anchors.Count - 1].Out = Tangent(fromDegrees, sweep > 0) * k;
        Anchors.Add(new Anchor { At = Polar(center, radius, toDegrees), In = Tangent(toDegrees, sweep > 0) * -k });
    }

    static P Tangent(double degrees, bool counterclockwise)
    {
        double radians = degrees * Math.PI / 180;
        return counterclockwise ? new P(-Math.Sin(radians), Math.Cos(radians)) : new P(Math.Sin(radians), -Math.Cos(radians));
    }

    /// <summary>Ends the outline: a last anchor on the first one is the first one, with its in handle</summary>
    public void Close()
    {
        var first = Anchors[0];
        var last = Anchors[Anchors.Count - 1];
        if (Near(first.At, last.At))
        {
            first.In = last.In;
            Anchors.RemoveAt(Anchors.Count - 1);
        }
    }

    /// <summary>Whether the outline goes on in the same direction through the anchor, as against turning a corner</summary>
    public bool IsSmooth(int index)
    {
        int count = Anchors.Count;
        var anchor = Anchors[index];
        var arriving = anchor.In.IsZero ? anchor.At - Anchors[(index + count - 1) % count].At : anchor.In * -1;
        var leaving = anchor.Out.IsZero ? Anchors[(index + 1) % count].At - anchor.At : anchor.Out;
        arriving = arriving.Normalized;
        leaving = leaving.Normalized;
        double cross = arriving.X * leaving.Y - arriving.Y * leaving.X;
        double dot = arriving.X * leaving.X + arriving.Y * leaving.Y;
        return dot > 0 && Math.Abs(cross) < Math.Sin(3 * Math.PI / 180);
    }

    /// <summary>The Path attribute: a piece per anchor, L when both handles are on their anchors, else C with the offsets scaled</summary>
    public string PathText(double scale)
    {
        var pieces = new List<string>();
        int count = Anchors.Count;
        for (int i = 0; i < count; i++)
        {
            var outHandle = Anchors[i].Out;
            var inHandle = Anchors[(i + 1) % count].In;
            pieces.Add(outHandle.IsZero && inHandle.IsZero
                ? "L"
                : $"C {Offset(outHandle, scale)} {Offset(inHandle, scale)}");
        }

        return string.Join(" ", pieces);
    }

    static string Offset(P handle, double scale) => Format(handle.X * scale) + "," + Format(handle.Y * scale);
}

/// <summary>A piece of a letter: a closed outline, and the hole it leaves out, if any</summary>
record Glyph(Contour Outline, Contour Hole);

/// <summary>A big letter's hue: the name its styles start with, and the color under each theme</summary>
record Hue(string Name, string Light, string Dark);

/// <summary>The names of a big letter's figures, which its copies are built on</summary>
record GlyphNames(string Path, IList<string> Anchors, string Hole, IList<string> HoleAnchors);

/// <summary>A segment of a letter's skeleton: where the pen goes</summary>
abstract record Seg
{
    public abstract P Start { get; }

    public abstract P End { get; }

    public abstract P StartDirection { get; }

    public abstract P EndDirection { get; }

    /// <summary>The pen's right edge, from the start to the end</summary>
    public abstract void RightSide(Contour c, double half);

    /// <summary>The pen's left edge, from the end back to the start</summary>
    public abstract void LeftSideBack(Contour c, double half);
}

record LineSeg(P A, P B) : Seg
{
    public override P Start => A;

    public override P End => B;

    public override P StartDirection => (B - A).Normalized;

    public override P EndDirection => StartDirection;

    public override void RightSide(Contour c, double half) => c.LineTo(B + Rotate(StartDirection, -90) * half);

    public override void LeftSideBack(Contour c, double half) => c.LineTo(A + Rotate(StartDirection, 90) * half);
}

/// <summary>An arc of a circle, counterclockwise when To is greater than From</summary>
record ArcSeg(P Center, double Radius, double From, double To) : Seg
{
    bool Counterclockwise => To > From;

    public override P Start => Polar(Center, Radius, From);

    public override P End => Polar(Center, Radius, To);

    public override P StartDirection => Tangent(From);

    public override P EndDirection => Tangent(To);

    P Tangent(double degrees)
    {
        double radians = degrees * Math.PI / 180;
        return Counterclockwise ? new P(-Math.Sin(radians), Math.Cos(radians)) : new P(Math.Sin(radians), -Math.Cos(radians));
    }

    // going counterclockwise the center is on the left, so the right edge is the outer one
    public override void RightSide(Contour c, double half) => c.ArcTo(Center, Counterclockwise ? Radius + half : Radius - half, From, To);

    public override void LeftSideBack(Contour c, double half) => c.ArcTo(Center, Counterclockwise ? Radius - half : Radius + half, To, From);
}
