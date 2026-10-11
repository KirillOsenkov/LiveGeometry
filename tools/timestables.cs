#:property Nullable=disable
#:property PublishAot=false

// timestables - writes the "Times Tables on a Circle" gallery drawing: 200 points around a
// circle, numbered 0 to 199, and a chord from every point i to the point k * i (mod 200),
// where k is a slider. The chords are segments whose far ends are points by coordinates
// over k, so the picture morphs as the slider moves: a cardioid at 2, a nephroid at 3...
//
//   dotnet tools/timestables.cs -- <out.lgf>

using System.Globalization;
using System.Text;

if (args.Length < 1)
{
    Console.WriteLine("usage: timestables <out.lgf>");
    return 1;
}

const int Count = 200;
const double Radius = 4;
var invariant = CultureInfo.InvariantCulture;
var text = new StringBuilder();

void Write(string line)
{
    text.Append(line).Append("\r\n");
}

string Format(double value)
{
    return value.ToString("0.######", invariant);
}

Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
Write("<Drawing Version=\"1\" Creator=\"LiveGeometry.App\">");
Write("  <Viewport Left=\"-6\" Top=\"5\" Right=\"6\" Bottom=\"-6\" />");
Write("  <Styles>");
Write("    <ShapeStyle Name=\"Ring\" IsFilled=\"false\" Color=\"#FF8A94A6\" StrokeWidth=\"1.5\" />");
Write("    <LineStyle Name=\"Chord\" Color=\"#902F7BD6\" StrokeWidth=\"1\">");
Write("      <Dark Color=\"#A062AEFF\" />");
Write("    </LineStyle>");
Write("  </Styles>");
Write("  <Figures>");
Write("    <PointByCoordinates Name=\"O\" Visible=\"false\" X=\"0\" Y=\"0\" />");
Write($"    <PointByCoordinates Name=\"Rim\" Visible=\"false\" X=\"{Format(Radius)}\" Y=\"0\" />");
Write("    <Circle Name=\"Ring\" Style=\"Ring\">");
Write("      <Dependency Name=\"O\" />");
Write("      <Dependency Name=\"Rim\" />");
Write("    </Circle>");
Write("    <Slider Name=\"k\" X=\"-5\" Y=\"-5.3\" Value=\"2\" />");

for (int i = 0; i < Count; i++)
{
    double angle = 2 * Math.PI * i / Count;
    Write($"    <PointByCoordinates Name=\"P{i}\" Visible=\"false\" X=\"{Format(Radius * Math.Cos(angle))}\" Y=\"{Format(Radius * Math.Sin(angle))}\" />");
    Write($"    <PointByCoordinates Name=\"Q{i}\" Visible=\"false\" X=\"{Format(Radius)} * cos(k * {Format(angle)})\" Y=\"{Format(Radius)} * sin(k * {Format(angle)})\">");
    Write("      <Dependency Name=\"k\" />");
    Write("    </PointByCoordinates>");
    Write($"    <Segment Name=\"Chord{i}\" Style=\"Chord\">");
    Write($"      <Dependency Name=\"P{i}\" />");
    Write($"      <Dependency Name=\"Q{i}\" />");
    Write("    </Segment>");
}

Write("    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"Times Tables on a Circle\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"200 dots sit around the circle, numbered 0 to 199. Each dot is joined by a line to the dot with k times its number: with k at 2, dot 7 goes to dot 14, dot 50 to dot 100, and past 199 the count wraps around to 0 again.\\n\\nThat is all. Nobody drew the heart: it just appears out of the 2 times table. Drag the slider to 3 and the heart grows a second lobe; 4 gives three, 5 gives four. Slide slowly between the whole numbers and watch one pattern melt into the next.\\n\\nTry 51, 99, 101, 199 and 201. Why does 199 look like 2? Because 199 times a number is one less than 200 times it, which wraps around to the same dot as minus one times it.\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("  </Figures>");
text.Append("</Drawing>");
File.WriteAllText(args[0], text.ToString(), new UTF8Encoding(false));
Console.WriteLine("wrote " + args[0]);
return 0;
