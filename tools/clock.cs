#:property Nullable=disable
#:property PublishAot=false

// clock - writes the "Clock Hands" gallery drawing: a wall clock whose hands are turned by
// the slider "time" (hours, 0 to 12). The hour and minute hands are tapered polygons whose
// corners are rotated points about the center, the angle a hidden label over the slider;
// the second hand sweeps once per minute of clock time. A wooden frame, a brass bezel,
// sixty ticks and twelve numerals, and the angle between the hands written under it. The
// discs are polygons, not circles: polygons draw under circles, and the hands are polygons.
//
//   dotnet tools/clock.cs -- <out.lgf>

using System.Globalization;
using System.Text;

if (args.Length < 1)
{
    Console.WriteLine("usage: clock <out.lgf>");
    return 1;
}

const double FaceRadius = 3;
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

void WriteFixed(string name, double x, double y)
{
    Write($"    <PointByCoordinates Name=\"{name}\" Visible=\"false\" X=\"{Format(x)}\" Y=\"{Format(y)}\" />");
}

// a disc as a 72-gon rather than a circle: polygons draw in a layer under circles, so the
// hands, which are polygons, would vanish under a filled circle; among polygons the list's
// order decides, and the discs come first
void WriteDisc(string name, string style, double radius)
{
    const int Sides = 72;
    for (int i = 0; i < Sides; i++)
    {
        double angle = 2 * Math.PI * i / Sides;
        WriteFixed($"{name}{i}", radius * Math.Cos(angle), radius * Math.Sin(angle));
    }

    Write($"    <Polygon Name=\"{name}\" Style=\"{style}\">");
    for (int i = 0; i < Sides; i++)
    {
        Write($"      <Dependency Name=\"{name}{i}\" />");
    }

    Write("    </Polygon>");
}

Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
Write("<Drawing Version=\"1\" Creator=\"LiveGeometry.App\">");
Write("  <Viewport Left=\"-7\" Top=\"4.5\" Right=\"7\" Bottom=\"-5.8\" />");
Write("  <Styles>");
Write("    <ShapeStyle Name=\"Frame\" IsFilled=\"true\" Color=\"#FF2E1A0C\" StrokeWidth=\"1.5\">");
Write("      <Fill>");
Write("        <LinearGradientBrush StartPoint=\"0,0\" EndPoint=\"1,1\">");
Write("          <GradientStop Color=\"#FF9C6B3C\" Offset=\"0\" />");
Write("          <GradientStop Color=\"#FF4A2E16\" Offset=\"1\" />");
Write("        </LinearGradientBrush>");
Write("      </Fill>");
Write("    </ShapeStyle>");
Write("    <ShapeStyle Name=\"Bezel\" IsFilled=\"true\" Color=\"#FF6B4F14\" StrokeWidth=\"1\">");
Write("      <Fill>");
Write("        <LinearGradientBrush StartPoint=\"0,0\" EndPoint=\"1,1\">");
Write("          <GradientStop Color=\"#FFF2D98A\" Offset=\"0\" />");
Write("          <GradientStop Color=\"#FF9C7A2A\" Offset=\"1\" />");
Write("        </LinearGradientBrush>");
Write("      </Fill>");
Write("    </ShapeStyle>");
Write("    <ShapeStyle Name=\"Face\" IsFilled=\"true\" Color=\"#FF8A6A3A\" StrokeWidth=\"1\">");
Write("      <Fill>");
Write("        <LinearGradientBrush StartPoint=\"0,0\" EndPoint=\"1,1\">");
Write("          <GradientStop Color=\"#FFFFFEF8\" Offset=\"0\" />");
Write("          <GradientStop Color=\"#FFEFE3C8\" Offset=\"1\" />");
Write("        </LinearGradientBrush>");
Write("      </Fill>");
Write("    </ShapeStyle>");
Write("    <LineStyle Name=\"MinuteTick\" Color=\"#FF6B5A3A\" StrokeWidth=\"1.25\" />");
Write("    <LineStyle Name=\"HourTick\" Color=\"#FF3A2E1A\" StrokeWidth=\"3\" />");
Write("    <TextStyle Name=\"Numeral\" FontSize=\"20\" Color=\"#FF2E1A0C\" FontFamily=\"Segoe UI\" Bold=\"true\">");
Write("      <Dark Color=\"#FF2E1A0C\" />");
Write("    </TextStyle>");
Write("    <ShapeStyle Name=\"Hand\" IsFilled=\"true\" Color=\"#FF101418\" StrokeWidth=\"1\">");
Write("      <Fill>");
Write("        <LinearGradientBrush StartPoint=\"0,0\" EndPoint=\"1,1\">");
Write("          <GradientStop Color=\"#FF5A6B85\" Offset=\"0\" />");
Write("          <GradientStop Color=\"#FF1A2130\" Offset=\"1\" />");
Write("        </LinearGradientBrush>");
Write("      </Fill>");
Write("    </ShapeStyle>");
Write("    <LineStyle Name=\"SecondHand\" Color=\"#FFD83B3B\" StrokeWidth=\"2\" />");
Write("    <PointStyle Name=\"Hub\" Size=\"16\" Fill=\"#FF3A4558\" Color=\"#FF101418\" StrokeWidth=\"1.5\" />");
Write("    <PointStyle Name=\"HubCap\" Size=\"7\" Fill=\"#FFD83B3B\" Color=\"#FF8E1F14\" StrokeWidth=\"0.5\" />");
Write("    <ShapeStyle Name=\"Wedge\" IsFilled=\"true\" Fill=\"#60F28C28\" Color=\"#FFD2691E\" StrokeWidth=\"2\" />");
Write("    <TextStyle Name=\"Degrees\" FontSize=\"20\" Color=\"#FFD2691E\" FontFamily=\"Segoe UI\" Bold=\"true\">");
Write("      <Dark Color=\"#FFFFA552\" />");
Write("    </TextStyle>");
Write("  </Styles>");
Write("  <Figures>");
WriteFixed("O", 0, 0);
WriteDisc("Frame", "Frame", FaceRadius + 0.5);
WriteDisc("Bezel", "Bezel", FaceRadius + 0.22);
WriteDisc("Face", "Face", FaceRadius);

// the ticks: sixty, every fifth longer and heavier
for (int i = 0; i < 60; i++)
{
    bool hour = i % 5 == 0;
    double angle = Math.PI / 2 - 2 * Math.PI * i / 60;
    double inner = FaceRadius - (hour ? 0.3 : 0.14);
    double outer = FaceRadius - 0.06;
    WriteFixed($"Tick{i}In", inner * Math.Cos(angle), inner * Math.Sin(angle));
    WriteFixed($"Tick{i}Out", outer * Math.Cos(angle), outer * Math.Sin(angle));
    Write($"    <Segment Name=\"Tick{i}\" Style=\"{(hour ? "HourTick" : "MinuteTick")}\">");
    Write($"      <Dependency Name=\"Tick{i}In\" />");
    Write($"      <Dependency Name=\"Tick{i}Out\" />");
    Write("    </Segment>");
}

// the numerals: hidden points named after them, their names shown, centered by a pixel
// offset (about 11 px per digit at this size)
for (int hour = 1; hour <= 12; hour++)
{
    double angle = Math.PI / 2 - 2 * Math.PI * hour / 12;
    double radius = FaceRadius - 0.62;
    string name = hour.ToString(invariant);
    WriteFixed(name, radius * Math.Cos(angle), radius * Math.Sin(angle));
    double width = name.Length * 11;
    Write($"    <PointLabel Name=\"Numeral{hour}\" IsHitTestVisible=\"false\" Style=\"Numeral\" OffsetX=\"{Format(-width / 2)}\" OffsetY=\"-14\" ShowName=\"true\" ShowCoordinates=\"false\">");
    Write($"      <Dependency Name=\"{name}\" />");
    Write("    </PointLabel>");
}

Write("    <Slider Name=\"time\" X=\"-6\" Y=\"-5\" Value=\"2.3\" Maximum=\"12\" />");
Write("    <Label Name=\"HourAngle\" Visible=\"false\" Text=\"[rad(-30 * time)]\" X=\"0\" Y=\"-8\">");
Write("      <Dependency Name=\"time\" />");
Write("    </Label>");
Write("    <Label Name=\"MinuteAngle\" Visible=\"false\" Text=\"[rad(-360 * time)]\" X=\"0\" Y=\"-9\">");
Write("      <Dependency Name=\"time\" />");
Write("    </Label>");
Write("    <Label Name=\"SecondAngle\" Visible=\"false\" Text=\"[rad(-360 * (60 * time - floor(60 * time)))]\" X=\"0\" Y=\"-10\">");
Write("      <Dependency Name=\"time\" />");
Write("    </Label>");

// a hand: its outline drawn pointing up (at 12), every corner a rotated point about the
// center by the hand's angle
void WriteHand(string name, string angle, (double X, double Y)[] outline)
{
    var corners = new List<string>();
    for (int i = 0; i < outline.Length; i++)
    {
        string fixedName = $"{name}Shape{i}";
        string turned = $"{name}{i}";
        WriteFixed(fixedName, outline[i].X, outline[i].Y);
        Write($"    <RotatedPoint Name=\"{turned}\" Visible=\"false\">");
        Write($"      <Dependency Name=\"{fixedName}\" />");
        Write("      <Dependency Name=\"O\" />");
        Write($"      <Dependency Name=\"{angle}\" />");
        Write("    </RotatedPoint>");
        corners.Add(turned);
    }

    Write($"    <Polygon Name=\"{name}\" Style=\"Hand\">");
    foreach (var corner in corners)
    {
        Write($"      <Dependency Name=\"{corner}\" />");
    }

    Write("    </Polygon>");
}

WriteHand("HourHand", "HourAngle", new[] { (-0.13, -0.4), (0.13, -0.4), (0.08, 1.55), (0.0, 1.8), (-0.08, 1.55) });
WriteHand("MinuteHand", "MinuteAngle", new[] { (-0.1, -0.5), (0.1, -0.5), (0.05, 2.45), (0.0, 2.7), (-0.05, 2.45) });
WriteFixed("SecondTail", 0, -0.7);
WriteFixed("SecondTipUp", 0, 2.75);
Write("    <RotatedPoint Name=\"SecondBack\" Visible=\"false\">");
Write("      <Dependency Name=\"SecondTail\" />");
Write("      <Dependency Name=\"O\" />");
Write("      <Dependency Name=\"SecondAngle\" />");
Write("    </RotatedPoint>");
Write("    <RotatedPoint Name=\"SecondTip\" Visible=\"false\">");
Write("      <Dependency Name=\"SecondTipUp\" />");
Write("      <Dependency Name=\"O\" />");
Write("      <Dependency Name=\"SecondAngle\" />");
Write("    </RotatedPoint>");
Write("    <Segment Name=\"SecondHand\" Style=\"SecondHand\">");
Write("      <Dependency Name=\"SecondBack\" />");
Write("      <Dependency Name=\"SecondTip\" />");
Write("    </Segment>");
Write("    <PointByCoordinates Name=\"Hub\" Style=\"Hub\" X=\"0\" Y=\"0\" />");
Write("    <PointByCoordinates Name=\"HubCap\" Style=\"HubCap\" X=\"0\" Y=\"0\" />");

// the angle between the hands: its wedge at the hub, its number under the clock (a label
// draws under a filled face, so the number can't sit on it)
Write("    <AngleArc Name=\"Wedge\" Style=\"Wedge\" Radius=\"40\" Sweep=\"Smaller\">");
Write("      <Dependency Name=\"O\" />");
Write("      <Dependency Name=\"HourHand3\" />");
Write("      <Dependency Name=\"MinuteHand3\" />");
Write("    </AngleArc>");
Write("    <Label Name=\"Between\" Style=\"Degrees\" Text=\"[deg(ang(HourHand3, O, MinuteHand3))]° between the hands\" DecimalsToShow=\"1\" X=\"-2.4\" Y=\"-3.85\">");
Write("      <Dependency Name=\"HourHand3\" />");
Write("      <Dependency Name=\"O\" />");
Write("      <Dependency Name=\"MinuteHand3\" />");
Write("    </Label>");
Write("    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"Clock Hands\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"Drag the slider to turn the clock forward. It's [floor(time)] hours and [floor(60 * (time - floor(time)))] minutes, and the angle between the hands is written under the clock.\\n\\nThe minute hand goes around twelve times while the hour hand goes around once, so the minute hand catches up with the hour hand only every 65 minutes and change, not every hour. Count how many times they overlap between 12 and 12: it's 11, not 12. The red second hand goes around sixty times for every lap of the minute hand, which is why it spins so fast when you slide.\\n\\nA classic puzzle: at what time after 3 o'clock are the hands exactly on top of each other? And when do they make an exact right angle? Slide slowly and find out.\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\">");
Write("      <Dependency Name=\"time\" />");
Write("    </Label>");
Write("  </Figures>");
text.Append("</Drawing>");
File.WriteAllText(args[0], text.ToString(), new UTF8Encoding(false));
Console.WriteLine("wrote " + args[0]);
return 0;
