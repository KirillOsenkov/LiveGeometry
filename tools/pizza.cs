#:property Nullable=disable
#:property PublishAot=false

// pizza - writes the "Slice the Pizza" gallery drawing: a pizza cut into twelve slices, which
// a slider takes apart and lines up point up, point down into a row that is nearly a
// rectangle: half the crust along the top, half along the bottom, a radius high. The area of
// the circle, pi r squared, read off a pizza. Each slice is a sector whose center and ends are
// points by coordinates, moving and turning between their two places by the slider.
//
//   dotnet tools/pizza.cs -- <out.lgf>

using System.Globalization;
using System.Text;

if (args.Length < 1)
{
    Console.WriteLine("usage: pizza <out.lgf>");
    return 1;
}

const int Slices = 12;
const double Radius = 1.6;
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

// where the pizza sits, and where the row of slices goes: the slices with their points down
// have their points on the row's top line, the others on its bottom line, each slice half
// a chord further along than the one before
var pizzaCenter = (X: 0.0, Y: 1.9);
double sliceAngle = 2 * Math.PI / Slices;
double chord = 2 * Radius * Math.Sin(sliceAngle / 2);
double height = Radius * Math.Cos(sliceAngle / 2);
double rowWidth = (Slices / 2.0 + 0.5) * chord;
double rowTop = -1.2;

Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
Write("<Drawing Version=\"1\" Creator=\"LiveGeometry.App\">");
Write("  <Viewport Left=\"-6\" Top=\"4\" Right=\"6\" Bottom=\"-4.5\" />");
Write("  <Styles>");
Write("    <ShapeStyle Name=\"Slice\" IsFilled=\"true\" Fill=\"#FFF6C34A\" Color=\"#FFC8782A\" StrokeWidth=\"2.5\">");
Write("      <Dark Fill=\"#FFD9A832\" Color=\"#FFA85E1E\" />");
Write("    </ShapeStyle>");
Write("    <PointStyle Name=\"Pepperoni\" Size=\"13\" Fill=\"#FFC0392B\" Color=\"#FF8E1F14\" StrokeWidth=\"1\" />");
Write("  </Styles>");
Write("  <Figures>");
Write("    <Slider Name=\"unroll\" X=\"-5.5\" Y=\"-3.8\" Value=\"0.6\" Maximum=\"4\" />");

for (int i = 0; i < Slices; i++)
{
    // the slice's middle direction in the pizza, and in the row: down for even slices, up
    // for odd, going round the shorter way
    double startAngle = i * sliceAngle + sliceAngle / 2;
    double targetAngle = i % 2 == 0 ? -Math.PI / 2 : Math.PI / 2;
    while (targetAngle - startAngle > Math.PI)
    {
        targetAngle -= 2 * Math.PI;
    }

    while (startAngle - targetAngle > Math.PI)
    {
        targetAngle += 2 * Math.PI;
    }

    double targetX = (i + 1) / 2.0 * chord - rowWidth / 2;
    double targetY = i % 2 == 0 ? rowTop : rowTop - height;

    // t from 0 (in the pizza) to 1 (in the row), the slider running 0 to 4; the slices
    // first spread apart and start turning only on the way, so that a small nudge of the
    // slider gives an exploded pizza, not a whirl
    string t = "unroll / 4";
    string turn = "clamp((unroll / 4 - 0.3) / 0.7, 0, 1)";
    string centerX = $"{Format(pizzaCenter.X)} + ({Format(targetX - pizzaCenter.X)}) * {t}";
    string centerY = $"{Format(pizzaCenter.Y)} + ({Format(targetY - pizzaCenter.Y)}) * {t}";
    string middle = $"{Format(startAngle)} + ({Format(targetAngle - startAngle)}) * {turn}";
    Write($"    <PointByCoordinates Name=\"Center{i}\" Visible=\"false\" X=\"{centerX}\" Y=\"{centerY}\">");
    Write("      <Dependency Name=\"unroll\" />");
    Write("    </PointByCoordinates>");
    Write($"    <PointByCoordinates Name=\"Start{i}\" Visible=\"false\" X=\"Center{i}.X + {Format(Radius)} * cos({middle} - {Format(sliceAngle / 2)})\" Y=\"Center{i}.Y + {Format(Radius)} * sin({middle} - {Format(sliceAngle / 2)})\">");
    Write($"      <Dependency Name=\"Center{i}\" />");
    Write("      <Dependency Name=\"unroll\" />");
    Write("    </PointByCoordinates>");
    Write($"    <PointByCoordinates Name=\"End{i}\" Visible=\"false\" X=\"Center{i}.X + {Format(Radius)} * cos({middle} + {Format(sliceAngle / 2)})\" Y=\"Center{i}.Y + {Format(Radius)} * sin({middle} + {Format(sliceAngle / 2)})\">");
    Write($"      <Dependency Name=\"Center{i}\" />");
    Write("      <Dependency Name=\"unroll\" />");
    Write("    </PointByCoordinates>");
    Write($"    <CircleSector Name=\"Slice{i}\" Style=\"Slice\">");
    Write($"      <Dependency Name=\"Center{i}\" />");
    Write($"      <Dependency Name=\"Start{i}\" />");
    Write($"      <Dependency Name=\"End{i}\" />");
    Write("    </CircleSector>");
    Write($"    <PointByCoordinates Name=\"Pepperoni{i}\" Style=\"Pepperoni\" X=\"Center{i}.X + {Format(Radius * 0.62)} * cos({middle})\" Y=\"Center{i}.Y + {Format(Radius * 0.62)} * sin({middle})\">");
    Write($"      <Dependency Name=\"Center{i}\" />");
    Write("      <Dependency Name=\"unroll\" />");
    Write("    </PointByCoordinates>");
}

Write("    <Label Name=\"Title\" Style=\"GalleryTitle\" Text=\"Slice the Pizza\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"16\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("    <Label Name=\"Description\" Style=\"GalleryText\" Text=\"Why is the area of a circle π times the radius squared? Cut a pizza into slices and drag the slider: the slices line up, point up, point down, into a shape that is almost a rectangle.\\n\\nHalf the crust runs along its top and half along its bottom, so the rectangle is half the circumference wide: π times the radius. And it is one slice high: the radius. Width times height is π times the radius times the radius.\\n\\nWith twelve slices the edges are still bumpy. Imagine a hundred slices, or a million: the bumps vanish and the shape is exactly a rectangle. That is the idea behind calculus, two thousand years before calculus was invented.\" Pin=\"TopRight\" OffsetX=\"16\" OffsetY=\"64\" WrapWidth=\"400\" Backdrop=\"true\" />");
Write("  </Figures>");
text.Append("</Drawing>");
File.WriteAllText(args[0], text.ToString(), new UTF8Encoding(false));
Console.WriteLine("wrote " + args[0]);
return 0;
