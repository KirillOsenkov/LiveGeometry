#:property TargetFramework=net10.0-windows
#:property UseWindowsForms=true
#:property Nullable=disable
#:property PublishAot=false

// contactsheet - tiles the PNGs of a folder (the output of `LiveGeometry.Desktop.exe --check`)
// into one image with a caption under each, to eyeball a whole folder of drawings at once.
//
//   dotnet tools/contactsheet.cs -- <png folder> <out.png> [columns=4] [tile width=420]

using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

if (args.Length < 2)
{
    Console.WriteLine("usage: contactsheet <png folder> <out.png> [columns] [tile width]");
    return 1;
}

var folder = args[0];
var output = args[1];
int columns = args.Length > 2 ? int.Parse(args[2]) : 4;
int tileWidth = args.Length > 3 ? int.Parse(args[3]) : 420;
const int captionHeight = 22;
const int gap = 10;

var files = Directory.GetFiles(folder, "*.png").OrderBy(f => f).ToArray();
if (files.Length == 0)
{
    Console.WriteLine("no PNGs in " + folder);
    return 1;
}

// all tiles share the aspect ratio of the first image (they come from one window size)
using var first = Image.FromFile(files[0]);
int tileHeight = tileWidth * first.Height / first.Width;
int rows = (files.Length + columns - 1) / columns;
int width = columns * (tileWidth + gap) + gap;
int height = rows * (tileHeight + captionHeight + gap) + gap;

using var sheet = new Bitmap(width, height);
using var graphics = Graphics.FromImage(sheet);
graphics.Clear(Color.FromArgb(0x60, 0x60, 0x60));
graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
using var font = new Font("Segoe UI", 10, FontStyle.Bold);
for (int i = 0; i < files.Length; i++)
{
    int x = gap + (i % columns) * (tileWidth + gap);
    int y = gap + (i / columns) * (tileHeight + captionHeight + gap);
    using var image = Image.FromFile(files[i]);
    graphics.DrawImage(image, x, y, tileWidth, tileHeight);
    graphics.DrawRectangle(Pens.Black, x, y, tileWidth, tileHeight);
    graphics.DrawString(Path.GetFileNameWithoutExtension(files[i]), font, Brushes.White, x, y + tileHeight + 2);
}

sheet.Save(output, ImageFormat.Png);
Console.WriteLine("saved " + output + " " + width + "x" + height + " (" + files.Length + " images)");
return 0;
