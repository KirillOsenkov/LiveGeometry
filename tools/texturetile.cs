// Makes the tile of a fill texture (ShapeStyle.Texture, FillTextures in the app and the player)
// from a scan of a hand-colored area: a patch of an evenly colored area, as a density per pixel -
// how much pigment the pixel has against the patch's average, measured from the paper's white,
// so 1 is the average and a fleck of paper is near 0 - flattened (the photo's light divided out:
// what varies over more than about --flatten pixels), made to repeat without a seam, and written
// as data both sides read: Main/Player/src/styles/<name>Tile.js and
// Main/Avalonia/LiveGeometry/Styles/<Name>Tile.cs.
//
//   dotnet run tools/texturetile.cs -- <Name> <image> <x> <y> [--size 256] [--overlap 48] [--flatten 40] [--preview <png>]
//
// The tiles in the repository are from Catriona Agg's puzzle drawings, their 1200 px pictures
// on Mathforge Notes (the grain of the paper, the pencil and the pen, nothing of the drawings):
//   Crayon   the green lower half of A Triangle Inside a Hexagon, at 380, 630
//   FeltTip  the yellow hexagon of Three Hexagons on a Circle, at 330, 380, --flatten 160 (the
//            bands of the pen's strokes are the texture, and stay)

// The image is a PNG (8 bits per channel, RGB or RGBA, not interlaced), read here with nothing
// but .NET's zlib: no image library for one patch.

#nullable disable

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

if (args.Length < 4)
{
    Console.WriteLine("usage: dotnet run tools/texturetile.cs -- <Name> <image> <x> <y> [--size 256] [--overlap 48] [--flatten 40] [--preview <png>]");
    return 2;
}

string name = args[0];
int x0 = int.Parse(args[2]), y0 = int.Parse(args[3]);
int size = Option("--size") is string s ? int.Parse(s) : 256;
int overlap = Option("--overlap") is string o ? int.Parse(o) : 48;
int flatten = Option("--flatten") is string reach ? int.Parse(reach) : 40;
string preview = Option("--preview");
int patch = size + overlap;

var (width, height, rgb) = Png.Read(args[1]);
if (x0 + patch > width || y0 + patch > height)
{
    Console.WriteLine($"error: the patch {patch} px square at {x0}, {y0} doesn't fit in the image ({width} x {height})");
    return 1;
}

// the patch's average color, and each pixel's pigment against it, measured from white
var pixels = new (int R, int G, int B)[patch * patch];
for (int y = 0; y < patch; y++)
{
    for (int x = 0; x < patch; x++)
    {
        int at = 3 * ((y0 + y) * width + x0 + x);
        pixels[y * patch + x] = (rgb[at], rgb[at + 1], rgb[at + 2]);
    }
}

double meanR = pixels.Average(p => p.R), meanG = pixels.Average(p => p.G), meanB = pixels.Average(p => p.B);
double meanDistance = (255 - meanR) + (255 - meanG) + (255 - meanB);
var density = pixels.Select(p => ((255 - p.R) + (255 - p.G) + (255 - p.B)) / meanDistance).ToArray();

// the light of the photo: divided out, through a wide blur of the density
density = Flatten(density, patch, radius: flatten);

// repeating: the last columns cross-faded into the first, then the same for the rows
density = Wrap(density, patch, size, overlap);

double mean = density.Average();
double spread = Math.Sqrt(density.Average(d => (d - mean) * (d - mean))) / mean;
Console.WriteLine($"spread (standard deviation over the mean) as scanned: {spread:0.000}");

// a speck of dust is no texture: what lies more than 3.5 spreads from the average is pulled back to it
density = density.Select(d => Math.Clamp(d, mean * (1 - 3.5 * spread), mean * (1 + 3.5 * spread))).ToArray();

// --spread: the variation scaled to it, for a scan whose ink nearly saturates it (a bright
// yellow keeps little but its blue channel, and its strokes come out faint)
if (Option("--spread") is string wanted)
{
    double now = Math.Sqrt(density.Average(d => (d - mean) * (d - mean))) / mean;
    double factor = double.Parse(wanted, System.Globalization.CultureInfo.InvariantCulture) / now;
    density = density.Select(d => mean + (d - mean) * factor).ToArray();
}
var bytes = density.Select(d => (byte)Math.Clamp(Math.Round(d / mean * 128), 0, 255)).ToArray();
string data = Convert.ToBase64String(bytes);
string repository = Png.FindRepository();
string camel = char.ToLowerInvariant(name[0]) + name.Substring(1);
File.WriteAllText(Path.Combine(repository, "Main", "Player", "src", "styles", camel + "Tile.js"),
    $"// Written by tools/texturetile.cs: the density of the {name} texture's tile (FillTextures), size x size\r\n"
    + "// bytes, row by row, 128 the average amount of pigment, 0 none (the paper).\r\n\r\n"
    + $"const {name}Tile = {{\r\n"
    + $"    size: {size},\r\n"
    + $"    data: \"{data}\"\r\n"
    + "};\r\n", new UTF8Encoding(false));
File.WriteAllText(Path.Combine(repository, "Main", "Avalonia", "LiveGeometry", "Styles", name + "Tile.cs"),
    "namespace LiveGeometry;\r\n\r\n"
    + "/// <summary>\r\n"
    + $"/// Written by tools/texturetile.cs: the density of the {name} texture's tile (<see cref=\"FillTextures\"/>),\r\n"
    + "/// Size x Size bytes, row by row, 128 the average amount of pigment, 0 none (the paper)\r\n"
    + "/// </summary>\r\n"
    + $"public static class {name}Tile\r\n"
    + "{\r\n"
    + $"    public const int Size = {size};\r\n\r\n"
    + $"    public const string Data = \"{data}\";\r\n"
    + "}\r\n", new UTF8Encoding(false));

Console.WriteLine($"{size} x {size}: density from {bytes.Min()} to {bytes.Max()}, {bytes.Count(b => b < 64) * 100.0 / bytes.Length:0.0}% under half the average");
if (preview != null)
{
    // the tile twice across in one color on white, as a fill would show it
    var output = new byte[3 * 4 * size * size];
    var hex = Convert.ToInt32((Option("--color") ?? "84DCC0").TrimStart('#'), 16);
    var color = (R: (double)((hex >> 16) & 255), G: (double)((hex >> 8) & 255), B: (double)(hex & 255));
    for (int y = 0; y < 2 * size; y++)
    {
        for (int x = 0; x < 2 * size; x++)
        {
            double f = bytes[(y % size) * size + x % size] / 128.0;
            byte Channel(double c) => (byte)Math.Clamp(Math.Round(255 + (c - 255) * f), 0, 255);
            int at = 3 * (y * 2 * size + x);
            output[at] = Channel(color.R);
            output[at + 1] = Channel(color.G);
            output[at + 2] = Channel(color.B);
        }
    }

    Png.Write(preview, 2 * size, 2 * size, output);
}

return 0;

// the density divided by its own wide blur (a box blur three times, about a Gaussian), times its mean
static double[] Flatten(double[] values, int side, int radius)
{
    if (radius <= 0)
    {
        return values;
    }

    var blurred = values;
    for (int pass = 0; pass < 3; pass++)
    {
        blurred = BoxBlur(blurred, side, radius, horizontal: true);
        blurred = BoxBlur(blurred, side, radius, horizontal: false);
    }

    double mean = values.Average();
    return values.Select((v, i) => v / Math.Max(1e-6, blurred[i]) * mean).ToArray();
}

static double[] BoxBlur(double[] values, int side, int radius, bool horizontal)
{
    var result = new double[values.Length];
    for (int line = 0; line < side; line++)
    {
        for (int i = 0; i < side; i++)
        {
            double sum = 0;
            int count = 0;
            for (int k = Math.Max(0, i - radius); k <= Math.Min(side - 1, i + radius); k++)
            {
                sum += horizontal ? values[line * side + k] : values[k * side + line];
                count++;
            }

            result[horizontal ? line * side + i : i * side + line] = sum / count;
        }
    }

    return result;
}

// A patch of side + overlap made into a tile of side that repeats: the first `overlap` columns
// fade from the columns just past the tile's end into their own, so that the last column runs
// on into the first; the blend keeps the grain's contrast (two noises averaged are flatter, and
// the sum is divided by the length of the weights)
static double[] Wrap(double[] values, int patch, int side, int overlap)
{
    double mean = values.Average();
    var columns = new double[patch * side];
    for (int y = 0; y < patch; y++)
    {
        for (int x = 0; x < side; x++)
        {
            columns[y * side + x] = x < overlap
                ? Blend(values[y * patch + x + side], values[y * patch + x], (double)x / overlap, mean)
                : values[y * patch + x];
        }
    }

    var tile = new double[side * side];
    for (int y = 0; y < side; y++)
    {
        for (int x = 0; x < side; x++)
        {
            tile[y * side + x] = y < overlap
                ? Blend(columns[(y + side) * side + x], columns[y * side + x], (double)y / overlap, mean)
                : columns[y * side + x];
        }
    }

    return tile;
}

static double Blend(double from, double to, double t, double mean)
{
    double w = t * t * (3 - 2 * t);
    double a = 1 - w;
    return mean + ((from - mean) * a + (to - mean) * w) / Math.Sqrt(a * a + w * w);
}

string Option(string name)
{
    int index = Array.IndexOf(args, name);
    return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
}

/// <summary>The little of PNG this needs: 8-bit RGB or RGBA, not interlaced, read as RGB; written as RGB</summary>
static class Png
{
    public static (int Width, int Height, byte[] Rgb) Read(string path)
    {
        using var file = File.OpenRead(path);
        var reader = new BinaryReader(file);
        reader.ReadBytes(8);
        int width = 0, height = 0, channels = 0;
        var compressed = new MemoryStream();
        while (file.Position < file.Length)
        {
            int length = BigEndian(reader.ReadBytes(4));
            string type = Encoding.ASCII.GetString(reader.ReadBytes(4));
            byte[] chunk = reader.ReadBytes(length);
            reader.ReadBytes(4);
            if (type == "IHDR")
            {
                width = BigEndian(chunk[0..4]);
                height = BigEndian(chunk[4..8]);
                if (chunk[8] != 8 || (chunk[9] != 2 && chunk[9] != 6) || chunk[12] != 0)
                {
                    throw new Exception("only 8-bit RGB or RGBA PNGs without interlacing");
                }

                channels = chunk[9] == 6 ? 4 : 3;
            }
            else if (type == "IDAT")
            {
                compressed.Write(chunk);
            }
            else if (type == "IEND")
            {
                break;
            }
        }

        compressed.Position = 0;
        var raw = new MemoryStream();
        using (var inflate = new ZLibStream(compressed, CompressionMode.Decompress))
        {
            inflate.CopyTo(raw);
        }

        byte[] data = raw.ToArray();
        int stride = width * channels;
        var previous = new byte[stride];
        var rgb = new byte[width * height * 3];
        for (int y = 0; y < height; y++)
        {
            int filter = data[y * (stride + 1)];
            var line = data.AsSpan(y * (stride + 1) + 1, stride).ToArray();
            for (int i = 0; i < stride; i++)
            {
                int left = i >= channels ? line[i - channels] : 0;
                int up = previous[i];
                int upLeft = i >= channels ? previous[i - channels] : 0;
                line[i] = (byte)(line[i] + filter switch
                {
                    1 => left,
                    2 => up,
                    3 => (left + up) / 2,
                    4 => Paeth(left, up, upLeft),
                    _ => 0,
                });
            }

            for (int x = 0; x < width; x++)
            {
                rgb[3 * (y * width + x)] = line[x * channels];
                rgb[3 * (y * width + x) + 1] = line[x * channels + 1];
                rgb[3 * (y * width + x) + 2] = line[x * channels + 2];
            }

            previous = line;
        }

        return (width, height, rgb);
    }

    public static void Write(string path, int width, int height, byte[] rgb)
    {
        var raw = new MemoryStream();
        for (int y = 0; y < height; y++)
        {
            raw.WriteByte(0);
            raw.Write(rgb, y * width * 3, width * 3);
        }

        var compressed = new MemoryStream();
        using (var deflate = new ZLibStream(compressed, CompressionLevel.Optimal, leaveOpen: true))
        {
            raw.Position = 0;
            raw.CopyTo(deflate);
        }

        using var file = File.Create(path);
        file.Write([137, 80, 78, 71, 13, 10, 26, 10]);
        var header = new List<byte>();
        header.AddRange(BigEndianBytes(width));
        header.AddRange(BigEndianBytes(height));
        header.AddRange(new byte[] { 8, 2, 0, 0, 0 });
        Chunk(file, "IHDR", header.ToArray());
        Chunk(file, "IDAT", compressed.ToArray());
        Chunk(file, "IEND", []);
    }

    static void Chunk(Stream file, string type, byte[] data)
    {
        file.Write(BigEndianBytes(data.Length));
        var typed = Encoding.ASCII.GetBytes(type).Concat(data).ToArray();
        file.Write(typed);
        file.Write(BigEndianBytes((int)Crc(typed)));
    }

    static uint Crc(byte[] data)
    {
        uint crc = 0xFFFFFFFF;
        foreach (byte b in data)
        {
            crc ^= b;
            for (int k = 0; k < 8; k++)
            {
                crc = (crc & 1) != 0 ? 0xEDB88320 ^ (crc >> 1) : crc >> 1;
            }
        }

        return crc ^ 0xFFFFFFFF;
    }

    static int Paeth(int a, int b, int c)
    {
        int p = a + b - c, pa = Math.Abs(p - a), pb = Math.Abs(p - b), pc = Math.Abs(p - c);
        return pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
    }

    static int BigEndian(byte[] bytes) => (bytes[0] << 24) | (bytes[1] << 16) | (bytes[2] << 8) | bytes[3];

    static byte[] BigEndianBytes(int value) => [(byte)(value >> 24), (byte)(value >> 16), (byte)(value >> 8), (byte)value];

    public static string FindRepository()
    {
        for (var folder = new DirectoryInfo(Directory.GetCurrentDirectory()); folder != null; folder = folder.Parent)
        {
            if (File.Exists(Path.Combine(folder.FullName, "AGENTS.md")) && Directory.Exists(Path.Combine(folder.FullName, "Main")))
            {
                return folder.FullName;
            }
        }

        throw new Exception("run this from inside the repository");
    }
}
