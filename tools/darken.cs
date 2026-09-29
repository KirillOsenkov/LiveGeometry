#:property TargetFramework=net10.0
#:property Nullable=disable
#:property PublishAot=false

// darken - gives the styles of the .lgf files in a folder a Dark override where their color
// would sink into the dark paper: a dark stroke, text or fill (lightness under 0.35) is
// lightened, keeping its hue and its alpha, into a <Dark .../> child of the style. Chromatic
// colors that read on both papers are left alone, and so is a drawing with a paper of its
// own (the Castle's sky, the Spiral's dark gray) and a style that has a Dark override
// already. Lists what it would do; --apply writes it.
//
//   dotnet tools/darken.cs -- <folder of .lgf> [--apply]

using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;

if (args.Length < 1)
{
    Console.WriteLine("usage: darken <folder of .lgf> [--apply]");
    return 1;
}

bool apply = args.Contains("--apply");
const double darkLightness = 0.35;
int changedFiles = 0;

foreach (var file in Directory.GetFiles(args[0], "*.lgf").OrderBy(f => f))
{
    var document = XDocument.Load(file);
    var viewport = document.Root.Element("Viewport");
    if (viewport?.Attribute("Color") != null || viewport?.Element("Background") != null)
    {
        Console.WriteLine($"{Path.GetFileName(file)}: paper of its own, left alone");
        continue;
    }

    bool changed = false;
    foreach (var style in document.Root.Element("Styles")?.Elements() ?? Enumerable.Empty<XElement>())
    {
        if (style.Element("Dark") != null)
        {
            continue;
        }

        var overrides = new List<XAttribute>();
        foreach (var attributeName in new[] { "Color", "Fill" })
        {
            var attribute = style.Attribute(attributeName);
            if (attribute == null || !TryParse(attribute.Value, out var color))
            {
                continue;
            }

            if (color.A < 0x20)
            {
                continue;
            }

            var lightened = Lighten(color);
            if (lightened != color)
            {
                overrides.Add(new XAttribute(attributeName, Format(lightened)));
            }
        }

        if (overrides.Count == 0)
        {
            continue;
        }

        Console.WriteLine($"{Path.GetFileName(file)}: {style.Name.LocalName} {(string)style.Attribute("Name")}: {string.Join(" ", overrides.Select(a => a.Name + "=" + a.Value))}");
        style.Add(new XElement("Dark", overrides));
        changed = true;
    }

    if (changed && apply)
    {
        var settings = new XmlWriterSettings() { Indent = true, Encoding = new UTF8Encoding(false), NewLineChars = "\r\n" };
        using (var writer = XmlWriter.Create(file, settings))
        {
            document.Save(writer);
        }

        changedFiles++;
    }
}

Console.WriteLine(apply ? $"darkened {changedFiles} files" : "dry run: add --apply to write");
return 0;

static bool TryParse(string text, out (byte A, byte R, byte G, byte B) color)
{
    color = default;
    if (text == null || !text.StartsWith("#") || text.Length != 9)
    {
        return false;
    }

    color = (
        byte.Parse(text.Substring(1, 2), NumberStyles.HexNumber),
        byte.Parse(text.Substring(3, 2), NumberStyles.HexNumber),
        byte.Parse(text.Substring(5, 2), NumberStyles.HexNumber),
        byte.Parse(text.Substring(7, 2), NumberStyles.HexNumber));
    return true;
}

static string Format((byte A, byte R, byte G, byte B) color)
{
    return $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";
}

// a dark color to the same hue and saturation at a lightness that reads on dark paper:
// black to 0.82 (the theme's ink is there), and a color that was already lighter to a
// little less than that, so that two dark colors stay apart
static (byte A, byte R, byte G, byte B) Lighten((byte A, byte R, byte G, byte B) color)
{
    ToHsl(color, out double hue, out double saturation, out double lightness);
    if (lightness >= darkLightness)
    {
        return color;
    }

    return FromHsl(color.A, hue, saturation, 0.82 - 0.4 * lightness);
}

static void ToHsl((byte A, byte R, byte G, byte B) color, out double hue, out double saturation, out double lightness)
{
    double r = color.R / 255.0, g = color.G / 255.0, b = color.B / 255.0;
    double max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b));
    lightness = (max + min) / 2;
    double delta = max - min;
    if (delta == 0)
    {
        hue = 0;
        saturation = 0;
        return;
    }

    saturation = delta / (1 - Math.Abs(2 * lightness - 1));
    if (max == r)
    {
        hue = 60 * (((g - b) / delta) % 6);
    }
    else if (max == g)
    {
        hue = 60 * ((b - r) / delta + 2);
    }
    else
    {
        hue = 60 * ((r - g) / delta + 4);
    }

    if (hue < 0)
    {
        hue += 360;
    }
}

static (byte A, byte R, byte G, byte B) FromHsl(byte alpha, double hue, double saturation, double lightness)
{
    double chroma = (1 - Math.Abs(2 * lightness - 1)) * saturation;
    double x = chroma * (1 - Math.Abs(hue / 60 % 2 - 1));
    double m = lightness - chroma / 2;
    (double r, double g, double b) = ((int)(hue / 60)) switch
    {
        0 => (chroma, x, 0.0),
        1 => (x, chroma, 0.0),
        2 => (0.0, chroma, x),
        3 => (0.0, x, chroma),
        4 => (x, 0.0, chroma),
        _ => (chroma, 0.0, x)
    };
    return (alpha, Channel(r + m), Channel(g + m), Channel(b + m));

    static byte Channel(double value)
    {
        return (byte)Math.Clamp(Math.Round(value * 255), 0, 255);
    }
}
