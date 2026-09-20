using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Avalonia.Media;

namespace DynamicGeometry;

public readonly record struct NamedColor(string Name, Color Color);

/// <summary>
/// A set of colors offered as swatches, already in the order and shape (number of columns)
/// they should be shown in. <see cref="SwatchPage"/> just lays them out row by row, so a
/// different palette - or a different arrangement of the same colors - is a new instance
/// of this class, not new UI code.
/// </summary>
public class ColorPalette
{
    public ColorPalette(IReadOnlyList<NamedColor> colors, int columns)
    {
        Colors = colors;
        Columns = columns;
    }

    /// <summary>Row by row</summary>
    public IReadOnlyList<NamedColor> Colors { get; }

    public int Columns { get; }

    static ColorPalette webColors;

    /// <summary>
    /// The named web colors (the ones everybody knows from HTML/CSS and System.Windows.Media.Colors)
    /// in the hand-arranged 14x10 map of the Helix color picker: a gray ramp down the first
    /// column, greens to blues left to right, a pale band across the middle with Transparent
    /// in it, and yellows, oranges, reds and purples below.
    /// </summary>
    public static ColorPalette WebColors => webColors ??= FromNames(columns: 14, @"
Black DarkSlateGray DarkGreen Green Olive DarkOliveGreen SeaGreen DarkSeaGreen LimeGreen YellowGreen MediumAquamarine LightSeaGreen DarkCyan Teal
DimGray SlateGray LightSlateGray ForestGreen DarkKhaki OliveDrab MediumSeaGreen Chartreuse LightGreen MediumSpringGreen Aquamarine Turquoise MediumTurquoise CadetBlue
Gray DarkGray Gainsboro LightGoldenrodYellow Beige GreenYellow LawnGreen Lime SpringGreen Honeydew PaleTurquoise LightBlue DarkTurquoise SteelBlue
Silver LightGray WhiteSmoke Snow White LemonChiffon PaleGreen Ivory FloralWhite MintCream LightCyan Cyan DeepSkyBlue CornflowerBlue
Khaki BlanchedAlmond Bisque Cornsilk Transparent OldLace White Azure GhostWhite AliceBlue PowderBlue LightSkyBlue DodgerBlue MediumBlue
PaleGoldenrod Wheat Moccasin LightYellow PapayaWhip PeachPuff AntiqueWhite SeaShell LavenderBlush Lavender LightSteelBlue SkyBlue RoyalBlue Blue
Yellow BurlyWood NavajoWhite Orange Coral DarkSalmon LightSalmon LightPink MistyRose Linen Thistle MediumSlateBlue SlateBlue DarkSlateBlue
Gold Tan SandyBrown DarkOrange Tomato Salmon LightCoral RosyBrown Violet Pink Plum BlueViolet MediumPurple DarkBlue
Goldenrod Peru Chocolate OrangeRed Firebrick Crimson IndianRed PaleVioletRed Orchid MediumOrchid DarkViolet DarkOrchid Indigo Navy
DarkGoldenrod SaddleBrown Sienna Red DarkRed Maroon Brown HotPink Magenta DeepPink MediumVioletRed DarkMagenta Purple MidnightBlue");

    static Dictionary<string, Color> colorsByName;

    /// <summary>Name to color, for the web colors plus Transparent; case-insensitive.</summary>
    public static IReadOnlyDictionary<string, Color> ColorsByName
    {
        get
        {
            if (colorsByName == null)
            {
                colorsByName = new Dictionary<string, Color>(StringComparer.OrdinalIgnoreCase);
                foreach (var named in ParseWebColors())
                {
                    colorsByName[named.Name] = named.Color;
                }

                colorsByName["Transparent"] = Avalonia.Media.Colors.Transparent;

                // the other names of Cyan and Magenta; Avalonia's Color.ToString() prefers these
                colorsByName["Aqua"] = colorsByName["Cyan"];
                colorsByName["Fuchsia"] = colorsByName["Magenta"];
            }

            return colorsByName;
        }
    }

    /// <returns>The name of the web color, or null if the color has none</returns>
    public static string GetName(Color color)
    {
        foreach (var pair in ColorsByName)
        {
            if (pair.Value == color)
            {
                return pair.Key;
            }
        }

        return null;
    }

    /// <summary>A palette from color names listed row by row, separated by whitespace.</summary>
    public static ColorPalette FromNames(int columns, string names)
    {
        var colors = names
            .Split(new[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(name => new NamedColor(name, ColorsByName[name]))
            .ToList();
        return new ColorPalette(colors, columns);
    }

    /// <summary>
    /// Lays colors out in columns of <paramref name="rows"/>: near-grays first (light to
    /// dark), then the rest sorted around the hue wheel and cut into columns, each column
    /// sorted light to dark.
    /// </summary>
    public static ColorPalette ArrangeByHue(IEnumerable<NamedColor> colors, int rows)
    {
        var all = colors.ToList();
        var grays = all
            .Where(c => HsvColor.FromColor(c.Color).Saturation < 0.12)
            .OrderByDescending(c => Lightness(c.Color))
            .ToList();
        var chromatic = all
            .Except(grays)
            .OrderBy(c => RotatedHue(c.Color))
            .ToList();

        var columns = new List<List<NamedColor>>();
        foreach (var source in new[] { grays, chromatic })
        {
            for (int start = 0; start < source.Count; start += rows)
            {
                columns.Add(source
                    .Skip(start)
                    .Take(rows)
                    .OrderByDescending(c => Lightness(c.Color))
                    .ToList());
            }
        }

        var rowByRow = new List<NamedColor>();
        for (int row = 0; row < rows; row++)
        {
            foreach (var column in columns)
            {
                // a short last column leaves holes; transparent fills them
                rowByRow.Add(row < column.Count ? column[row] : new NamedColor("", Avalonia.Media.Colors.Transparent));
            }
        }

        return new ColorPalette(rowByRow, columns.Count);
    }

    /// <summary>Hue with the wheel cut in the magentas instead of in the middle of the reds.</summary>
    static double RotatedHue(Color color) => (HsvColor.FromColor(color).Hue + 15) % 360;

    static double Lightness(Color color) => 0.2126 * color.R + 0.7152 * color.G + 0.0722 * color.B;

    /// <summary>
    /// Spelled out rather than reflected from Avalonia.Media.Colors: the trimmer of the browser
    /// build removes the color properties nobody references, and reflection would come up short.
    /// Aqua/Cyan and Fuchsia/Magenta are the same color; one name each is kept.
    /// </summary>
    static IEnumerable<NamedColor> ParseWebColors()
    {
        const string table = @"
AliceBlue F0F8FF AntiqueWhite FAEBD7 Aquamarine 7FFFD4 Azure F0FFFF Beige F5F5DC Bisque FFE4C4
Black 000000 BlanchedAlmond FFEBCD Blue 0000FF BlueViolet 8A2BE2 Brown A52A2A BurlyWood DEB887
CadetBlue 5F9EA0 Chartreuse 7FFF00 Chocolate D2691E Coral FF7F50 CornflowerBlue 6495ED
Cornsilk FFF8DC Crimson DC143C Cyan 00FFFF DarkBlue 00008B DarkCyan 008B8B DarkGoldenrod B8860B
DarkGray A9A9A9 DarkGreen 006400 DarkKhaki BDB76B DarkMagenta 8B008B DarkOliveGreen 556B2F
DarkOrange FF8C00 DarkOrchid 9932CC DarkRed 8B0000 DarkSalmon E9967A DarkSeaGreen 8FBC8F
DarkSlateBlue 483D8B DarkSlateGray 2F4F4F DarkTurquoise 00CED1 DarkViolet 9400D3 DeepPink FF1493
DeepSkyBlue 00BFFF DimGray 696969 DodgerBlue 1E90FF Firebrick B22222 FloralWhite FFFAF0
ForestGreen 228B22 Gainsboro DCDCDC GhostWhite F8F8FF Gold FFD700 Goldenrod DAA520 Gray 808080
Green 008000 GreenYellow ADFF2F Honeydew F0FFF0 HotPink FF69B4 IndianRed CD5C5C Indigo 4B0082
Ivory FFFFF0 Khaki F0E68C Lavender E6E6FA LavenderBlush FFF0F5 LawnGreen 7CFC00
LemonChiffon FFFACD LightBlue ADD8E6 LightCoral F08080 LightCyan E0FFFF
LightGoldenrodYellow FAFAD2 LightGray D3D3D3 LightGreen 90EE90 LightPink FFB6C1
LightSalmon FFA07A LightSeaGreen 20B2AA LightSkyBlue 87CEFA LightSlateGray 778899
LightSteelBlue B0C4DE LightYellow FFFFE0 Lime 00FF00 LimeGreen 32CD32 Linen FAF0E6
Magenta FF00FF Maroon 800000 MediumAquamarine 66CDAA MediumBlue 0000CD MediumOrchid BA55D3
MediumPurple 9370DB MediumSeaGreen 3CB371 MediumSlateBlue 7B68EE MediumSpringGreen 00FA9A
MediumTurquoise 48D1CC MediumVioletRed C71585 MidnightBlue 191970 MintCream F5FFFA
MistyRose FFE4E1 Moccasin FFE4B5 NavajoWhite FFDEAD Navy 000080 OldLace FDF5E6 Olive 808000
OliveDrab 6B8E23 Orange FFA500 OrangeRed FF4500 Orchid DA70D6 PaleGoldenrod EEE8AA
PaleGreen 98FB98 PaleTurquoise AFEEEE PaleVioletRed DB7093 PapayaWhip FFEFD5 PeachPuff FFDAB9
Peru CD853F Pink FFC0CB Plum DDA0DD PowderBlue B0E0E6 Purple 800080 Red FF0000
RosyBrown BC8F8F RoyalBlue 4169E1 SaddleBrown 8B4513 Salmon FA8072 SandyBrown F4A460
SeaGreen 2E8B57 SeaShell FFF5EE Sienna A0522D Silver C0C0C0 SkyBlue 87CEEB SlateBlue 6A5ACD
SlateGray 708090 Snow FFFAFA SpringGreen 00FF7F SteelBlue 4682B4 Tan D2B48C Teal 008080
Thistle D8BFD8 Tomato FF6347 Turquoise 40E0D0 Violet EE82EE Wheat F5DEB3 White FFFFFF
WhiteSmoke F5F5F5 Yellow FFFF00 YellowGreen 9ACD32";

        var tokens = table.Split(new[] { ' ', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        for (int i = 0; i + 1 < tokens.Length; i += 2)
        {
            uint rgb = uint.Parse(tokens[i + 1], NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            yield return new NamedColor(tokens[i], Color.FromUInt32(0xFF000000 | rgb));
        }
    }
}
