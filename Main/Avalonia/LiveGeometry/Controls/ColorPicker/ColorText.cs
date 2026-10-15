using System.Globalization;
using Avalonia;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// Colors as text (names and hex codes), and the checkerboard that stands for transparency.
/// </summary>
public static class ColorText
{
    /// <summary>#RRGGBB, or #AARRGGBB when the color isn't opaque</summary>
    public static string ToHex(Color color)
    {
        return color.A == 255
            ? $"#{color.R:X2}{color.G:X2}{color.B:X2}"
            : $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";
    }

    /// <summary>
    /// #AARRGGBB, always. This is how colors go into drawing files (it is what WPF's
    /// Color.ToString() produced); Avalonia's Color.ToString() is no good for that because
    /// it substitutes the name for a known color.
    /// </summary>
    public static string ToArgbHex(Color color)
    {
        return $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";
    }

    /// <summary>"Red #FF0000" for a named color, just the hex code otherwise</summary>
    public static string Describe(Color color)
    {
        var name = ColorPalette.GetName(color);
        return name != null ? name + " " + ToHex(color) : ToHex(color);
    }

    /// <summary>
    /// Accepts a web color name, RGB, RRGGBB or AARRGGBB hex with or without the #, or a
    /// "Name #hex" pair as produced by <see cref="Describe"/>.
    /// </summary>
    public static bool TryParse(string text, out Color color)
    {
        color = default;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        text = text.Trim();
        int space = text.IndexOf(' ');
        if (space > 0)
        {
            return TryParse(text.Substring(space + 1), out color) || TryParse(text.Substring(0, space), out color);
        }

        if (ColorPalette.ColorsByName.TryGetValue(text, out color))
        {
            return true;
        }

        var hex = text.TrimStart('#');
        if (hex.Length == 3)
        {
            hex = $"{hex[0]}{hex[0]}{hex[1]}{hex[1]}{hex[2]}{hex[2]}";
        }

        if (hex.Length == 6)
        {
            hex = "FF" + hex;
        }

        if (hex.Length == 8 && uint.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out uint argb))
        {
            color = Color.FromUInt32(argb);
            return true;
        }

        return false;
    }

    static IBrush checkerboardBrush;

    /// <summary>Light gray and white squares, to show through transparent colors.</summary>
    public static IBrush CheckerboardBrush
    {
        get
        {
            if (checkerboardBrush == null)
            {
                const double cell = 4;
                var squares = new GeometryGroup();
                squares.Children.Add(new RectangleGeometry(new Rect(0, 0, cell, cell)));
                squares.Children.Add(new RectangleGeometry(new Rect(cell, cell, cell, cell)));
                var drawing = new DrawingGroup();
                drawing.Children.Add(new GeometryDrawing()
                {
                    Brush = Brushes.White,
                    Geometry = new RectangleGeometry(new Rect(0, 0, 2 * cell, 2 * cell))
                });
                drawing.Children.Add(new GeometryDrawing()
                {
                    Brush = new SolidColorBrush(Color.FromRgb(0xD0, 0xD0, 0xD0)),
                    Geometry = squares
                });
                checkerboardBrush = new DrawingBrush(drawing)
                {
                    TileMode = TileMode.Tile,
                    Stretch = Stretch.None,
                    DestinationRect = new RelativeRect(0, 0, 2 * cell, 2 * cell, RelativeUnit.Absolute)
                };
            }

            return checkerboardBrush;
        }
    }
}
