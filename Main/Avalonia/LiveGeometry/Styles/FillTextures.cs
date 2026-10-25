using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Avalonia;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform;

namespace LiveGeometry;

/// <summary>How a shape's fill is laid down (<see cref="ShapeStyle.Texture"/>)</summary>
public enum FillTexture
{
    None,

    /// <summary>With a colored pencil: the paper's tooth shows through in flecks</summary>
    Crayon,

    /// <summary>With a felt-tip pen: smooth ink, soft blotches and the bands of the strokes</summary>
    FeltTip,
}

/// <summary>
/// Fills that look laid down by hand rather than poured. Each texture's tile is a scan of a
/// hand-colored area (<see cref="CrayonTile"/>, <see cref="FeltTipTile"/>, written by
/// tools/texturetile.cs): how much pigment each pixel has against the average. A fill of a color
/// is the paper moved toward the color by that much, so it averages to the color, its pale spots
/// fade to the paper and where the pen or pencil went over again it deepens. Opaque on the paper:
/// a pale spot shows the paper, not whatever is under the shape, as each region of a hand-colored
/// drawing is colored on the paper itself. The player has the twin
/// (Main/Player/src/styles/fillTextures.js).
/// </summary>
public static class FillTextures
{
    /// <summary>How much of the canvas a pixel of a tile covers: a pixel each (drawn smaller, it blurred and lost its grain)</summary>
    const double TileScale = 1;

    static readonly Dictionary<FillTexture, byte[]> densities = new();
    static readonly Dictionary<(FillTexture Texture, Color Color, Color Paper), IBrush> brushes = new();

    /// <summary>The fill of the color on the paper, laid down with the texture</summary>
    public static IBrush Brush(FillTexture texture, Color color, Color paper)
    {
        if (brushes.TryGetValue((texture, color, paper), out var brush))
        {
            return brush;
        }

        var (size, data) = texture == FillTexture.FeltTip ? (FeltTipTile.Size, FeltTipTile.Data) : (CrayonTile.Size, CrayonTile.Data);
        if (!densities.TryGetValue(texture, out var density))
        {
            density = Convert.FromBase64String(data);
            densities[texture] = density;
        }

        var pixels = new byte[size * size * 4];
        // a translucent color is that much less of a step from the paper
        double strength = color.A / 255.0;
        for (int i = 0; i < density.Length; i++)
        {
            double amount = density[i] / 128.0 * strength;
            // blue green red alpha
            pixels[4 * i] = Channel(paper.B, color.B, amount);
            pixels[4 * i + 1] = Channel(paper.G, color.G, amount);
            pixels[4 * i + 2] = Channel(paper.R, color.R, amount);
            pixels[4 * i + 3] = 255;
        }

        var bitmap = new WriteableBitmap(new PixelSize(size, size), new Avalonia.Vector(96, 96), PixelFormat.Bgra8888, AlphaFormat.Premul);
        using (var buffer = bitmap.Lock())
        {
            for (int row = 0; row < size; row++)
            {
                Marshal.Copy(pixels, row * size * 4, buffer.Address + row * buffer.RowBytes, size * 4);
            }
        }

        brush = new ImageBrush(bitmap)
        {
            TileMode = TileMode.Tile,
            Stretch = Stretch.Fill,
            AlignmentX = AlignmentX.Left,
            AlignmentY = AlignmentY.Top,
            DestinationRect = new RelativeRect(0, 0, size * TileScale, size * TileScale, RelativeUnit.Absolute),
        };
        if (brushes.Count > 200)
        {
            brushes.Clear();
        }

        brushes[(texture, color, paper)] = brush;
        return brush;
    }

    static byte Channel(byte paper, byte color, double amount) => (byte)System.Math.Clamp(System.Math.Round(paper + (color - paper) * amount), 0, 255);

    /// <summary>The one color of a fill: a solid brush's, the middle of a gradient's stops</summary>
    public static Color ColorOf(IBrush brush)
    {
        if (brush is ISolidColorBrush solid)
        {
            return solid.Color;
        }

        if (brush is IGradientBrush gradient && gradient.GradientStops.Count > 0)
        {
            var first = gradient.GradientStops[0].Color;
            var last = gradient.GradientStops[gradient.GradientStops.Count - 1].Color;
            return Color.FromArgb(
                (byte)System.Math.Round((first.A + last.A) / 2.0),
                (byte)System.Math.Round((first.R + last.R) / 2.0),
                (byte)System.Math.Round((first.G + last.G) / 2.0),
                (byte)System.Math.Round((first.B + last.B) / 2.0));
        }

        return Colors.Transparent;
    }
}
