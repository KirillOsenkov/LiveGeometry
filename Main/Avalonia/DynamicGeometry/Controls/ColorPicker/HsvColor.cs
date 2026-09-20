using System;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// Hue (0..360), saturation, value and alpha (0..1). The picker works in this space so that
/// dragging to black or gray doesn't lose the hue the way a round trip through RGB would.
/// </summary>
public readonly struct HsvColor : IEquatable<HsvColor>
{
    public HsvColor(double hue, double saturation, double value, double alpha = 1)
    {
        Hue = ((hue % 360) + 360) % 360;
        Saturation = System.Math.Clamp(saturation, 0, 1);
        Value = System.Math.Clamp(value, 0, 1);
        Alpha = System.Math.Clamp(alpha, 0, 1);
    }

    public double Hue { get; }
    public double Saturation { get; }
    public double Value { get; }
    public double Alpha { get; }

    public HsvColor WithHue(double hue) => new HsvColor(hue, Saturation, Value, Alpha);

    public HsvColor WithSaturationAndValue(double saturation, double value) => new HsvColor(Hue, saturation, value, Alpha);

    public HsvColor WithAlpha(double alpha) => new HsvColor(Hue, Saturation, Value, alpha);

    public static HsvColor FromColor(Color color)
    {
        double red = color.R / 255.0;
        double green = color.G / 255.0;
        double blue = color.B / 255.0;
        double max = System.Math.Max(red, System.Math.Max(green, blue));
        double min = System.Math.Min(red, System.Math.Min(green, blue));
        double delta = max - min;

        double hue = 0;
        if (delta > 0)
        {
            if (max == red)
            {
                hue = 60 * (((green - blue) / delta) % 6);
            }
            else if (max == green)
            {
                hue = 60 * (((blue - red) / delta) + 2);
            }
            else
            {
                hue = 60 * (((red - green) / delta) + 4);
            }
        }

        double saturation = max == 0 ? 0 : delta / max;
        return new HsvColor(hue, saturation, max, color.A / 255.0);
    }

    public Color ToColor()
    {
        double chroma = Value * Saturation;
        double sector = Hue / 60;
        double second = chroma * (1 - System.Math.Abs((sector % 2) - 1));
        double match = Value - chroma;

        (double red, double green, double blue) = (int)sector switch
        {
            0 => (chroma, second, 0.0),
            1 => (second, chroma, 0.0),
            2 => (0.0, chroma, second),
            3 => (0.0, second, chroma),
            4 => (second, 0.0, chroma),
            _ => (chroma, 0.0, second),
        };

        return Color.FromArgb(
            ToByte(Alpha),
            ToByte(red + match),
            ToByte(green + match),
            ToByte(blue + match));
    }

    /// <summary>The fully saturated, fully bright color of this hue.</summary>
    public Color PureHue => new HsvColor(Hue, 1, 1).ToColor();

    static byte ToByte(double component) => (byte)System.Math.Round(System.Math.Clamp(component, 0, 1) * 255);

    public bool Equals(HsvColor other) =>
        Hue == other.Hue && Saturation == other.Saturation && Value == other.Value && Alpha == other.Alpha;

    public override bool Equals(object obj) => obj is HsvColor other && Equals(other);

    public override int GetHashCode() => HashCode.Combine(Hue, Saturation, Value, Alpha);
}
