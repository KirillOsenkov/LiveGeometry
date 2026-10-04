using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// A hue of the styles a drawing starts with (<see cref="StyleManager.AddDefaultStyles"/>).
/// Every kind of figure has a column per hue in its style picker, in the order of
/// <see cref="All"/>, rows of eight: a hue means the same for points, lines and shapes, and
/// the picker of a circle, which offers lines and shapes, lines up. A hue is a stroke for
/// each paper, the light one and the dark one; its fills are worked out from the stroke
/// (<see cref="Tint"/>), so that the hues stay about as light as each other.
/// </summary>
public class StyleHue
{
    /// <param name="fillHue">The hue of the fills on the light paper, when it isn't the stroke's (0 to 360)</param>
    /// <param name="darkFillHue">The same on the dark paper</param>
    public StyleHue(
        string name,
        string stroke,
        string darkStroke,
        double? fillHue = null,
        double? darkFillHue = null)
    {
        Name = name;
        Stroke = Color.Parse(stroke);
        DarkStroke = Color.Parse(darkStroke);
        this.fillHue = fillHue;
        this.darkFillHue = darkFillHue;
    }

    readonly double? fillHue;
    readonly double? darkFillHue;

    public static StyleHue Gray { get; } = new StyleHue("Gray", "#6E6E6E", "#9E9E9E");

    /// <summary>
    /// A golden brown outline, filled as the ribbon's shapes are, cream to gold: the column
    /// of the yellow point and the yellow fill of a new polygon. A fill of the outline's own
    /// hue was too near orange's; on the dark paper that hue went olive in the deep end.
    /// </summary>
    public static StyleHue Brown { get; } = new StyleHue("Brown", "#A87B05", "#E3B341", fillHue: 52, darkFillHue: 45);

    /// <summary>The column of the green point and the classic green fill</summary>
    public static StyleHue Green { get; } = new StyleHue("Green", "#2E9E4F", "#4CD27A");

    public static IReadOnlyList<StyleHue> All { get; } = new[]
    {
        Gray,
        new StyleHue("Red", "#D83B3B", "#FF6B6B"),
        new StyleHue("Orange", "#E08A00", "#FFA940"),
        Brown,
        Green,
        new StyleHue("Cyan", "#00A3BF", "#33D1EB"),
        new StyleHue("Blue", "#2F7BD6", "#6AA5FF"),
        new StyleHue("Purple", "#8854D0", "#B48CFF")
    };

    /// <summary>The hues but gray, whose styles are the defaults of the theme (the line, the thick line)</summary>
    public static IEnumerable<StyleHue> Colors => All.Where(hue => hue != Gray);

    public string Name { get; }

    /// <summary>A line or an outline on the light paper</summary>
    public Color Stroke { get; }

    /// <summary>A line or an outline on the dark paper</summary>
    public Color DarkStroke { get; }

    /// <summary>How much of the paper and what is under a shape shows through its fill</summary>
    public static byte FillAlpha = 0xC8;

    /// <summary>
    /// A shape's fill on the light paper: light at the top left, deeper at the bottom right,
    /// as the ribbon's shapes are filled
    /// </summary>
    public Brush Fill => Gradient(LightFillEnd, Tint(lightness: 0.78, FillAlpha));

    /// <summary>
    /// A shape's fill on the dark paper: bright at the top left, deep at the bottom right, at
    /// the stroke's full saturation (capped, orange and gold came out brown and olive)
    /// </summary>
    public Brush DarkFill => Gradient(DarkFillLightEnd, DarkTint(lightness: 0.17, FillAlpha, maxSaturation: 1));

    /// <summary>A flat fill on the light paper: the lighter end of <see cref="Fill"/></summary>
    public Brush SolidFill => new SolidColorBrush(LightFillEnd);

    /// <summary>A flat fill on the dark paper: the lighter end of <see cref="DarkFill"/></summary>
    public Brush DarkSolidFill => new SolidColorBrush(DarkFillLightEnd);

    Color LightFillEnd => Tint(lightness: 0.97, FillAlpha);

    Color DarkFillLightEnd => DarkTint(lightness: 0.6, FillAlpha, maxSaturation: 1);

    /// <summary>A bead (a big point) on the light paper: a light highlight running into the hue</summary>
    public Brush BeadFill => Gradient(Tint(lightness: 0.9), Tint(lightness: 0.55));

    public Brush DarkBeadFill => Gradient(DarkTint(lightness: 0.88), DarkTint(lightness: 0.42));

    /// <summary>The rim of a bead, a shade deeper than its fill</summary>
    public Color BeadRim => Tint(lightness: 0.3);

    public Color DarkBeadRim => DarkTint(lightness: 0.25);

    /// <summary>
    /// Light tints of the strongest hues (cyan, orange) glare next to the others: a tint is
    /// at most this saturated, unless asked otherwise
    /// </summary>
    const double MaxTintSaturation = 0.75;

    /// <summary>For the light paper: the stroke's color in the fill hue, as light as asked (0 black, 1 white), and not too saturated</summary>
    Color Tint(double lightness, byte alpha = 0xFF, double maxSaturation = MaxTintSaturation)
    {
        return Tint(Stroke, fillHue, lightness, alpha, maxSaturation);
    }

    /// <summary>The same for the dark paper, from the dark stroke</summary>
    Color DarkTint(double lightness, byte alpha = 0xFF, double maxSaturation = MaxTintSaturation)
    {
        return Tint(DarkStroke, darkFillHue, lightness, alpha, maxSaturation);
    }

    static Color Tint(
        Color color,
        double? hue,
        double lightness,
        byte alpha,
        double maxSaturation)
    {
        var hsl = color.ToHsl();
        var result = new HslColor(alpha: 1, hue ?? hsl.H, System.Math.Min(hsl.S, maxSaturation), lightness).ToRgb();
        return Color.FromArgb(alpha, result.R, result.G, result.B);
    }

    /// <summary>A diagonal gradient across the box of whatever it fills, top left to bottom right</summary>
    public static Brush Gradient(Color from, Color to)
    {
        return new LinearGradientBrush()
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(1, 1, RelativeUnit.Relative),
            GradientStops =
            {
                new GradientStop(from, offset: 0),
                new GradientStop(to, offset: 1)
            }
        };
    }
}
