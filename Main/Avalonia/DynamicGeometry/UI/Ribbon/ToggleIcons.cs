using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// Icons of the drawing-aid toggles (the on/off state is shown by the button itself).
/// </summary>
public static class ToggleIcons
{
    static readonly Color guide = Color.FromRgb(0x8A, 0x94, 0xA6);
    static readonly Color accent = Color.FromRgb(0x2F, 0x7F, 0xD8);

    /// <summary>Movement constrained to horizontal/vertical</summary>
    public static FrameworkElement Ortho()
    {
        return IconBuilder.BuildIcon()
            .Line(strokeThickness: 2, accent, 0.25, 0.2, 0.25, 0.75)
            .Line(strokeThickness: 2, accent, 0.25, 0.75, 0.8, 0.75)
            .Line(guide, 0.25, 0.6, 0.4, 0.6)
            .Line(guide, 0.4, 0.6, 0.4, 0.75)
            .Point(0.25, 0.75)
            .Canvas;
    }

    /// <summary>Movement constrained to fixed angle steps</summary>
    public static FrameworkElement Polar()
    {
        return IconBuilder.BuildIcon()
            .Line(guide, 0.2, 0.78, 0.85, 0.78)
            .Line(guide, 0.2, 0.78, 0.76, 0.46)
            .Line(strokeThickness: 2, accent, 0.2, 0.78, 0.58, 0.2)
            .Arc(0.2, 0.78, 0.55, 0.78, 0.39, 0.49)
            .Point(0.2, 0.78)
            .Canvas;
    }

    public static FrameworkElement SnapToGrid()
    {
        var builder = IconBuilder.BuildIcon();
        foreach (var position in new[] { 0.2, 0.5, 0.8 })
        {
            builder.Line(guide, position, 0.1, position, 0.9);
            builder.Line(guide, 0.1, position, 0.9, position);
        }

        return builder.Point(0.5, 0.5).Canvas;
    }

    public static FrameworkElement SnapToPoint()
    {
        return IconBuilder.BuildIcon()
            .Circle(0.5, 0.5, radius: 0.32)
            .Line(accent, 0.5, 0.06, 0.5, 0.3)
            .Line(accent, 0.5, 0.7, 0.5, 0.94)
            .Line(accent, 0.06, 0.5, 0.3, 0.5)
            .Line(accent, 0.7, 0.5, 0.94, 0.5)
            .Point(0.5, 0.5)
            .Canvas;
    }

    public static FrameworkElement SnapToCenter()
    {
        return IconBuilder.BuildIcon()
            .Line(0.12, 0.72, 0.88, 0.28)
            .Line(accent, 0.44, 0.38, 0.56, 0.62)
            .DependentPoint(0.12, 0.72)
            .DependentPoint(0.88, 0.28)
            .Point(0.5, 0.5)
            .Canvas;
    }

    public static FrameworkElement LabelNewPoints()
    {
        return IconBuilder.BuildIcon()
            .Point(0.36, 0.64)
            .Text(Colors.Black, 0.5, 0.08, text: "A")
            .Canvas;
    }
}
