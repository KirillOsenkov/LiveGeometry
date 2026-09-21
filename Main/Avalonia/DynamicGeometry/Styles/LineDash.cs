using Avalonia.Collections;

namespace DynamicGeometry;

/// <summary>
/// The same five the original DG had (VB6 DrawStyle 0-4), in the same order.
/// The names are what gets written into drawings.
/// </summary>
public enum LineDash
{
    Solid,
    Dash,
    Dot,
    DashDot,
    DashDotDot
}

public static class LineDashes
{
    /// <summary>
    /// Dash and gap lengths in pixels. A pattern looks the same on a thin and on a thick line,
    /// only a thick line stretches it enough for a gap to stay a gap.
    /// </summary>
    static double[] GetPattern(LineDash dash)
    {
        switch (dash)
        {
            case LineDash.Dash:
                return new double[] { 8, 5 };
            case LineDash.Dot:
                return new double[] { 2, 4 };
            case LineDash.DashDot:
                return new double[] { 9, 4, 2, 4 };
            case LineDash.DashDotDot:
                return new double[] { 9, 4, 2, 4, 2, 4 };
            default:
                return null;
        }
    }

    /// <summary>
    /// What to put into Shape.StrokeDashArray, which counts in stroke widths; null for solid.
    /// </summary>
    public static AvaloniaList<double> GetDashArray(LineDash dash, double strokeThickness)
    {
        var pattern = GetPattern(dash);
        if (pattern == null)
        {
            return null;
        }

        // up to 3 px wide the lengths are as given; beyond that they grow with the line
        var unit = System.Math.Max(strokeThickness, 0.1);
        var stretch = System.Math.Max(1, unit / 3);
        var result = new AvaloniaList<double>();
        foreach (var length in pattern)
        {
            result.Add(length * stretch / unit);
        }

        return result;
    }

    /// <summary>The style of a line in a .dgf file of the original DG</summary>
    public static LineDash FromVB6DrawStyle(int drawStyle)
    {
        return drawStyle >= 1 && drawStyle <= 4 ? (LineDash)drawStyle : LineDash.Solid;
    }
}
