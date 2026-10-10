namespace DynamicGeometry;

/// <summary>
/// Which of the two angles between two rays out of a point a figure means: an angle's mark
/// and number, a bisector, an arc with its sector and segment are all one of them. Let θ be
/// the counterclockwise angle from the first ray to the second, in [0, 2π): the regions
/// measure θ and 2π - θ. Counterclockwise and Clockwise are fixed choices, which never jump
/// and whose number runs on past 180° and wraps at 360°. Smaller and Larger are conditional:
/// they trade places at 180°, where the number bounces and a figure built on the region (a
/// bisector) turns round. Both are the right thing for an angle of a triangle, which stays
/// the inside one when the triangle flattens and gets reflected. At exactly 180° both
/// regions are the same size, and the counterclockwise one is taken.
/// </summary>
public enum AngleSweep
{
    Counterclockwise,
    Clockwise,

    [PropertyGridName("Under 180°")]
    Smaller,

    [PropertyGridName("Over 180°")]
    Larger
}

/// <summary>A figure that is one of the two angles between two rays (see <see cref="AngleSweep"/>)</summary>
public interface IHasSweep
{
    AngleSweep Sweep { get; set; }
}

public static class AngleSweepExtensions
{
    /// <summary>
    /// Whether the region goes clockwise from the first ray to the second, given the
    /// counterclockwise angle between them (0 to 2π)
    /// </summary>
    public static bool IsClockwise(this AngleSweep sweep, double counterclockwise)
    {
        switch (sweep)
        {
            case AngleSweep.Clockwise:
                return true;
            case AngleSweep.Smaller:
                return counterclockwise > Math.PI;
            case AngleSweep.Larger:
                return counterclockwise < Math.PI;
            default:
                return false;
        }
    }

    /// <summary>
    /// The measure of the region, 0 to 2π, given the counterclockwise angle between the
    /// rays. Two rays along one line measure 0 whichever way round (as <see cref="Math.OAngle"/>
    /// says a full turn is no turn), so an angle doesn't blink between 0° and 360°.
    /// </summary>
    public static double Measure(this AngleSweep sweep, double counterclockwise)
    {
        return sweep.IsClockwise(counterclockwise) && counterclockwise > 0
            ? 2 * Math.PI - counterclockwise
            : counterclockwise;
    }

    /// <summary>
    /// The sweep of the mirror image: a reflection turns a counterclockwise region
    /// clockwise, and leaves the conditional choices as they are (the image of the smaller
    /// angle is the smaller angle)
    /// </summary>
    public static AngleSweep Mirrored(this AngleSweep sweep)
    {
        switch (sweep)
        {
            case AngleSweep.Counterclockwise:
                return AngleSweep.Clockwise;
            case AngleSweep.Clockwise:
                return AngleSweep.Counterclockwise;
            default:
                return sweep;
        }
    }

    /// <summary>The attribute's value read back by name; the given default when it is missing or unknown</summary>
    public static AngleSweep ReadSweep(this System.Xml.Linq.XElement element, AngleSweep defaultSweep)
    {
        return System.Enum.TryParse(element.ReadString("Sweep"), out AngleSweep read) ? read : defaultSweep;
    }
}
