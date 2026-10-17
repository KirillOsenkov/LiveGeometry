using System.Runtime.CompilerServices;

namespace LiveGeometry;

/// <summary>
/// The length and the direction last typed for an end (<see cref="LengthConstraint.SetDistance"/>,
/// <see cref="DirectionConstraint.SetDirection(IPoint, IPoint, double)"/>), so that fixing them
/// keeps the number typed: worked back out of the coordinates the end was moved to, a length
/// of 8 came back as 7.999999999999999, and the file saved that (2026-10-17). Only while the
/// end is still there but for rounding: dragged or set otherwise since, the end is fixed where
/// it is.
/// </summary>
public static class TypedDimensions
{
    /// <summary>Relative, of a distance: how far apart a typed value and the one worked back out of coordinates may be</summary>
    const double DistanceTolerance = 1e-9;

    /// <summary>In degrees</summary>
    const double DirectionTolerance = 1e-9;

    static readonly ConditionalWeakTable<IPoint, StrongBox<(IPoint Pivot, double Value)>> distances = new();
    static readonly ConditionalWeakTable<IPoint, StrongBox<(IPoint Pivot, double Value)>> directions = new();

    public static void RememberDistance(IPoint end, IPoint pivot, double distance)
    {
        distances.AddOrUpdate(end, new StrongBox<(IPoint, double)>((pivot, System.Math.Abs(distance))));
    }

    /// <summary>In degrees, from the pivot to the end</summary>
    public static void RememberDirection(IPoint end, IPoint pivot, double degrees)
    {
        directions.AddOrUpdate(end, new StrongBox<(IPoint, double)>((pivot, degrees)));
    }

    /// <summary>The distance as typed, with the sign of the current one, when the current one is that but for rounding; else the current one</summary>
    public static double Distance(IPoint end, IPoint pivot, double current)
    {
        if (distances.TryGetValue(end, out var typed) && typed.Value.Pivot == pivot)
        {
            double value = typed.Value.Value;
            if (System.Math.Abs(System.Math.Abs(current) - value) <= DistanceTolerance * System.Math.Max(1, value))
            {
                return current < 0 ? -value : value;
            }
        }

        return current;
    }

    /// <summary>The direction as typed, in degrees, as many turns from it as the current one is, when the current one is that but for rounding; else the current one</summary>
    public static double Direction(IPoint end, IPoint pivot, double current)
    {
        if (directions.TryGetValue(end, out var typed) && typed.Value.Pivot == pivot)
        {
            double value = typed.Value.Value;
            double sameTurn = value + 360 * System.Math.Round((current - value) / 360);
            if (System.Math.Abs(current - sameTurn) <= DirectionTolerance)
            {
                return sameTurn;
            }
        }

        return current;
    }
}
