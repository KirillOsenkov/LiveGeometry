using GuiLabs.Undo;

namespace DynamicGeometry;

/// <summary>
/// A figure with a direction the user can type, fix and free - the way from its first point
/// to its second: a segment, a vector. <see cref="LengthPanel"/> shows it under the length
/// right after one is made, with Horizontal and Vertical.
/// </summary>
public interface IFixableDirection : IFigure, IConditionalProperties
{
    /// <summary>In degrees, counterclockwise from the x axis</summary>
    double Angle { get; set; }

    IPoint DirectionEndpoint { get; }

    void FixDirection();

    void FreeDirection();

    void MakeHorizontal();

    void MakeVertical();
}

/// <summary>Undo puts back the end that turned and the Number holding a fixed direction, as <see cref="LengthPropertyValue"/> does for a length</summary>
public class DirectionPropertyValue : ConditionalPropertyValue, IRestorableValue
{
    public virtual IFixableDirection Figure => (IFixableDirection)Parent;

    public object CaptureState()
    {
        var endpoint = Figure.DirectionEndpoint;
        var place = endpoint as IRestorablePlace;
        var number = (endpoint as TranslatedPoint)?.DirectionSource as Number;
        return (place, place?.CapturePlace(), number, number?.Value);
    }

    public void RestoreState(object state)
    {
        var (place, position, number, value) = ((IRestorablePlace, object, Number, double?))state;
        if (number != null)
        {
            number.Value = value.Value;
        }

        place?.RestorePlace(position);
        Figure.Drawing.Recalculate();
    }
}

/// <summary>
/// The rules and the surgery behind "Fix direction" and "Free direction", the counterpart of
/// <see cref="LengthConstraint"/>. An end whose direction from the pivot is fixed is a
/// <see cref="TranslatedPoint"/> from the pivot with a Number for its direction, as a fixed
/// length is one with a Number for its distance - the same point, so the two are fixed and
/// freed one by one through its Free direction and Free distance: fixed in direction alone,
/// the end slides along the line through the pivot; fixed both ways, it is fully determined.
/// </summary>
public static class DirectionConstraint
{
    /// <summary>
    /// Whether the end can be turned about the pivot: a free point the pivot isn't built on
    /// (else the pivot would turn with it), or a translated point from the pivot with a free
    /// direction (a fixed length). A point on a figure or inside a shape would only get near.
    /// </summary>
    public static bool CanTurn(IFigure end, IFigure pivot)
    {
        if (end.Locked || pivot.DependsOn(end) || PointSnapping.IsHeldByLocus(end))
        {
            return false;
        }

        if (end is TranslatedPoint translated)
        {
            return translated.Source == pivot && translated.IsDirectionFree;
        }

        return end is FreePoint && end is not PointOnFigure && end is not PointInRegion;
    }

    /// <summary>The translated point holding a fixed direction from the pivot, or null</summary>
    public static TranslatedPoint FixedEnd(IFigure end, IFigure pivot)
    {
        return end is TranslatedPoint translated && translated.Source == pivot && translated.DirectionSource is Number
            ? translated
            : null;
    }

    /// <summary>Whether the fixed direction lies along the axis at <paramref name="axis"/> degrees, either way</summary>
    public static bool IsAlong(TranslatedPoint fixedEnd, double axis)
    {
        double remainder = Normalize(((Number)fixedEnd.DirectionSource).Value - axis) % 180;
        return remainder < 1e-9 || 180 - remainder < 1e-9;
    }

    /// <summary>The way along the axis at <paramref name="axis"/> degrees nearer to the direction: the axis or its opposite</summary>
    public static double NearestAlong(double direction, double axis)
    {
        return System.Math.Cos((direction - axis).ToRadians()) >= 0 ? axis : Normalize(axis + 180);
    }

    /// <summary>
    /// Points the end from the pivot at the direction, in degrees: by changing the Number when
    /// the direction is fixed, by turning the end once about the pivot at its distance
    /// otherwise. Dragging then changes an unfixed direction again.
    /// </summary>
    public static void SetDirection(IPoint end, IPoint pivot, double degrees)
    {
        var fixedEnd = FixedEnd(end, pivot);
        if (fixedEnd != null)
        {
            // the Number is the way the line through the pivot points, and an end behind the
            // pivot (a negative distance) points the other way
            ((Number)fixedEnd.DirectionSource).Value = Normalize(fixedEnd.Distance < 0 ? degrees + 180 : degrees);
            return;
        }

        if (!CanTurn(end, pivot))
        {
            return;
        }

        if (end is TranslatedPoint translated)
        {
            translated.Direction = Normalize(translated.Distance < 0 ? degrees + 180 : degrees);
            return;
        }

        var center = pivot.Coordinates;
        double distance = center.Distance(end.Coordinates);
        if (distance == 0)
        {
            return;
        }

        ((IMovable)end).MoveTo(Math.GetTranslationPoint(center, distance, degrees.ToRadians()));
        end.RecalculateAndUpdateVisual();
        end.RecalculateAllDependents();
    }

    /// <summary>
    /// The end keeps its direction from the pivot from now on: it becomes a translated point
    /// from the pivot whose direction is a Number at the current one, distance free - so it
    /// slides along the line through the pivot, and the pivot carries it along. An end with a
    /// fixed length gets its direction fixed as well. Nothing moves; one undo step.
    /// </summary>
    public static void Fix(IPoint end, IPoint pivot)
    {
        var drawing = end.Drawing;
        if (end is TranslatedPoint translated)
        {
            drawing.ActionManager.SetProperty(translated, nameof(TranslatedPoint.FreeDirection), false);
            return;
        }

        using (Transaction.Create(drawing.ActionManager, false))
        {
            var direction = Number.CreateAuxiliary(drawing, Math.GetAngle(pivot.Coordinates, end.Coordinates).ToDegrees());
            Actions.Add(drawing, direction);
            var fixedEnd = Factory.CreateTranslatedPoint(drawing, pivot, distanceSource: null, directionSource: direction);
            // the free distance: where the end is now
            fixedEnd.MoveTo(end.Coordinates);
            Actions.ReplacePoint((PointBase)end, fixedEnd);
        }
    }

    /// <summary>Whether <see cref="Free"/> applies: not to a point a locus is drawn from that would have to become another kind of point</summary>
    public static bool CanFree(TranslatedPoint fixedEnd)
    {
        return !fixedEnd.IsDistanceFree || !PointSnapping.IsHeldByLocus(fixedEnd);
    }

    /// <summary>
    /// The opposite of <see cref="Fix"/>: an end fixed in direction alone becomes a free point
    /// where it is (its Number goes with it); one with a fixed length keeps that and turns
    /// about the pivot again.
    /// </summary>
    public static void Free(TranslatedPoint fixedEnd)
    {
        var drawing = fixedEnd.Drawing;
        if (fixedEnd.IsDistanceFree)
        {
            Actions.ReplacePoint(fixedEnd, Factory.CreateFreePoint(drawing, fixedEnd.Coordinates));
        }
        else
        {
            drawing.ActionManager.SetProperty(fixedEnd, nameof(TranslatedPoint.FreeDirection), true);
        }
    }

    /// <summary>In [0, 360)</summary>
    static double Normalize(double degrees)
    {
        double result = degrees % 360;
        return result < 0 ? result + 360 : result;
    }
}
