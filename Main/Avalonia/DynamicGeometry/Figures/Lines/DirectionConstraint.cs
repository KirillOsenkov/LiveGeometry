using GuiLabs.Undo;

namespace DynamicGeometry;

/// <summary>
/// A figure with a direction the user can type, fix and free - the way from its first point
/// to its second: a segment, a vector, a ray, a line through two points.
/// <see cref="DimensionsPanel"/> shows it right after one is made, with Horizontal and
/// Vertical. (A parallel, a perpendicular and a bisector are lines too, built on a line or
/// an angle: they say no to all of it.)
/// </summary>
public interface IFixableDirection : IFigure, IConditionalProperties
{
    /// <summary>In degrees, counterclockwise from the x axis</summary>
    double Direction { get; set; }

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
        return IsAlong(((Number)fixedEnd.DirectionSource).Value, axis);
    }

    /// <summary>Whether the direction, in degrees, lies along the axis at <paramref name="axis"/> degrees, either way</summary>
    public static bool IsAlong(double degrees, double axis)
    {
        double remainder = Normalize(degrees - axis) % 180;
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

    #region A figure through two points

    // What a segment, a ray and a line through two points do with their direction, from
    // their first point to their second: each has the row and the verbs, and hands them here.

    /// <summary>The direction of the line from its first point to its second, in degrees from 0 to 360</summary>
    public static double Of(PointPair line)
    {
        // a full turn but for rounding is no turn: a first end fixed at 180° from the second,
        // a hair below it (sin 180° is no 0 in doubles), said 360
        double angle = Math.GetAngle(line.P1, line.P2).ToDegrees();
        return 360 - angle < 1e-9 ? 0 : angle;
    }

    /// <summary>Whether the name is the direction row's or one of its verbs': what a figure's CanEdit hands to <see cref="CanEdit"/></summary>
    public static bool IsMember(string name)
    {
        return name == nameof(IFixableDirection.Direction)
            || name == nameof(IFixableDirection.FixDirection)
            || name == nameof(IFixableDirection.FreeDirection)
            || name == nameof(IFixableDirection.MakeHorizontal)
            || name == nameof(IFixableDirection.MakeVertical);
    }

    /// <summary>
    /// The row while an end can be turned or holds a fixed direction; Fix direction while none
    /// holds one and an end can be turned; Free direction while one does; Horizontal and
    /// Vertical unless the direction is fixed along that axis already (they would change
    /// nothing). None of it on a figure not built on two points.
    /// </summary>
    public static bool CanEdit(IFigure figure, string name)
    {
        switch (name)
        {
            case nameof(IFixableDirection.Direction):
                return EndIndex(figure) >= 0;
            case nameof(IFixableDirection.FixDirection):
                return FixedEndIndex(figure) < 0 && EndToTurn(figure) >= 0;
            case nameof(IFixableDirection.FreeDirection):
                return CanFreeDirection(figure);
            case nameof(IFixableDirection.MakeHorizontal):
                return CanAlign(figure, axis: 0);
            case nameof(IFixableDirection.MakeVertical):
                return CanAlign(figure, axis: 90);
            default:
                return true;
        }
    }

    /// <summary>The end that answers for the direction (<see cref="EndIndex"/>), or null</summary>
    public static IPoint Endpoint(IFigure figure)
    {
        int end = EndIndex(figure);
        return end >= 0 ? End(figure, end) : null;
    }

    /// <summary>
    /// Points the figure the way given, in degrees: the end that answers for the direction
    /// is turned about the other, or its Number changed (<see cref="SetDirection(IPoint, IPoint, double)"/>);
    /// the first end is turned the opposite way. Nothing when no end can be.
    /// </summary>
    public static void SetDirection(IFigure figure, double degrees)
    {
        int end = EndIndex(figure);
        if (end >= 0)
        {
            SetDirection(End(figure, end), End(figure, 1 - end), end == 1 ? degrees : degrees + 180);
        }
    }

    /// <summary>The figure keeps its direction from now on (<see cref="Fix"/>)</summary>
    public static void FixDirection(IFigure figure)
    {
        int end = EndToTurn(figure);
        if (FixedEndIndex(figure) < 0 && end >= 0)
        {
            Fix(End(figure, end), End(figure, 1 - end));
        }
    }

    /// <summary>The opposite of <see cref="FixDirection"/> (<see cref="Free"/>)</summary>
    public static void FreeDirection(IFigure figure)
    {
        if (CanFreeDirection(figure))
        {
            Free((TranslatedPoint)figure.Dependencies[FixedEndIndex(figure)]);
        }
    }

    /// <summary>
    /// The direction fixed along the axis at <paramref name="axis"/> degrees, the nearer way
    /// (rightward or leftward, up or down: the second point stays on its side of the first).
    /// One undo step: the fix, then the direction set through the figure's row, so that its
    /// undo puts the Number back.
    /// </summary>
    public static void Align(IFixableDirection figure, double axis)
    {
        if (!CanAlign(figure, axis))
        {
            return;
        }

        double direction = NearestAlong(figure.Direction, axis);
        var actionManager = figure.Drawing.ActionManager;
        using (Transaction.Create(actionManager, false))
        {
            FixDirection(figure);
            Actions.SetProperty(
                actionManager,
                PropertyDiscoveryStrategy.CreateValueProvider(figure, nameof(IFixableDirection.Direction)),
                direction);
        }
    }

    /// <summary>The end holding a fixed direction from the other (see <see cref="FixedEnd"/>); -1 when none does</summary>
    static int FixedEndIndex(IFigure figure)
    {
        return IndexOfEnd(figure, (end, pivot) => FixedEnd(end, pivot) != null);
    }

    /// <summary>The end a new direction turns (see <see cref="CanTurn"/>), the second preferred; -1 when neither can</summary>
    static int EndToTurn(IFigure figure)
    {
        return IndexOfEnd(figure, CanTurn);
    }

    /// <summary>The end that answers for the direction: the fixed one, else the one that can be turned</summary>
    static int EndIndex(IFigure figure)
    {
        int fixedEnd = FixedEndIndex(figure);
        return fixedEnd >= 0 ? fixedEnd : EndToTurn(figure);
    }

    static bool CanFreeDirection(IFigure figure)
    {
        int end = FixedEndIndex(figure);
        return end >= 0 && CanFree((TranslatedPoint)figure.Dependencies[end]);
    }

    static bool CanAlign(IFigure figure, double axis)
    {
        if (EndIndex(figure) < 0)
        {
            return false;
        }

        int fixedEnd = FixedEndIndex(figure);
        return fixedEnd < 0 || !IsAlong((TranslatedPoint)figure.Dependencies[fixedEnd], axis);
    }

    static IPoint End(IFigure figure, int index)
    {
        return (IPoint)figure.Dependencies[index];
    }

    /// <summary>The end, second first, that qualifies with the other for its pivot; -1 when neither, or when the figure isn't built on two points</summary>
    static int IndexOfEnd(IFigure figure, System.Func<IFigure, IFigure, bool> qualifies)
    {
        var ends = figure.Dependencies;
        if (ends == null || ends.Count != 2 || ends[0] is not IPoint || ends[1] is not IPoint)
        {
            return -1;
        }

        for (int end = 1; end >= 0; end--)
        {
            if (qualifies(ends[end], ends[1 - end]))
            {
                return end;
            }
        }

        return -1;
    }

    #endregion
}
