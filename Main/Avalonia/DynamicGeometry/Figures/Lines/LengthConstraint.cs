using System.Collections.Generic;
using System.Linq;
using Avalonia;
using GuiLabs.Undo;

namespace DynamicGeometry;

/// <summary>
/// A figure with a length the user can type, fix and free: a segment, a vector, a regular
/// polygon (its side). <see cref="LengthPanel"/> shows exactly that right after one is made.
/// </summary>
public interface IFixableLength : IFigure, IConditionalProperties
{
    double Length { get; set; }

    void FixLength();

    void FreeLength();

    /// <summary>
    /// What a distance measurement showing this length depends on (the segment itself, a
    /// circle's center and rim point); null when there is no figure to hang one on.
    /// </summary>
    IList<IFigure> MeasuredFigures { get; }
}

/// <summary>
/// The rules and the surgery behind "Fix length" and "Free length", for any end kept at a
/// distance from a pivot: a segment's second end from its first, a regular polygon's vertex
/// from its center.
/// </summary>
public static class LengthConstraint
{
    /// <summary>A new distance measurement of the figure's length, not yet in the drawing; null when it has none</summary>
    public static DistanceMeasurement CreateMeasurement(IFixableLength figure)
    {
        var measured = figure.MeasuredFigures;
        return measured == null ? null : Factory.CreateDistanceMeasurement(figure.Drawing, measured);
    }

    /// <summary>
    /// The distance measurement already showing the figure's length, or null: one on the
    /// same figures, or - for a segment - on its two points, which is what the Distance
    /// tool makes.
    /// </summary>
    public static DistanceMeasurement FindMeasurement(IFixableLength figure)
    {
        var measured = figure.MeasuredFigures;
        if (measured == null)
        {
            return null;
        }

        var ends = measured.Count == 1 && measured[0] is Segment segment ? segment.Dependencies : null;
        return figure.Drawing.Figures
            .OfType<DistanceMeasurement>()
            .FirstOrDefault(m => m.Dependencies.SequenceEqual(measured) || (ends != null && m.Dependencies.SequenceEqual(ends)));
    }

    /// <summary>
    /// Whether the end can be moved to a new distance from the pivot: a free point the pivot
    /// isn't built on (else the pivot would follow and the distance stay), a point on a line
    /// that runs through the pivot (<see cref="LineThrough"/>), or a translated point sliding
    /// along the line from the pivot (free distance, fixed direction). A point on some other
    /// figure would only get near.
    /// </summary>
    public static bool CanStretch(IFigure end, IFigure pivot)
    {
        // (nor a point a locus is drawn from: fixing the length would put another kind of
        // point in its place, and the locus would trace nothing - PointSnapping.IsHeldByLocus)
        if (end.Locked || pivot.DependsOn(end) || PointSnapping.IsHeldByLocus(end))
        {
            return false;
        }

        if (end is PointOnFigure onFigure)
        {
            return LineThrough(onFigure, pivot) != null;
        }

        if (end is FreePoint)
        {
            return true;
        }

        return end is TranslatedPoint translated
            && translated.IsDistanceFree
            && !translated.IsDirectionFree
            && translated.Source == pivot;
    }

    /// <summary>
    /// The line the end slides on when the pivot is on that line too: moving the end away from
    /// the pivot keeps it on the line, and fixing the distance leaves it no freedom at all.
    /// Null for a point on any other figure.
    /// </summary>
    static ILine LineThrough(PointOnFigure end, IFigure pivot)
    {
        return end.LinearFigure is ILine line && pivot is IPoint point && IsOnLine(point, line) ? line : null;
    }

    /// <summary>
    /// Whether the point is on the line by construction, not by chance: a point on it, an
    /// intersection with it, or a point the line is built on and runs through (the points of a
    /// line, segment or ray, the point of a parallel or a perpendicular, the vertex of an angle
    /// bisector - not the two points of a segment bisector).
    /// </summary>
    static bool IsOnLine(IPoint point, ILine line)
    {
        if (point is PointOnFigure onFigure && onFigure.LinearFigure == line)
        {
            return true;
        }

        if (point is IntersectionPoint && point.Dependencies.Contains(line))
        {
            return true;
        }

        if (!line.Dependencies.Contains(point))
        {
            return false;
        }

        var coordinates = point.Coordinates;
        return Math.GetDistanceToLine(coordinates, line.Coordinates) <= Tolerance(coordinates);
    }

    static double Tolerance(Point coordinates)
    {
        return Math.TangencyTolerance(System.Math.Max(System.Math.Abs(coordinates.X), System.Math.Abs(coordinates.Y)));
    }

    /// <summary>The translated point holding a fixed distance from the pivot, or null</summary>
    public static TranslatedPoint FixedEnd(IFigure end, IFigure pivot)
    {
        return end is TranslatedPoint translated && translated.Source == pivot && translated.DistanceSource is Number
            ? translated
            : null;
    }

    /// <summary>
    /// Puts the end at the distance from the pivot: by changing the Number when the distance
    /// is fixed, by moving the end once along the current direction otherwise. Dragging then
    /// changes an unfixed distance again.
    /// </summary>
    public static void SetDistance(IPoint end, IPoint pivot, double distance)
    {
        var fixedEnd = FixedEnd(end, pivot);
        if (fixedEnd != null)
        {
            // the Number is signed (an end behind the pivot along a fixed direction); the length isn't
            fixedEnd.Distance = fixedEnd.Distance < 0 ? -distance : distance;
            return;
        }

        var current = pivot.Coordinates.Distance(end.Coordinates);
        if (!CanStretch(end, pivot) || current == 0 || distance == current)
        {
            return;
        }

        var target = Math.GetDilationPoint(end.Coordinates, pivot.Coordinates, distance / current);
        ((IMovable)end).MoveTo(target);
        end.RecalculateAndUpdateVisual();
        var dependents = DependencyAlgorithms.FindDescendants(f => f.Dependents, end.AsEnumerable<IFigure>());
        dependents.Reverse();
        foreach (var dependent in dependents)
        {
            dependent.RecalculateAndUpdateVisual();
        }
    }

    /// <summary>
    /// The end keeps its distance from the pivot from now on: it becomes a translated point
    /// from the pivot at the current distance held by a Number, direction free - so it drags
    /// around the pivot and the pivot carries it along. An end on a line through the pivot
    /// takes the line's direction instead (the distance signed along it) and is fully
    /// determined. An end already sliding along the line just has its distance fixed. Nothing
    /// moves; one undo step.
    /// </summary>
    public static void Fix(IPoint end, IPoint pivot)
    {
        var drawing = end.Drawing;
        if (end is TranslatedPoint sliding)
        {
            drawing.ActionManager.SetProperty(sliding, "FreeDistance", false);
            return;
        }

        var line = end is PointOnFigure onFigure ? LineThrough(onFigure, pivot) : null;
        using (Transaction.Create(drawing.ActionManager, false))
        {
            var distance = Number.CreateAuxiliary(drawing, line == null
                ? pivot.Coordinates.Distance(end.Coordinates)
                : DistanceAlong(line, pivot.Coordinates, end.Coordinates));
            Actions.Add(drawing, distance);
            var fixedEnd = Factory.CreateTranslatedPoint(drawing, pivot, distance, directionSource: line);
            // the free direction, if any: where the end is now
            fixedEnd.MoveTo(end.Coordinates);
            Actions.ReplacePoint((PointBase)end, fixedEnd);
        }
    }

    /// <summary>How far the point is from the origin the way the line points: negative behind it</summary>
    static double DistanceAlong(ILine line, Point origin, Point point)
    {
        var direction = Math.GetAngle(line.Coordinates.P1, line.Coordinates.P2);
        return (point.X - origin.X) * System.Math.Cos(direction) + (point.Y - origin.Y) * System.Math.Sin(direction);
    }

    /// <summary>
    /// The opposite of <see cref="Fix"/>: the fixed end becomes a free point where it is (its
    /// Number goes with it). When its direction is fixed too, it keeps sliding along that
    /// line: as a point on the line when the pivot is on it as well (what Fix made of such a
    /// point) and the end is within the line's extent, else as a translated point with the
    /// distance free.
    /// </summary>
    public static void Free(TranslatedPoint fixedEnd)
    {
        var drawing = fixedEnd.Drawing;
        if (PointSnapping.IsHeldByLocus(fixedEnd))
        {
            // a point a locus is drawn from stays the point it is: only its distance is let go
            drawing.ActionManager.SetProperty(fixedEnd, "FreeDistance", true);
        }
        else if (fixedEnd.IsDirectionFree)
        {
            Actions.ReplacePoint(fixedEnd, Factory.CreateFreePoint(drawing, fixedEnd.Coordinates));
        }
        else if (fixedEnd.DirectionSource is ILine line && IsOnLine(fixedEnd.Source, line) && IsWithin(line, fixedEnd.Coordinates))
        {
            Actions.ReplacePoint(fixedEnd, Factory.CreatePointOnFigure(drawing, line, fixedEnd.Coordinates));
        }
        else
        {
            drawing.ActionManager.SetProperty(fixedEnd, "FreeDistance", true);
        }
    }

    /// <summary>Whether a point on the line through a segment or a ray is on the segment or ray itself</summary>
    static bool IsWithin(ILine line, Point point)
    {
        var nearest = line.GetPointFromParameter(line.GetNearestParameterFromPoint(point));
        return nearest.Distance(point) <= Tolerance(point);
    }
}
