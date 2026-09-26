using System.Collections.Generic;
using System.Linq;
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
    /// isn't built on (else the pivot would follow and the distance stay), or a translated
    /// point sliding along the line from the pivot (free distance, fixed direction). A point
    /// on some other figure would only get near.
    /// </summary>
    public static bool CanStretch(IFigure end, IFigure pivot)
    {
        if (end.Locked || pivot.DependsOn(end))
        {
            return false;
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
            fixedEnd.Distance = distance;
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
    /// around the pivot and the pivot carries it along. An end already sliding along the line
    /// just has its distance fixed. Nothing moves; one undo step.
    /// </summary>
    public static void Fix(IPoint end, IPoint pivot)
    {
        var drawing = end.Drawing;
        if (end is TranslatedPoint sliding)
        {
            drawing.ActionManager.SetProperty(sliding, "FreeDistance", false);
            return;
        }

        using (Transaction.Create(drawing.ActionManager, false))
        {
            var distance = Number.CreateAuxiliary(drawing, pivot.Coordinates.Distance(end.Coordinates));
            Actions.Add(drawing, distance);
            var fixedEnd = Factory.CreateTranslatedPoint(drawing, pivot, distance, directionSource: null);
            // the free direction: where the end is now
            fixedEnd.MoveTo(end.Coordinates);
            Actions.ReplacePoint((PointBase)end, fixedEnd);
        }
    }

    /// <summary>
    /// The opposite of <see cref="Fix"/>: the fixed end becomes a free point where it is (its
    /// Number goes with it), or, when its direction is fixed too, keeps sliding along the line
    /// with the distance free.
    /// </summary>
    public static void Free(TranslatedPoint fixedEnd)
    {
        var drawing = fixedEnd.Drawing;
        if (fixedEnd.IsDirectionFree)
        {
            Actions.ReplacePoint(fixedEnd, Factory.CreateFreePoint(drawing, fixedEnd.Coordinates));
        }
        else
        {
            drawing.ActionManager.SetProperty(fixedEnd, "FreeDistance", true);
        }
    }
}
