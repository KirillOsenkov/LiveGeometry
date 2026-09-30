using System.Collections.Generic;
using System.Linq;
using GuiLabs.Undo;

namespace DynamicGeometry;

/// <summary>
/// Turns a point into another kind of point where it is, keeping its name, label and what is
/// built on it (<see cref="Actions.ReplacePoint"/>, one undo step): a free point snapped onto a
/// figure, and a point tied to figures - on a figure, at an intersection, a midpoint - released
/// into a free point. A point dragged onto another point joins it (<see cref="Join"/>). The
/// property grid, the context menu and the Alt-drag of the <see cref="Dragger"/> all come here.
/// </summary>
public static class PointSnapping
{
    /// <summary>The kinds of point a click puts on figures (<see cref="PointPlacement"/>), which can be let go</summary>
    public static bool CanRelease(IFigure point)
    {
        return point is PointOnFigure || point is IntersectionPoint || point is MidPoint;
    }

    /// <summary>The point becomes a free point where it is</summary>
    public static FreePoint Release(PointBase point)
    {
        var free = Factory.CreateFreePoint(point.Drawing, point.Coordinates);
        Replace(point, free);
        return free;
    }

    /// <summary>
    /// The figures a point can be snapped onto from where it is: those passing through it
    /// (within the cursor's reach) that can hold a point, except the one it is on already and
    /// those built on it, which would make a loop. None for a locked point: snapping moves it.
    /// </summary>
    public static IList<IFigure> FiguresToSnapTo(PointBase point)
    {
        var drawing = point.Drawing;
        if (drawing == null || point.Locked)
        {
            return new IFigure[0];
        }

        var current = (point as PointOnFigure)?.LinearFigure;
        return drawing.Figures.HitTestMany(point.Coordinates)
            .Where(f => f.IsHitTestVisible
                && f != current
                && PointOnFigure.CanBeOnFigure(f)
                && !f.DependsOn(point))
            .Reverse()
            .ToArray();
    }

    /// <summary>
    /// How a menu names a figure to snap to: "line AB", "segment AB", "ray AB", "circle C" -
    /// a default name alone (AB, Circle1) doesn't say what kind of figure it is. A name the
    /// user gave, or a figure not built on its points alone, is called by its name.
    /// </summary>
    public static string Describe(IFigure figure)
    {
        var dependencies = figure.Dependencies;
        if (!figure.HasDefaultName || dependencies.Count != 2 || !(dependencies[0] is IPoint first))
        {
            return figure.Name;
        }

        if (figure is Circle)
        {
            return "circle " + first.Name;
        }

        if (dependencies[1] is IPoint second)
        {
            var kind = figure is Segment ? "segment"
                : figure is Ray ? "ray"
                : figure is LineTwoPoints ? "line"
                : figure is Vector ? "vector"
                : null;
            if (kind != null)
            {
                // the default name: the points in order (line AB, not BA)
                return kind + " " + figure.Name;
            }
        }

        return figure.Name;
    }

    /// <summary>The point goes onto the figure, at its nearest place, and slides along it from then on</summary>
    public static PointOnFigure SnapTo(PointBase point, IFigure figure)
    {
        var onFigure = Factory.CreatePointOnFigure(point.Drawing, figure, point.Coordinates);
        Replace(point, onFigure);
        return onFigure;
    }

    /// <summary>
    /// The point becomes what the placement says (on a figure, an intersection, a midpoint),
    /// or joins the existing point it names
    /// </summary>
    public static IPoint Snap(PointBase point, PointPlacement placement)
    {
        if (placement.ExistingPoint != null)
        {
            Join(point, placement.ExistingPoint);
            return placement.ExistingPoint;
        }

        var snapped = (PointBase)placement.Create(point.Drawing);
        Replace(point, snapped);
        return snapped;
    }

    /// <summary>
    /// Whether the point can be joined into the target: not one built on it (a loop), and no
    /// figure uses both, which would then use the target twice (a segment between them).
    /// </summary>
    public static bool CanJoin(PointBase point, IPoint target)
    {
        return target != point
            && !target.DependsOn(point)
            && !point.Dependents.Any(d => !(d is PointLabel) && d.Dependencies.Contains(target));
    }

    /// <summary>
    /// The point joins the target: what was built on it is built on the target from now on,
    /// and the point goes, its name and label with it. One undo step. There is no way back
    /// but undo: taking the point out again would have to guess which of the target's
    /// dependents were its own.
    /// </summary>
    public static void Join(PointBase point, IPoint target)
    {
        var drawing = point.Drawing;
        bool selected = point.Selected;
        using (Transaction.Create(drawing.ActionManager, false))
        {
            foreach (var dependent in point.Dependents.Where(d => !(d is PointLabel)).ToArray())
            {
                Actions.ReplaceDependency(dependent, point, target);
            }

            Actions.Remove(point);
        }

        drawing.Recalculate();
        if (selected)
        {
            point.Selected = false;
            target.Selected = true;
            drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());
        }
    }

    /// <summary>
    /// <see cref="Actions.ReplacePoint"/>, and a selected point hands the selection over, so
    /// the property grid shows the new one
    /// </summary>
    public static void Replace(PointBase point, PointBase replacement)
    {
        var drawing = point.Drawing;
        bool selected = point.Selected;
        Actions.ReplacePoint(point, replacement);
        if (selected)
        {
            point.Selected = false;
            replacement.Selected = true;
            drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());
        }
    }
}
