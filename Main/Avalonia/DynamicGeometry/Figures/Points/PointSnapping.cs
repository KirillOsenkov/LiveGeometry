using System.Collections.Generic;
using System.Linq;

namespace DynamicGeometry;

/// <summary>
/// Turns a point into another kind of point where it is, keeping its name, label and what is
/// built on it (<see cref="Actions.ReplacePoint"/>, one undo step): a free point snapped onto a
/// figure, and a point tied to figures - on a figure, at an intersection, a midpoint - released
/// into a free point. The property grid, the context menu and the Alt-drag of the
/// <see cref="Dragger"/> all come here.
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
    /// the default names (LineTwoPoints1) mean nothing to a reader. A name the user gave, or
    /// a figure not built on its points alone, is called by its name.
    /// </summary>
    public static string Describe(IFigure figure)
    {
        var dependencies = figure.Dependencies;
        bool defaultName = figure.Name == null || figure.Name.StartsWith(figure.GetType().Name);
        if (!defaultName || dependencies.Count != 2 || !(dependencies[0] is IPoint first))
        {
            return figure.Name;
        }

        if (figure is Circle)
        {
            return "circle " + first.Name;
        }

        if (dependencies[1] is IPoint second)
        {
            var kind = figure is Segment ? "segment" : figure is Ray ? "ray" : figure is LineTwoPoints ? "line" : null;
            if (kind != null)
            {
                return kind + " " + first.Name + second.Name;
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

    /// <summary>The point becomes what the placement says (on a figure, an intersection, a midpoint)</summary>
    public static PointBase Snap(PointBase point, PointPlacement placement)
    {
        var snapped = (PointBase)placement.Create(point.Drawing);
        Replace(point, snapped);
        return snapped;
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
