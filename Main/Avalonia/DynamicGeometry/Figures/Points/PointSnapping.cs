using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using GuiLabs.Undo;

namespace DynamicGeometry;

/// <summary>
/// Turns a point into another kind of point where it is, keeping its name, label and what is
/// built on it (<see cref="Actions.ReplacePoint"/>, one undo step): a free point snapped onto a
/// figure, and a point tied to figures - on a figure, at an intersection, a midpoint - released
/// into a free point. A free point also becomes a point by coordinates and back
/// (<see cref="ConvertToPointByCoordinates"/>). A point dragged onto another point joins it
/// (<see cref="Join"/>). The property grid, the context menu and the Alt-drag of the
/// <see cref="Dragger"/> all come here.
/// </summary>
public static class PointSnapping
{
    /// <summary>The kinds of point a click puts on figures (<see cref="PointPlacement"/>), which can be let go</summary>
    public static bool CanRelease(IFigure point)
    {
        return (point is PointOnFigure || point is IntersectionPoint || point is MidPoint)
            && !IsHeldByLocus(point);
    }

    /// <summary>
    /// A locus is drawn from two points: the one that slides along its figure, and the one
    /// whose trace it is, built on the first. Neither can become another kind of point or
    /// join one: the locus would be the trace of nothing. (Let go with Alt, the sliding
    /// point left a locus that threw on every move, a segment on a point that was not in
    /// the drawing, and an undo that did not bring the drawing back.)
    /// </summary>
    public static bool IsHeldByLocus(IFigure point)
    {
        return point.Dependents.OfType<Locus>().Any();
    }

    /// <summary>
    /// What "Free point" is offered for: the points that can be let go, and a point by
    /// coordinates that is somewhere. Dragging with Alt doesn't free that one
    /// (<see cref="CanRelease"/>): its place was typed on purpose.
    /// </summary>
    public static bool CanFree(IFigure point)
    {
        return CanRelease(point)
            || point is PointByCoordinates byCoordinates
                && byCoordinates.Exists
                && byCoordinates.Coordinates.Exists()
                && !IsHeldByLocus(point);
    }

    /// <summary>The point becomes a free point where it is</summary>
    public static FreePoint Release(PointBase point)
    {
        var free = Factory.CreateFreePoint(point.Drawing, point.Coordinates);
        Replace(point, free);
        return free;
    }

    /// <summary>A free point proper: a point on a figure is a <see cref="FreePoint"/> only by inheritance</summary>
    public static bool CanConvertToPointByCoordinates(IFigure point)
    {
        return point is FreePoint && !(point is PointOnFigure);
    }

    /// <summary>
    /// The free point becomes a point by coordinates: it stays where its X and Y say, which
    /// start as the numbers it is at, as the grid shows them, to be typed over (the grid
    /// puts the keyboard into X). The way back is <see cref="Release"/>.
    /// </summary>
    public static PointByCoordinates ConvertToPointByCoordinates(FreePoint point)
    {
        var drawing = point.Drawing;
        var coordinates = point.Coordinates;
        var converted = Factory.CreatePointByCoordinates(drawing, ConstantText(coordinates.X), ConstantText(coordinates.Y));
        converted.Recalculate();
        Replace(point, converted);
        if (converted.Selected)
        {
            drawing.RaiseDisplayProperties(converted, focusProperty: converted.XExpression.Name);
        }

        return converted;
    }

    /// <summary>
    /// A number as an expression says it: rounded as the grid shows numbers
    /// (<see cref="Settings.DisplayDecimals"/>), and never with an exponent, which the
    /// expression language doesn't read
    /// </summary>
    static string ConstantText(double value)
    {
        // + 0.0: a small negative number rounds to -0, which would read "-0"
        var rounded = System.Math.Round(value, Settings.DisplayDecimals) + 0.0;
        return rounded.ToString("0." + new string('#', Settings.DisplayDecimals), CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// The figures a point can be snapped onto from where it is: those passing through it
    /// (within the cursor's reach) that can hold a point, except the one it is on already and
    /// those built on it, which would make a loop. None for a locked point: snapping moves it.
    /// None for a point a locus is drawn from (<see cref="IsHeldByLocus"/>).
    /// </summary>
    public static IList<IFigure> FiguresToSnapTo(PointBase point)
    {
        var drawing = point.Drawing;
        if (drawing == null || point.Locked || IsHeldByLocus(point))
        {
            return new IFigure[0];
        }

        var current = (point as PointOnFigure)?.LinearFigure;
        return drawing.Figures.HitTestMany(point.Coordinates)
            .Where(f => f.IsHitTestVisible)
            .Select(BezierPath.PointHolder)
            .Distinct()
            .Where(f => f != current
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
    /// Whether the point can be joined into the target: not one built on it (a loop). Nor
    /// when an expression names both (a label [dist(B, C)]): what is built on both points
    /// otherwise collapses and goes (<see cref="Collapse"/>), but an expression of the two
    /// still has a meaning. Nor into a point without a name (a vertex a regular polygon
    /// works out) when expressions name the point: they would have nothing to call the
    /// target. Nor a point a locus is drawn from (<see cref="IsHeldByLocus"/>).
    /// </summary>
    public static bool CanJoin(PointBase point, IPoint target)
    {
        return target != point
            && !IsHeldByLocus(point)
            && !target.DependsOn(point)
            && !FindBuiltOnBoth(point, target).Any(dependent => dependent is IRenamableExpressions)
            && !(string.IsNullOrEmpty(target.Name) && point.Dependents.OfType<IRenamableExpressions>().Any())
            && CanJoinOnBezierPaths(point, target);
    }

    /// <summary>
    /// Nothing joins a handle of a Bezier path, nor is joined into one; and two anchors of a
    /// path join only when they are next to each other and the path keeps two (the point
    /// leaves the path, <see cref="BezierPath.CanDropAnchorInto"/>) - not to take the path
    /// away, as the collapse of anything else built on both would
    /// </summary>
    static bool CanJoinOnBezierPaths(PointBase point, IPoint target)
    {
        if (point is BezierPath.BezierPathHandle || target is BezierPath.BezierPathHandle)
        {
            return false;
        }

        return point.Dependents
            .OfType<BezierPath>()
            .Where(path => path.IsAnchor(point) && path.IsAnchor(target))
            .All(path => path.CanDropAnchorInto(point, target));
    }

    /// <summary>The figures built on the point that use the target too (segment BC, for C and B)</summary>
    static IFigure[] FindBuiltOnBoth(PointBase point, IPoint target)
    {
        return point.Dependents
            .Where(dependent => !(dependent is PointLabel) && dependent.Dependencies.Contains(target))
            .ToArray();
    }

    /// <summary>
    /// The point joins the target: what was built on it is built on the target from now on,
    /// and the point goes, its name and label with it. What is built on both collapses
    /// (<see cref="Collapse"/>). One undo step. There is no way back but undo: taking the
    /// point out again would have to guess which of the target's dependents were its own.
    /// </summary>
    public static void Join(PointBase point, IPoint target)
    {
        var drawing = point.Drawing;
        bool selected = point.Selected;
        using (Transaction.Create(drawing.ActionManager, false))
        {
            Collapse(point, target);
            var dependents = point.Dependents.Where(d => !(d is PointLabel)).ToArray();

            // Expressions that name the point ([A.X] in a label, the X of a point by
            // coordinates) name the target from now on. Compiled, they hold the point
            // itself: left alone they kept showing where it was last, and the saved
            // file named a point that is gone. The text is rewritten once the point has
            // left (its name then means the target alone), and put back as it was on
            // undo, where it is compiled again last of all, when the point is back.
            var expressions = dependents.OfType<IRenamableExpressions>().ToArray();
            void Rebind()
            {
                foreach (var holder in expressions)
                {
                    holder.RebindExpressions();
                }
            }

            if (expressions.Length > 0)
            {
                drawing.ActionManager.RecordAction(new CallMethodAction(() => { }, Rebind));
            }

            // The target, with what it is built on, goes before the first figure that is
            // about to be built on it: the list is in dependency order, and a file is
            // written in the order of the list (joined into a point made later, what was
            // built on the point came before the target, and a saved file came back in
            // another order).
            var figures = drawing.Figures;
            var first = dependents
                .Select(dependent => figures.FindTopLevel(dependent))
                .Where(listed => listed != null)
                .OrderBy(listed => figures.IndexOf(listed))
                .FirstOrDefault();
            if (first != null)
            {
                Actions.MoveBefore(drawing, target, first);
            }

            foreach (var dependent in dependents)
            {
                // a part of a composite (a side of a regular polygon built on the point)
                // has gone over with its composite already
                if (dependent.Dependencies.Contains(point))
                {
                    Actions.ReplaceDependency(dependent, point, target);
                }
            }

            Actions.Remove(point);

            if (expressions.Length > 0)
            {
                string pointName = point.Name;
                var texts = expressions.Select(holder => holder.ExpressionTexts).ToArray();
                drawing.ActionManager.RecordAction(new CallMethodAction(
                    () =>
                    {
                        var renamer = new ExpressionRenamer(drawing, new Dictionary<IFigure, string>() { { target, pointName } });
                        foreach (var holder in expressions)
                        {
                            holder.RenameInExpressions(renamer);
                        }

                        Rebind();
                    },
                    () =>
                    {
                        for (int i = 0; i < expressions.Length; i++)
                        {
                            expressions[i].ExpressionTexts = texts[i];
                        }
                    }));
            }
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
    /// What is built on both the point and the target would be built on the target twice
    /// once they are joined: segment BC when C joins B (ABCD becomes ABD), a midpoint, a
    /// circle around one through the other. A polygon or polyline in which the two are
    /// neighbors loses the point; anything else goes, with what is built on it, as Delete
    /// would take it. Rewired instead, a segment BB was left, and undo of the rewiring
    /// swapped its ends.
    /// </summary>
    static void Collapse(PointBase point, IPoint target)
    {
        var figures = point.Drawing.Figures;
        foreach (var collapsing in FindBuiltOnBoth(point, target))
        {
            // gone already with one before it (a midpoint of segment BC built on the segment)
            if (!point.Dependents.Contains(collapsing))
            {
                continue;
            }

            if (collapsing is BezierPath path)
            {
                // two anchors next to each other: the path goes on without the point, the
                // target taking its handle on the far side. Otherwise one of the two is a
                // handle, which is the target from now on, as anything built on the point.
                if (path.CanDropAnchorInto(point, target))
                {
                    point.Drawing.ActionManager.RecordAction(path.CreateDropAnchorAction(point, target));
                }
            }
            else if (CanDropVertex(collapsing, point, target))
            {
                Actions.RemoveDependency(collapsing, point);
            }
            else
            {
                // (a part of a composite goes with its composite)
                Actions.Remove(figures.FindTopLevel(collapsing) ?? collapsing);
            }
        }
    }

    /// <summary>
    /// Whether the point is a vertex of a polygon or polyline next to the target, each there
    /// once, and the figure keeps enough vertices without it
    /// </summary>
    static bool CanDropVertex(IFigure figure, PointBase point, IPoint target)
    {
        bool isPolygon = figure.GetType() == typeof(Polygon);
        if (!isPolygon && !(figure is Polyline))
        {
            return false;
        }

        var vertices = figure.Dependencies;
        if (vertices.Count(vertex => vertex == point) != 1
            || vertices.Count(vertex => vertex == target) != 1
            || vertices.Count - 1 < (isPolygon ? 3 : 2))
        {
            return false;
        }

        int apart = System.Math.Abs(vertices.IndexOf(point) - vertices.IndexOf(target));
        return apart == 1 || isPolygon && apart == vertices.Count - 1;
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
