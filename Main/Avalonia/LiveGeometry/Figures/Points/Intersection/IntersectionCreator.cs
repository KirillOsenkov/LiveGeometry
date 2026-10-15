using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;

namespace DynamicGeometry;

/// <summary>
/// A point where two figures cross, made by clicking the two figures: lines, rays, segments,
/// circles, arcs, and an ellipse with a line (<see cref="IntersectionPoint.GetAlgorithms"/>).
/// The Point tool makes the same point from a click on the crossing itself; this is for
/// crossings that are crowded or hard to hit. Where two figures cross twice, the click on
/// the second one picks the crossing nearer to it.
/// </summary>
[Category(BehaviorCategories.Points)]
[Order(3)]
public class IntersectionCreator : FigureCreator
{
    static readonly DependencyList figureFigure = new DependencyList(typeof(IFigure), typeof(IFigure));

    PointPlacement hoverIntersection;

    protected override DependencyList InitExpectedDependencies()
    {
        return figureFigure;
    }

    /// <summary>
    /// First any figure that can be intersected; then one that crosses it near the cursor
    /// where there is no point yet - the click and the hover preview both ask this
    /// </summary>
    protected override IReadOnlyList<IFigure> FindExpectedDependencies(Point coordinates)
    {
        return Drawing.Figures.HitTestAll(coordinates, figure =>
        {
            if (figure == null || !figure.Visible || !figure.IsHitTestVisible)
            {
                return false;
            }

            if (FoundDependencies.Count == 0)
            {
                return (figure is ILine || figure is IEllipse) && !(figure is AngleArc);
            }

            return figure != FoundDependencies[0] && FindNewIntersection(figure, coordinates) != null;
        });
    }

    /// <summary>
    /// Where the second figure crosses the first nearest to the coordinates, unless a point
    /// is there already (two circles through the same two points): null then, or if they
    /// don't cross
    /// </summary>
    PointPlacement FindNewIntersection(IFigure second, Point coordinates)
    {
        var placement = PointPlacement.Intersection(FoundDependencies[0], second, coordinates);
        if (placement == null)
        {
            return null;
        }

        var there = placement.Coordinates;
        bool occupied = Drawing.Figures
            .OfType<IPoint>()
            .Any(p => p.Visible && p.Exists && p.Coordinates.EqualsWithPrecision(there));
        return occupied ? null : placement;
    }

    protected override IEnumerable<IFigure> CreateFigures()
    {
        var placement = FindNewIntersection(FoundDependencies[1], ClickedUnconstrainedCoordinates);
        if (placement != null)
        {
            yield return placement.Create(Drawing);
        }
    }

    public override void MouseMove(object sender, MouseEventArgs e)
    {
        base.MouseMove(sender, e);
        hoverIntersection = null;
        if (FoundDependencies.Count == 1)
        {
            var coordinates = Coordinates(e, false, false, false);
            var second = LookForExpectedDependencyUnderCursor(coordinates);
            if (second != null)
            {
                hoverIntersection = FindNewIntersection(second, coordinates);
            }
        }
    }

    // the point the click would make, with a halo on both figures
    protected override PointPlacement GetClickPreview(MouseEventArgs e)
    {
        return hoverIntersection ?? base.GetClickPreview(e);
    }

    public override string ConstructionHintText(Drawing.ConstructionStepCompleteEventArgs args)
    {
        if (FoundDependencies.Count == 1)
        {
            return "Click a figure that crosses it, near the crossing you want.";
        }

        return base.ConstructionHintText(args);
    }

    public override string Name
    {
        get { return "Intersection"; }
    }

    public override string HintText
    {
        get { return "Click two figures that cross (lines, segments, rays, circles) to put a point where they cross."; }
    }

    public override FrameworkElement CreateIcon()
    {
        // a diagonal through the center crosses the circle at 1:30 and 7:30,
        // 0.35 * cos 45° = 0.2475 away from the center each way
        return IconBuilder.BuildIcon()
            .Circle(0.5, 0.5, 0.35)
            .Line(0, 1, 1, 0)
            .Point(0.7475, 0.2525, nameof(AppTheme.IntersectionPointFill))
            .Point(0.2525, 0.7475, nameof(AppTheme.IntersectionPointFill))
            .Canvas;
    }
}
