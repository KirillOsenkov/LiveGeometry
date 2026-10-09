using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;

namespace DynamicGeometry;

/// <summary>
/// One click on a shape with a perimeter (<see cref="IPerimeter"/>): a polygon, a circle, an
/// ellipse, a sector, a circular segment, a closed Bezier path. The Area tool's shape, without
/// its list of points.
/// </summary>
[Category(BehaviorCategories.Measure)]
[Order(4)]
public class PerimeterMeasurementCreator : FigureCreator
{
    protected override DependencyList InitExpectedDependencies()
    {
        return DependencyList.Create<IPerimeter>();
    }

    /// <summary>Only a shape that has a perimeter right now: an open Bezier path has none</summary>
    protected override IReadOnlyList<IFigure> FindExpectedDependencies(Point coordinates)
    {
        return base.FindExpectedDependencies(coordinates)
            .Where(figure => ((IPerimeter)figure).Perimeter.IsValidValue())
            .ToList();
    }

    protected override IEnumerable<IFigure> CreateFigures()
    {
        yield return Factory.CreatePerimeterMeasurement(Drawing, FoundDependencies);
    }

    public override string Name
    {
        get { return "Perimeter"; }
    }

    public override string HintText
    {
        get { return "Click a polygon, circle, ellipse, sector or closed path to measure its perimeter."; }
    }

    public override FrameworkElement CreateIcon()
    {
        // The Area tool's pentagon, hollow, with a second outline in the accent color a
        // little way in along its walls: the outline is what is measured.
        var pentagon = new[]
        {
            new Point(0.5, 0.08),
            new Point(0.93, 0.39),
            new Point(0.77, 0.9),
            new Point(0.23, 0.9),
            new Point(0.07, 0.39)
        };

        // the same pentagon shrunk about its center: for a regular one that is an inset by
        // the same distance along every wall
        var center = new Point(0.5, 0.55);
        const double shrink = 0.78;
        var inner = new Point[pentagon.Length];
        for (int i = 0; i < pentagon.Length; i++)
        {
            inner[i] = center + (pentagon[i] - center) * shrink;
        }

        return IconBuilder
            .BuildIcon()
            .Polyline(IconBuilder.AccentThickness, nameof(AppTheme.Ink), pentagon, isClosed: true)
            .Polyline(IconBuilder.AccentThickness, nameof(AppTheme.LineAccent), inner, isClosed: true)
            .Canvas;
    }
}
