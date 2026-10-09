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
        // The Area tool's pentagon, hollow, its outline marked off like a ruler: the outline
        // is what is measured.
        var pentagon = new[]
        {
            new Point(0.5, 0.08),
            new Point(0.93, 0.39),
            new Point(0.77, 0.9),
            new Point(0.23, 0.9),
            new Point(0.07, 0.39)
        };

        var builder = IconBuilder
            .BuildIcon()
            .Polyline(IconBuilder.AccentThickness, nameof(AppTheme.Ink), pentagon, isClosed: true);

        // ticks across each side, at a third and two thirds of it, pointing inward
        const double tickLength = 0.09;
        for (int i = 0; i < pentagon.Length; i++)
        {
            var from = pentagon[i];
            var to = pentagon[(i + 1) % pentagon.Length];
            var along = to - from;
            double length = System.Math.Sqrt(along.X * along.X + along.Y * along.Y);
            var inward = new Point(-along.Y / length, along.X / length);
            foreach (double fraction in new[] { 1.0 / 3, 2.0 / 3 })
            {
                var start = from + along * fraction;
                var end = start + inward * tickLength;
                builder.Line(nameof(AppTheme.ScaleMarks), start.X, start.Y, end.X, end.Y);
            }
        }

        return builder.Canvas;
    }
}
