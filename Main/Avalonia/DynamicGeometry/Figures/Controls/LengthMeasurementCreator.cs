using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;

namespace DynamicGeometry;

/// <summary>
/// The Length tool: two points for the distance between them, or one click on a segment, a
/// vector, a polyline, a Bezier curve or path, an arc, a sector, a circular segment, a circle,
/// an ellipse or a polygon for its length (<see cref="LengthMeasurement.IsMeasurable"/>). A
/// figure with more than one length (a circle's radius, diameter and circumference) gets the
/// first, and the side panel offers the others (<see cref="LengthMeasurementPanel"/>).
/// </summary>
[Category(BehaviorCategories.Measure)]
[Order(1)]
public class LengthMeasurementCreator : FigureCreator
{
    protected override DependencyList InitExpectedDependencies()
    {
        return DependencyList.PointPoint;
    }

    protected override IEnumerable<IFigure> CreateFigures()
    {
        yield return Factory.CreateLengthMeasurement(Drawing, FoundDependencies);
    }

    /// <summary>
    /// A figure with a length under the cursor with no point on top of it: a first click
    /// measures it instead of putting a point on it. Once the first point of a distance is
    /// there the next click is its other end, on the figure if it is on one (a click on a
    /// segment then measured the segment, and the first point was left behind on its own).
    /// </summary>
    protected override IReadOnlyList<IFigure> FindFiguresInsteadOfPoint(Point unconstrainedCoordinates)
    {
        if (FoundDependencies.Count > 0 || Drawing.Figures.HitTest<IPoint>(unconstrainedCoordinates) != null)
        {
            return System.Array.Empty<IFigure>();
        }

        // (only what is in view: without the last two tests a click on empty paper measured
        // a hidden segment lying there - the long axis every ellipse keeps; and a path is hit
        // by a side or its inside, parts of it, while the path is what is measured)
        return Drawing.Figures
            .HitTestAll(
                unconstrainedCoordinates,
                f => LengthMeasurement.IsMeasurable(FigureParts.Whole(f))
                    && f.Visible
                    && f.IsHitTestVisible
                    && !FigureParts.Whole(f).DependsOn(TempPoint))
            .Select(FigureParts.Whole)
            .Distinct()
            .ToList();
    }

    public override void MouseDown(object sender, MouseButtonEventArgs e)
    {
        // where the cursor is, as the hover preview asks: not where Shift snaps it to
        var underMouse = FindFigureInsteadOfPoint(Coordinates(e, false, false, false));
        if (underMouse != null)
        {
            FoundDependencies.Clear();
            FoundDependencies.Add(underMouse);
            RemoveIntermediateFigureIfNecessary();
            RemoveTempPointIfNecessary();
            AddFiguresAndRestart();
            return;
        }

        base.MouseDown(sender, e);
    }

    /// <summary>A figure with more than one length: the panel to pick another</summary>
    protected override void ShowCreatedFigure(IList<IFigure> figures)
    {
        var measurement = figures.OfType<LengthMeasurement>().FirstOrDefault();
        if (measurement != null && measurement.Quantities.Count > 1)
        {
            LengthMeasurementPanel.Show(measurement);
        }
    }

    public override string Name
    {
        get { return "Length"; }
    }

    public override string HintText
    {
        get { return "Click two points for the distance between them, or a segment, an arc, a curve, a circle or a shape for its length."; }
    }

    public override FrameworkElement CreateIcon()
    {
        var builder = IconBuilder
            .BuildIcon()
            .Polygon(
                nameof(AppTheme.RulerFill),
                nameof(AppTheme.RulerOutline),
                new Point(0.1, 0.8),
                new Point(0.3, 1),
                new Point(1, 0.3),
                new Point(0.8, 0.1))
            .Line(0, 0.7, 0.7, 0);
        for (double i = 0.2; i <= 0.7; i += 0.1)
        {
            builder.Line(nameof(AppTheme.ScaleMarks), i, 0.9 - i, i + 0.1, 1 - i);
        }

        for (double i = 0.15; i <= 0.75; i += 0.1)
        {
            builder.Line(nameof(AppTheme.ScaleMarks), i, 0.9 - i, i + 0.05, 0.95 - i);
        }

        return builder.Canvas;
    }
}
