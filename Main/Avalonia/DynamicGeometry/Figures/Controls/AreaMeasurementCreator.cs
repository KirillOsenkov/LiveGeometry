using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;
using Avalonia.Controls;

namespace DynamicGeometry;

[Category(BehaviorCategories.Measure)]
[Order(3)]
public class AreaMeasurementCreator : FigureCreator
{
    /// <summary>
    /// A shape under the cursor with no point on top of it, before any point of a list:
    /// a click measures it. Found here, where the hover looks too, so that the cursor is
    /// the hand and the shape has its halo over a shape (it was the cross of a free point
    /// over a filled polygon, which the click then measured); and not once points are
    /// being listed, when a click inside a shape is the next point (it threw the points
    /// listed away and measured the shape).
    /// </summary>
    protected override IReadOnlyList<IFigure> FindFiguresInsteadOfPoint(Point unconstrainedCoordinates)
    {
        if (FoundDependencies.Count > 0 || Drawing.Figures.HitTest<IPoint>(unconstrainedCoordinates) != null)
        {
            return System.Array.Empty<IFigure>();
        }

        // (a path is hit by a side or its inside, parts of it, while the path is what is measured)
        return Drawing.Figures
            .HitTestAll(
                unconstrainedCoordinates,
                f => AreaMeasurement.IsMeasurable(FigureParts.Whole(f))
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
        var shape = FindFigureInsteadOfPoint(Coordinates(e, false, false, false));
        if (shape != null)
        {
            FoundDependencies.Clear();
            FoundDependencies.Add(shape);
            RemoveIntermediateFigureIfNecessary();
            RemoveTempPointIfNecessary();
            AddFiguresAndRestart();
            return;
        }

        base.MouseDown(sender, e);
    }

    protected override void Click(Point coordinates)
    {
        if (Drawing.Figures.HitTest(coordinates) is PointBase point && FoundDependencies.Contains(point) && TryFinish())
        {
            return;
        }

        base.Click(coordinates);
    }

    /// <summary>A click on a shape measures it: the choice is between the shapes there, not points in them</summary>
    protected override IReadOnlyList<object> FindClickOptions(MouseEventArgs e)
    {
        var shapes = FindFiguresInsteadOfPoint(Coordinates(e, false, false, false));
        return shapes.Count > 0 ? shapes.ToList<object>() : base.FindClickOptions(e);
    }

    /// <summary>Three points or more (and the one following the cursor): enough for an area</summary>
    bool CanFinish
    {
        get
        {
            return FoundDependencies.Count >= 4 && FoundDependencies.All(f => f is IPoint);
        }
    }

    bool TryFinish()
    {
        if (!CanFinish)
        {
            return false;
        }

        RemoveIntermediateFigureIfNecessary();
        RemoveTempPointIfNecessary();
        AddFiguresAndRestart();
        return true;
    }

    /// <summary>
    /// Enter and a right click finish the list of points, as they close a polygon
    /// (<see cref="PolygonCreator"/>). The only way was to click one of the points a
    /// second time, which nothing said; Enter did nothing and a right click threw the
    /// points away.
    /// </summary>
    public override void MouseRightClick(object sender, MouseButtonEventArgs e)
    {
        if (TryFinish())
        {
            return;
        }

        base.MouseRightClick(sender, e);
    }

    public override void KeyDown(object sender, Avalonia.Input.KeyEventArgs e)
    {
        if (e.Key == Avalonia.Input.Key.Enter && TryFinish())
        {
            e.Handled = true;
            return;
        }

        base.KeyDown(sender, e);
    }

    public override string ConstructionHintText(Drawing.ConstructionStepCompleteEventArgs args)
    {
        if (CanFinish)
        {
            return "Press Enter when done (or click one of the points again), or click more points.";
        }

        return base.ConstructionHintText(args);
    }

    protected override bool CanCreateTempResults()
    {
        return false;
    }

    protected override DependencyList InitExpectedDependencies()
    {
        return null;
    }

    protected override System.Type GetExpectedDependencyType()
    {
        return typeof(IPoint);
    }

    protected override IEnumerable<IFigure> CreateFigures()
    {
        yield return Factory.CreateAreaMeasurement(Drawing, FoundDependencies);
    }

    protected override IFigure CreateIntermediateFigure()
    {
        if (!FoundDependencies.All(f => f is IPoint))
        {
            return null;
        }
        else if (FoundDependencies.Count >= 3)
        {
            var result = Factory.CreateAreaMeasurement(Drawing, FoundDependencies);
            return result;
        }
        return null;
    }

    public override string Name
    {
        get { return "Area"; }
    }

    public override string HintText
    {
        get
        {
            return "Click a polygon, ellipse, circle, sector or closed path, or a list of points, to measure its area.";
        }
    }

    public override FrameworkElement CreateIcon()
    {
        // A hatched region: all drawn, no text (text in an icon doesn't scale with it, and the
        // tab header shows a copy of this icon at another font size).
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
            .Polygon(
                nameof(AppTheme.AreaFill),
                nameof(AppTheme.Ink),
                pentagon);

        // hatching: the parts of the lines x + y = c that are inside the pentagon
        var hatchColor = nameof(AppTheme.AreaHatch);
        for (double c = 0.55; c < 1.7; c += 0.2)
        {
            var ends = new List<Point>();
            for (int i = 0; i < pentagon.Length; i++)
            {
                var a = pentagon[i];
                var b = pentagon[(i + 1) % pentagon.Length];
                var sideA = a.X + a.Y - c;
                var sideB = b.X + b.Y - c;
                if (sideA * sideB < 0)
                {
                    var t = sideA / (sideA - sideB);
                    ends.Add(new Point(a.X + (b.X - a.X) * t, a.Y + (b.Y - a.Y) * t));
                }
            }

            if (ends.Count == 2)
            {
                builder.Line(
                    hatchColor,
                    ends[0].X,
                    ends[0].Y,
                    ends[1].X,
                    ends[1].Y);
            }
        }

        return builder.Canvas;
    }
}