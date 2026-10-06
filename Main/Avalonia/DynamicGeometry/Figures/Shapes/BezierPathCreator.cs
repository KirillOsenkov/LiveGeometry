using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;
using Avalonia.Media;
using AvaloniaPath = Avalonia.Controls.Shapes.Path;
using AvaloniaShape = Avalonia.Controls.Shapes.Shape;

namespace DynamicGeometry;

/// <summary>
/// Makes a <see cref="BezierPath"/> as the pen of a drawing program does: a click puts an
/// anchor (wherever the Point tool would put a point) with its handles on it, a press and drag
/// pulls the handles out, the out one under the cursor and the in one the opposite way, until
/// the release. A click on the first anchor closes the path (a drag from it sets its handles
/// too), Enter, a right click or a double click leaves it open.
/// </summary>
[Category(BehaviorCategories.Shapes)]
[Order(5)]
public class BezierPathCreator : FigureCreator
{
    // the handles of the anchors found so far, as offsets from them
    readonly List<Point> inOffsets = new List<Point>();
    readonly List<Point> outOffsets = new List<Point>();

    // the anchor (its index among those found) whose handles the press pulls out; -1 for none
    int pulledAnchor = -1;

    // whether the press has gone far enough to be a pull rather than a click
    bool pulling;

    // the press was on the first anchor: the path closes at the release
    bool closing;

    // what the path made at the end is: closed and filled, or open and not
    bool finishingClosed;

    public override void Started()
    {
        base.Started();
        inOffsets.Clear();
        outOffsets.Clear();
        pulledAnchor = -1;
        pulling = false;
        closing = false;
    }

    protected override DependencyList InitExpectedDependencies()
    {
        return null;
    }

    // a point, as many as it takes
    protected override Type GetExpectedDependencyType()
    {
        return typeof(IPoint);
    }

    // the path so far, to the point following the cursor
    protected override bool CanCreateTempResults()
    {
        return FoundDependencies.Count >= 2;
    }

    /// <summary>The anchors found, without the point following the cursor</summary>
    int AnchorsFound
    {
        get
        {
            return FoundDependencies.Count(found => found != TempPoint);
        }
    }

    protected override void AddFoundDependency(IFigure figure)
    {
        base.AddFoundDependency(figure);
        while (figure != TempPoint && inOffsets.Count < AnchorsFound)
        {
            inOffsets.Add(default);
            outOffsets.Add(default);
        }
    }

    protected override IEnumerable<IFigure> CreateFigures()
    {
        var anchors = FoundDependencies.ToList();
        var ins = anchors.Select((anchor, i) => i < inOffsets.Count ? inOffsets[i] : default).ToList();
        var outs = anchors.Select((anchor, i) => i < outOffsets.Count ? outOffsets[i] : default).ToList();
        yield return BezierPath.Create(
            Drawing,
            anchors,
            ins,
            outs,
            closed: finishingClosed,
            filled: finishingClosed);
    }

    /// <summary>The path so far, while there is one</summary>
    BezierPath Preview
    {
        get
        {
            return TempResults.OfType<BezierPath>().FirstOrDefault();
        }
    }

    /// <summary>The path from the anchors found, open or closed; false while there are fewer than two</summary>
    bool TryFinish(bool closed)
    {
        if (AnchorsFound < 2)
        {
            return false;
        }

        pulledAnchor = -1;
        pulling = false;
        closing = false;
        finishingClosed = closed;
        try
        {
            RemoveIntermediateFigureIfNecessary();
            RemoveTempPointIfNecessary();
            AddFiguresAndRestart();
        }
        finally
        {
            finishingClosed = false;
        }

        return true;
    }

    // nothing to adjust right after: no length, no tied values
    protected override void ShowCreatedFigure(IList<IFigure> figures)
    {
    }

    protected override void Click(Point coordinates)
    {
        var point = Drawing.Figures.HitTest<IPoint>(ClickedUnconstrainedCoordinates);
        if (point != null && AnchorsFound >= 2 && point == FoundDependencies[0])
        {
            // closed at the release, after a drag sets the first anchor's handles
            closing = true;
            pulledAnchor = 0;
            return;
        }

        int before = AnchorsFound;
        base.Click(coordinates);
        if (AnchorsFound > before)
        {
            pulledAnchor = AnchorsFound - 1;
        }
    }

    public override void MouseDown(object sender, MouseButtonEventArgs e)
    {
        pulledAnchor = -1;
        pulling = false;
        closing = false;

        // the second click of a double click: the path ends there, open
        if (e.ClickCount == 2 && TryFinish(closed: false))
        {
            return;
        }

        base.MouseDown(sender, e);
    }

    public override void MouseMove(object sender, MouseEventArgs e)
    {
        if (pulledAnchor < 0 || !IsMouseButtonDown || pulledAnchor >= AnchorsFound)
        {
            base.MouseMove(sender, e);
            return;
        }

        // a press with a wobble is still a click: the handles stay on the anchor
        var coordinateSystem = Drawing.CoordinateSystem;
        if (!pulling)
        {
            var wobble = coordinateSystem.ToPhysical(Coordinates(e, false, false, false))
                .Distance(coordinateSystem.ToPhysical(ClickedUnconstrainedCoordinates));
            if (wobble < Dragger.DragThreshold)
            {
                return;
            }

            pulling = true;
        }

        var anchor = (IPoint)FoundDependencies[pulledAnchor];
        var offset = Coordinates(e).Minus(anchor.Coordinates);
        outOffsets[pulledAnchor] = offset;
        inOffsets[pulledAnchor] = offset.Minus();

        // the point following the cursor waits at the anchor whose handles are pulled: a
        // piece to the cursor would be drawn from it meanwhile (when closing: the first one)
        if (TempPoint is IMovable following)
        {
            following.MoveTo(anchor.Coordinates);
        }

        Preview?.SetHandleOffsets(pulledAnchor, inOffsets[pulledAnchor], outOffsets[pulledAnchor]);
        Drawing.Recalculate();
    }

    // Not the base's: a release far from its press is not a second click here, it ends a pull
    public override void MouseUp(object sender, MouseButtonEventArgs e)
    {
        IsMouseButtonDown = false;
        bool close = closing;
        pulledAnchor = -1;
        pulling = false;
        closing = false;
        if (close)
        {
            TryFinish(closed: true);
        }
    }

    public override void MouseRightClick(object sender, MouseButtonEventArgs e)
    {
        if (TryFinish(closed: false))
        {
            return;
        }

        base.MouseRightClick(sender, e);
    }

    public override void KeyDown(object sender, Avalonia.Input.KeyEventArgs e)
    {
        if (e.Key == Avalonia.Input.Key.Enter && TryFinish(closed: false))
        {
            e.Handled = true;
            return;
        }

        base.KeyDown(sender, e);
    }

    public override string ConstructionHintText(Drawing.ConstructionStepCompleteEventArgs args)
    {
        if (AnchorsFound >= 2)
        {
            return "Click the first point to close the path, press Enter to leave it open, or click or drag the next point.";
        }

        return "Click the next point, or press and drag to pull out its handles.";
    }

    public override string Name
    {
        get
        {
            return "Bezier path";
        }
    }

    public override string HintText
    {
        get
        {
            return "Click points of a path, or press and drag to pull out handles. Click the first point again to close the path, or press Enter.";
        }
    }

    public override FrameworkElement CreateIcon()
    {
        // a heart: two anchors, top and bottom, and their four handles
        var builder = IconBuilder.BuildIcon();
        double size = builder.Canvas.Width;
        Point At(double x, double y)
        {
            return new Point(size * x, size * y);
        }

        var figure = new PathFigure()
        {
            StartPoint = At(0.5, 0.33),
            IsClosed = true,
            IsFilled = true,
            Segments = new PathSegmentCollection()
            {
                new BezierSegment() { Point1 = At(0.62, 0.09), Point2 = At(0.91, 0.45), Point3 = At(0.5, 0.9) },
                new BezierSegment() { Point1 = At(0.09, 0.45), Point2 = At(0.38, 0.09), Point3 = At(0.5, 0.33) }
            }
        };
        var heart = new AvaloniaPath()
        {
            Data = new PathGeometry() { Figures = new PathFigureCollection() { figure } },
            StrokeThickness = 1,
            StrokeJoin = PenLineJoin.Round
        };
        heart.BindTheme(AvaloniaShape.FillProperty, nameof(AppTheme.ShapeIconFill));
        heart.BindTheme(AvaloniaShape.StrokeProperty, nameof(AppTheme.ShapeOutline));
        builder.Canvas.Children.Add(heart);
        return builder
            .DashedLine(nameof(AppTheme.Ink), 0.5, 0.33, 0.62, 0.09)
            .DashedLine(nameof(AppTheme.Ink), 0.5, 0.33, 0.38, 0.09)
            .DashedLine(nameof(AppTheme.Ink), 0.5, 0.9, 0.91, 0.45)
            .DashedLine(nameof(AppTheme.Ink), 0.5, 0.9, 0.09, 0.45)
            .DependentPoint(0.62, 0.09)
            .DependentPoint(0.38, 0.09)
            .DependentPoint(0.91, 0.45)
            .DependentPoint(0.09, 0.45)
            .Point(0.5, 0.33)
            .Point(0.5, 0.9)
            .Canvas;
    }
}
