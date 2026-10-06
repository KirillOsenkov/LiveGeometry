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
/// too), Enter, a right click or a double click leaves it open. Alt+click between two anchors
/// takes the handles, in the order of a Bézier curve's points: the first the out handle of
/// the anchor before, the second the in handle of the anchor after - a point of the drawing
/// where the Point tool would take or make one (an existing point, a point on a figure, a
/// crossing), else an ordinary handle at the click.
/// </summary>
[Category(BehaviorCategories.Shapes)]
[Order(5)]
public class BezierPathCreator : FigureCreator
{
    // the handles of the anchors found so far
    readonly List<BezierPath.HandleSpec> ins = new List<BezierPath.HandleSpec>();
    readonly List<BezierPath.HandleSpec> outs = new List<BezierPath.HandleSpec>();

    // the in handle Alt+clicked for the anchor still to come: a point, or where it was
    // clicked (its offset is from that anchor, which isn't there yet)
    IPoint pendingPoint;
    Point? pendingSpot;

    // the handles Alt+clicked since the last anchor
    int handleClicks;

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
        ins.Clear();
        outs.Clear();
        ForgetPending();
        handleClicks = 0;
        pulledAnchor = -1;
        pulling = false;
        closing = false;
    }

    void ForgetPending()
    {
        pendingPoint = null;
        pendingSpot = null;
    }

    /// <summary>The in handle Alt+clicked for this anchor, as a handle of it; an ordinary one on it when there is none</summary>
    BezierPath.HandleSpec PendingFor(IPoint anchor)
    {
        if (pendingPoint != null)
        {
            return new BezierPath.HandleSpec(default, pendingPoint);
        }

        return new BezierPath.HandleSpec(pendingSpot.HasValue ? pendingSpot.Value.Minus(anchor.Coordinates) : default, null);
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
        if (figure == TempPoint || !(figure is IPoint anchor) || ins.Count >= AnchorsFound)
        {
            return;
        }

        // an anchor: it takes the in handle Alt+clicked for it
        ins.Add(PendingFor(anchor));
        outs.Add(default);
        ForgetPending();
        handleClicks = 0;
    }

    protected override IEnumerable<IFigure> CreateFigures()
    {
        // (the preview's last anchor, the point following the cursor, has none of its own:
        // it shows the in handle Alt+clicked for the next anchor)
        var anchors = FoundDependencies.ToList();
        var anchorIns = anchors.Select((anchor, i) => i < ins.Count ? ins[i] : PendingFor((IPoint)anchor)).ToList();
        var anchorOuts = anchors.Select((anchor, i) => i < outs.Count ? outs[i] : default).ToList();
        yield return BezierPath.Create(
            Drawing,
            anchors,
            anchorIns,
            anchorOuts,
            holes: Array.Empty<IFigure>(),
            closed: finishingClosed,
            filled: finishingClosed);
    }

    void RebuildPreview()
    {
        RemoveTempResultsIfNecessary();
        if (CanCreateTempResults())
        {
            CreateTempResults();
        }
    }

    /// <summary>
    /// Alt+click: the next handle - the out handle of the last anchor, then the in handle of
    /// the next one - a point the Point tool would take or make here, else an ordinary handle
    /// where the click is
    /// </summary>
    void TakeHandle(Point coordinates)
    {
        if (AnchorsFound == 0)
        {
            Drawing.RaiseStatusNotification("Click the first point of the path: " + KeyNames.Alt + "+click takes handles after it.");
            return;
        }

        if (handleClicks >= 2)
        {
            Drawing.RaiseStatusNotification("Both handles between these points are taken: click the next point.");
            return;
        }

        var placement = FindPointPlacement(ClickedUnconstrainedCoordinates, coordinates);
        var point = placement?.ExistingPoint;
        if (point == null && placement != null && placement.IsDependent)
        {
            point = placement.Create(Drawing);
            Actions.Add(Drawing, point);
        }

        var spot = placement?.Coordinates ?? coordinates;
        if (handleClicks == 0)
        {
            int last = AnchorsFound - 1;
            var anchor = (IPoint)FoundDependencies[last];
            outs[last] = point != null
                ? new BezierPath.HandleSpec(default, point)
                : new BezierPath.HandleSpec(spot.Minus(anchor.Coordinates), null);
        }
        else
        {
            pendingPoint = point;
            pendingSpot = point == null ? spot : null;
        }

        handleClicks++;
        RebuildPreview();
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

        // closed, the first anchor takes the in handle Alt+clicked before the click on it
        if (closed && (pendingPoint != null || pendingSpot.HasValue))
        {
            ins[0] = PendingFor((IPoint)FoundDependencies[0]);
        }

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
        if (IsAltPressed())
        {
            TakeHandle(coordinates);
            return;
        }

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

            // an in handle Alt+clicked at a spot stays there while the anchor it is for
            // follows the cursor
            if (pendingSpot.HasValue && TempPoint != null && Preview is BezierPath preview)
            {
                preview.SetHandleOffsets(preview.AnchorCount - 1, pendingSpot.Value.Minus(TempPoint.Coordinates), default);
            }

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

        // (a handle that is a point stays that point)
        var anchor = (IPoint)FoundDependencies[pulledAnchor];
        var offset = Coordinates(e).Minus(anchor.Coordinates);
        if (outs[pulledAnchor].Point == null)
        {
            outs[pulledAnchor] = new BezierPath.HandleSpec(offset, null);
        }

        if (ins[pulledAnchor].Point == null)
        {
            ins[pulledAnchor] = new BezierPath.HandleSpec(offset.Minus(), null);
        }

        // the point following the cursor waits at the anchor whose handles are pulled: a
        // piece to the cursor would be drawn from it meanwhile (when closing: the first one)
        if (TempPoint is IMovable following)
        {
            following.MoveTo(anchor.Coordinates);
        }

        Preview?.SetHandleOffsets(pulledAnchor, ins[pulledAnchor].Offset, outs[pulledAnchor].Offset);
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
        var handles = " " + KeyNames.Alt + "+click takes a handle (a point, or a spot).";
        if (AnchorsFound >= 2)
        {
            return "Click the first point to close the path, press Enter to leave it open, or click or drag the next point." + handles;
        }

        return "Click the next point, or press and drag to pull out its handles." + handles;
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
        // a heart through four anchors (the dip at the top, the two sides, the tip), of which
        // only the left one is drawn, with two handles: small gray squares on dotted lines,
        // as on the paper (all four, and their handles, crowded the icon). The handles are a
        // picture of handles, placed to read well, not the ones the curve is drawn with.
        var builder = IconBuilder.BuildIcon();
        double size = builder.Canvas.Width;
        Point At(double x, double y)
        {
            return new Point(size * x, size * y);
        }

        var left = new Point(0.05, 0.5);
        var top = new Point(0.5, 0.35);
        var right = new Point(0.95, 0.5);
        var tip = new Point(0.5, 1);
        var handles = new[]
        {
            (Anchor: left, Handle: new Point(-0.05, 0.25)),
            // (a pixel on the screen to the right of the mirror image of the other: the two
            // snap to pixels differently, and at 0.15 they looked lopsided)
            (Anchor: left, Handle: new Point(0.165, 0.75))
        };
        var figure = new PathFigure()
        {
            StartPoint = At(left.X, left.Y),
            IsClosed = true,
            IsFilled = true,
            Segments = new PathSegmentCollection()
            {
                new BezierSegment() { Point1 = At(0, 0.2), Point2 = At(0.35, 0.15), Point3 = At(top.X, top.Y) },
                new BezierSegment() { Point1 = At(0.65, 0.15), Point2 = At(1, 0.2), Point3 = At(right.X, right.Y) },
                new BezierSegment() { Point1 = At(0.9, 0.65), Point2 = At(0.6, 0.8), Point3 = At(tip.X, tip.Y) },
                new BezierSegment() { Point1 = At(0.4, 0.8), Point2 = At(0.1, 0.65), Point3 = At(left.X, left.Y) }
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

        foreach (var (anchor, handle) in handles)
        {
            var line = new Avalonia.Controls.Shapes.Line()
            {
                StartPoint = At(anchor.X, anchor.Y),
                EndPoint = At(handle.X, handle.Y),
                StrokeThickness = 1,
                StrokeDashArray = new Avalonia.Collections.AvaloniaList<double>() { 1, 1 },
                Opacity = 0.6
            };
            line.BindTheme(AvaloniaShape.StrokeProperty, nameof(AppTheme.Ink));
            builder.Canvas.Children.Add(line);
        }

        foreach (var (_, handle) in handles)
        {
            // tiny, solid: a rim would be all there is of it
            const double side = 2.5;
            var square = new Avalonia.Controls.Shapes.Rectangle()
            {
                Width = side,
                Height = side,
                Opacity = 0.6
            };
            square.BindTheme(AvaloniaShape.FillProperty, nameof(AppTheme.Ink));
            Avalonia.Controls.Canvas.SetLeft(square, size * handle.X - side / 2);
            Avalonia.Controls.Canvas.SetTop(square, size * handle.Y - side / 2);
            builder.Canvas.Children.Add(square);
        }

        return builder
            .Point(left.X, left.Y)
            .Canvas;
    }
}
