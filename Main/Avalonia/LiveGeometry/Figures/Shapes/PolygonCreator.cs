using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;
using Avalonia.Media;

namespace LiveGeometry;

[Category(BehaviorCategories.Shapes)]
[Order(3)]
public class PolygonCreator : ShapeCreator
{
    protected int FoundDependenciesMinimum = 3; // Includes TempPoint

    /// <summary>Enough vertices to close (the point following the cursor doesn't count)</summary>
    bool CanClose
    {
        get
        {
            return FoundDependencies.Count >= FoundDependenciesMinimum + 1 // Add 1 for TempPoint
                && FoundDependencies.All(f => f is IPoint);
        }
    }

    /// <summary>Makes the polygon from the vertices so far, if there are enough</summary>
    bool TryClose()
    {
        if (!CanClose)
        {
            return false;
        }

        RemoveIntermediateFigureIfNecessary();
        RemoveTempPointIfNecessary();
        AddFiguresAndRestart();
        return true;
    }

    protected override void Click(Point coordinates)
    {
        // the point the hover showed, as the base click would find it
        var point = LookForExpectedDependencyUnderCursor(ClickedUnconstrainedCoordinates);
        if (FinishesOn(point) && TryClose())
        {
            return;
        }

        base.Click(coordinates);
    }

    /// <summary>A click on a vertex closes the polygon once there are enough</summary>
    protected override bool FinishesOn(IFigure figure)
    {
        return CanClose && figure is IPoint && figure != TempPoint && FoundDependencies.Contains(figure);
    }

    protected override string GetHoverStatus(MouseEventArgs e)
    {
        var point = LookForExpectedDependencyUnderCursor(Coordinates(e, false, false, false));
        return FinishesOn(point) ? FinishStatus : null;
    }

    /// <summary>What the status says over a vertex a click finishes the figure on</summary>
    protected virtual string FinishStatus => "Click this point to close the polygon.";

    protected override void AddDependency(Point coordinates)
    {
        base.AddDependency(coordinates);
        UpdateStartHalo();
    }

    public override void Stopping()
    {
        base.Stopping();
        RemoveStartHalo();
    }

    #region The first vertex, marked

    // Once a click on the first vertex closes the polygon, the vertex has a green halo for
    // as long as the polygon is drawn: where to come back to. Under the hover's halo, which
    // is translucent as it is, so the two mix where the cursor is over it.
    Avalonia.Controls.Shapes.Ellipse startHalo;
    PointBase startPoint;
    Avalonia.Controls.Canvas startHaloCanvas;

    void UpdateStartHalo()
    {
        var first = CanClose ? FoundDependencies[0] as PointBase : null;
        if (first == startPoint)
        {
            return;
        }

        RemoveStartHalo();
        if (first?.Shape == null || Drawing?.Canvas == null)
        {
            return;
        }

        startPoint = first;
        startHaloCanvas = Drawing.Canvas;
        startHalo = ClickPreview.CreatePointHalo(first, ClickPreview.StartPointHaloBrush);
        startHalo.IsHitTestVisible = false;
        // under the hover's halo
        startHalo.ZIndex--;
        startHaloCanvas.Children.Add(startHalo);
        first.Shape.PropertyChanged += StartPointShape_PropertyChanged;
    }

    void RemoveStartHalo()
    {
        if (startPoint != null)
        {
            startPoint.Shape.PropertyChanged -= StartPointShape_PropertyChanged;
            startHaloCanvas.Children.Remove(startHalo);
        }

        startPoint = null;
        startHalo = null;
        startHaloCanvas = null;
    }

    // the point is placed anew when the view moves or zooms, and so is its halo
    void StartPointShape_PropertyChanged(object sender, AvaloniaPropertyChangedEventArgs e)
    {
        ClickPreview.PlacePointHalo(startHalo, startPoint);
    }

    #endregion

    /// <summary>
    /// Like in the original DG, a right-click closes the polygon once it has enough
    /// vertices; so does Enter.
    /// </summary>
    public override void MouseRightClick(object sender, MouseButtonEventArgs e)
    {
        if (TryClose())
        {
            return;
        }

        base.MouseRightClick(sender, e);
    }

    public override void KeyDown(object sender, Avalonia.Input.KeyEventArgs e)
    {
        if (e.Key == Avalonia.Input.Key.Enter && TryClose())
        {
            e.Handled = true;
            return;
        }

        base.KeyDown(sender, e);
    }

    public override string ConstructionHintText(Drawing.ConstructionStepCompleteEventArgs args)
    {
        if (CanClose)
        {
            return "Press Enter when done, or click more points.";
        }

        return base.ConstructionHintText(args);
    }

    // the length of one side means nothing for a polygon
    protected override void ShowCreatedFigure(IList<IFigure> figures)
    {
    }

    protected override DependencyList InitExpectedDependencies()
    {
        return null;
    }

    protected override bool CanCreateTempResults()
    {
        return FoundDependencies.Count >= 3;
    }

    protected override Type GetExpectedDependencyType()
    {
        return typeof(IPoint);
    }

    protected override IEnumerable<IFigure> CreateFigures()
    {
        var polygon = Factory.CreatePolygon(Drawing, FoundDependencies);
        //if (Style != null)
        //{
        //    polygon.Style = Style;
        //}
        yield return polygon;
        for (int i = 0; i < FoundDependencies.Count; i++)
        {
            // get two consecutive vertices of the polygon
            int j = (i + 1) % FoundDependencies.Count;
            IPoint p1 = FoundDependencies[i] as IPoint;
            IPoint p2 = FoundDependencies[j] as IPoint;
            // try to find if there is already a line connecting them
            if (Drawing.Figures.FindLine(p1, p2) == null)
            {
                // if not, create a new segment
                var segment = Factory.CreateSegment(Drawing, p1, p2);
                yield return segment;
            }
        }
    }

    protected override IFigure CreateIntermediateFigure()
    {
        if (!FoundDependencies.All(f => f is IPoint))
        {
            return null;
        }
        if (FoundDependencies.Count == 2)
        {
            return Factory.CreateSegment(Drawing, FoundDependencies);
        }
        return null;
    }

    public override string Name => "Polygon";

    public override string HintText => "Click points to construct a polygon. Press Enter (or click the first point again) to close it.";

    public override FrameworkElement CreateIcon()
    {
        return IconBuilder
            .BuildIcon()
            .Polygon(
                nameof(AppTheme.ShapeIconFill),
                nameof(AppTheme.ShapeOutline),
                new Point(0.2, 0.4),
                new Point(0.3, 0.8),
                new Point(0.7, 0.8),
                new Point(0.8, 0.4),
                new Point(0.6, 0.2))
            .Canvas;
    }
}