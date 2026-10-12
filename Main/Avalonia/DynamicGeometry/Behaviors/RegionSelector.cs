using System.ComponentModel;
using Avalonia;
using Avalonia.Collections;
using Avalonia.Controls.Shapes;
using Avalonia.Input;

namespace DynamicGeometry;

/// <summary>
/// The tools that select by dragging a region over figures (<see cref="RegionSelection"/>):
/// what the region takes is added to the selection, or with Alt taken out of it, so that a few
/// drags gather a selection. A drag starts anywhere, on a figure too. A click without a drag
/// selects as the Drag tool's does (on nothing: nothing). A right click goes back to Drag,
/// which moves the selection as one piece.
/// </summary>
public abstract class RegionSelector : Behavior
{
    // the press of the drag under way, and its region once it has gone far enough
    IPointer pointer;
    Point pressPosition;
    RegionSelection selection;

    protected abstract SelectionRegion CreateRegion(Point start);

    public override void MouseDown(object sender, MouseButtonEventArgs e)
    {
        // (a press whose release never came)
        AbandonGesture();
        pointer = e.Pointer;
        pressPosition = e.GetPosition(ParentCanvas);
    }

    public override void MouseMove(object sender, MouseEventArgs e)
    {
        if (pointer == null || e.Pointer != pointer)
        {
            return;
        }

        // the release never came (the window lost the mouse): the region stays as it is
        if (!e.GetCurrentPoint(ParentCanvas).Properties.IsLeftButtonPressed)
        {
            Finish(e);
            return;
        }

        var position = e.GetPosition(ParentCanvas);
        if (selection == null)
        {
            if (position.Distance(pressPosition) < Dragger.DragThreshold)
            {
                return;
            }

            selection = new RegionSelection(
                Drawing,
                CreateRegion(pressPosition),
                showsKeyHint: e.Pointer.Type != PointerType.Touch);
        }

        selection.Extend(position, removing: IsAltPressed());
    }

    public override void MouseUp(object sender, MouseButtonEventArgs e)
    {
        if (pointer == null || e.Pointer != pointer)
        {
            return;
        }

        if (selection != null)
        {
            Finish(e);
            return;
        }

        pointer = null;
        var found = Drawing.Figures.HitTest(Coordinates(e, false, false, false));

        // (an axis is paper; a handle of a Bezier path is never selected)
        if (found is AxisLine || found is BezierPath.BezierPathHandle)
        {
            found = null;
        }

        Dragger.SelectClicked(Drawing, found != null ? FigureParts.SelectionTarget(found) : null);
        Drawing.RaiseSelectionChanged(Drawing.GetSelectedFigures());
    }

    void Finish(MouseEventArgs e)
    {
        var finished = selection;
        selection = null;
        pointer = null;
        if (finished != null)
        {
            finished.Extend(e.GetPosition(ParentCanvas), removing: IsAltPressed());
            finished.Finish();
        }
    }

    public override bool AbandonGesture()
    {
        pointer = null;
        if (selection == null)
        {
            return false;
        }

        selection.Cancel();
        selection = null;
        return true;
    }

    public override void Stopping()
    {
        AbandonGesture();
        base.Stopping();
    }

    public override Drawing Drawing
    {
        get
        {
            return base.Drawing;
        }
        set
        {
            if (value != base.Drawing)
            {
                AbandonGesture();
            }

            base.Drawing = value;
        }
    }

    /// <summary>Two points and the segment between them, as the icons show what is selected</summary>
    protected static IconBuilder BuildIcon()
    {
        return IconBuilder.BuildIcon()
            .Line(0.38, 0.6, 0.62, 0.4)
            .Point(0.38, 0.6)
            .Point(0.62, 0.4);
    }

    /// <summary>The line just added to the icon dashed, as a region being selected is no figure</summary>
    protected static void DashLast(IconBuilder builder)
    {
        var shape = (Shape)builder.Canvas.Children[builder.Canvas.Children.Count - 1];
        shape.StrokeDashArray = new AvaloniaList<double>() { 2, 1.5 };
    }
}

[Category(BehaviorCategories.Selection)]
[Order(2)]
public class BoxSelector : RegionSelector
{
    protected override SelectionRegion CreateRegion(Point start)
    {
        return SelectionRegion.Rectangle(start);
    }

    public override FrameworkElement CreateIcon()
    {
        var builder = BuildIcon().Polyline(
            strokeThickness: 1,
            nameof(AppTheme.Ink),
            new[] { new Point(0.15, 0.22), new Point(0.85, 0.22), new Point(0.85, 0.78), new Point(0.15, 0.78) },
            isClosed: true);
        DashLast(builder);
        return builder.Canvas;
    }

    public override string Name
    {
        get { return "Box"; }
    }

    public override string HintText
    {
        get { return "Drag a box over figures to add them to the selection. Hold " + KeyNames.Alt + " to remove them instead."; }
    }
}

[Category(BehaviorCategories.Selection)]
[Order(3)]
public class LassoSelector : RegionSelector
{
    protected override SelectionRegion CreateRegion(Point start)
    {
        return SelectionRegion.Lasso(start);
    }

    public override FrameworkElement CreateIcon()
    {
        // a loop, and the end of the rope hanging from where it closes
        var builder = BuildIcon().Bezier(0.34, 0.8, 0.02, 0.66, 0.12, 0.14, 0.55, 0.14);
        DashLast(builder);
        builder.Bezier(0.55, 0.14, 0.98, 0.14, 0.96, 0.7, 0.34, 0.8);
        DashLast(builder);
        return builder
            .Bezier(0.34, 0.8, 0.22, 0.84, 0.3, 0.92, 0.2, 0.97)
            .Canvas;
    }

    public override string Name
    {
        get { return "Lasso"; }
    }

    public override string HintText
    {
        get { return "Draw a loop around figures to add them to the selection. Hold " + KeyNames.Alt + " to remove them instead."; }
    }
}
