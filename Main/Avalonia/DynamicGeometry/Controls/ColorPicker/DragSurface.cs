using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// An area the user presses and drags on. Reports the pointer position normalized to 0..1
/// in both directions (clamped, so dragging past the edge pins to it) and whether the
/// drag is still going on. Put the visuals in as children; they are not hit-testable.
/// </summary>
public class DragSurface : Panel
{
    public DragSurface()
    {
        // something to hit even where no child paints
        Background = Brushes.Transparent;
        Cursor = new Cursor(StandardCursorType.Cross);
    }

    /// <summary>(x, y) in 0..1, and true while the button is still down</summary>
    public event Action<Point, bool> Dragged;

    bool isDragging;

    protected override void OnPointerPressed(PointerPressedEventArgs e)
    {
        base.OnPointerPressed(e);
        if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            return;
        }

        isDragging = true;
        e.Pointer.Capture(this);
        Report(e, isStillDragging: true);
        e.Handled = true;
    }

    protected override void OnPointerMoved(PointerEventArgs e)
    {
        base.OnPointerMoved(e);
        if (isDragging)
        {
            Report(e, isStillDragging: true);
        }
    }

    protected override void OnPointerReleased(PointerReleasedEventArgs e)
    {
        base.OnPointerReleased(e);
        if (isDragging)
        {
            isDragging = false;
            e.Pointer.Capture(null);
            Report(e, isStillDragging: false);
        }
    }

    protected override void OnPointerCaptureLost(PointerCaptureLostEventArgs e)
    {
        base.OnPointerCaptureLost(e);
        isDragging = false;
    }

    void Report(PointerEventArgs e, bool isStillDragging)
    {
        if (Bounds.Width <= 0 || Bounds.Height <= 0)
        {
            return;
        }

        var position = e.GetPosition(this);
        Dragged?.Invoke(
            new Point(
                System.Math.Clamp(position.X / Bounds.Width, 0, 1),
                System.Math.Clamp(position.Y / Bounds.Height, 0, 1)),
            isStillDragging);
    }
}
