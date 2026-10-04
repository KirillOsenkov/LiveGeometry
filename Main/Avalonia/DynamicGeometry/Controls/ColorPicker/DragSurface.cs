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
/// Pressed, it takes the keyboard, and reports the arrow keys as steps (<see cref="Stepped"/>).
/// </summary>
public class DragSurface : Panel
{
    public DragSurface()
    {
        // something to hit even where no child paints
        Background = Brushes.Transparent;
        Cursor = new Cursor(StandardCursorType.Cross);
        Focusable = true;
    }

    /// <summary>(x, y) in 0..1, and true while the button is still down</summary>
    public event Action<Point, bool> Dragged;

    /// <summary>
    /// An arrow key while the surface has the keyboard: the steps across and down, one, or
    /// <see cref="LargeStep"/> with Shift, for whoever uses the surface to make a small change of
    /// </summary>
    public event Action<int, int> Stepped;

    /// <summary>How many steps an arrow key with Shift makes</summary>
    public const int LargeStep = 10;

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

        // the arrow keys go on from here
        Focus(NavigationMethod.Pointer);
        Report(e, isStillDragging: true);
        e.Handled = true;
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        if (e.Handled || (e.KeyModifiers & ~KeyModifiers.Shift) != KeyModifiers.None)
        {
            return;
        }

        int size = e.KeyModifiers == KeyModifiers.Shift ? LargeStep : 1;
        (int across, int down) = e.Key switch
        {
            Key.Left => (-size, 0),
            Key.Right => (size, 0),
            Key.Up => (0, -size),
            Key.Down => (0, size),
            _ => (0, 0)
        };
        if (across == 0 && down == 0)
        {
            return;
        }

        // at the edge too: the side panel's scroll viewer would take the key and scroll
        e.Handled = true;
        Stepped?.Invoke(across, down);
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
