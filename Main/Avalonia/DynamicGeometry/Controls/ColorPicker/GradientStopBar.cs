using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>One color of a gradient and where along it (0..1) the color sits.</summary>
public class ColorStop
{
    public ColorStop(double offset, Color color)
    {
        Offset = offset;
        Color = color;
    }

    public double Offset { get; set; }
    public Color Color { get; set; }
}

/// <summary>
/// A bar showing a gradient, with a handle for every stop underneath. Click a handle to
/// select the stop (whoever hosts the bar shows a color picker for it), drag it to move
/// the stop, click the bar to add a stop, drag a handle well away from the bar to remove it.
/// </summary>
public class GradientStopBar : Control
{
    const double BarHeight = 18;
    const double HandleRadius = 6.5;
    const double HandleCenterY = BarHeight + 4 + HandleRadius;
    const double RemoveDistance = 36;
    const double Inset = HandleRadius + 1;

    readonly List<ColorStop> stops = new List<ColorStop>();
    ColorStop dragged;
    bool isRemoving;

    public GradientStopBar()
    {
        Height = HandleCenterY + HandleRadius + 2;
        Cursor = new Cursor(StandardCursorType.Hand);
    }

    public int MinimumStops { get; set; } = 2;

    /// <summary>Sorted by offset</summary>
    public IReadOnlyList<ColorStop> Stops => stops;

    public ColorStop SelectedStop { get; private set; }

    /// <summary>A different stop is current now (possibly a newly added one).</summary>
    public event Action SelectionChanged;

    /// <summary>A stop moved, appeared, disappeared or changed color.</summary>
    public event Action StopsChanged;

    public void SetStops(IEnumerable<ColorStop> newStops)
    {
        stops.Clear();
        stops.AddRange(newStops.OrderBy(s => s.Offset));
        SelectedStop = stops.FirstOrDefault();
        InvalidateVisual();
        SelectionChanged?.Invoke();
    }

    /// <summary>Changes the color of the selected stop (the picker below the bar calls this).</summary>
    public void SetSelectedColor(Color color)
    {
        if (SelectedStop == null)
        {
            return;
        }

        SelectedStop.Color = color;
        InvalidateVisual();
        StopsChanged?.Invoke();
    }

    double TrackWidth => System.Math.Max(1, Bounds.Width - 2 * Inset);

    double ToX(double offset) => Inset + offset * TrackWidth;

    double ToOffset(double x) => System.Math.Clamp((x - Inset) / TrackWidth, 0, 1);

    public override void Render(DrawingContext context)
    {
        // hit testing goes by what is drawn: make the whole area count, gaps included
        context.DrawRectangle(Brushes.Transparent, pen: null, new Rect(Bounds.Size));

        var bar = new Rect(Inset, 0, TrackWidth, BarHeight);
        context.DrawRectangle(ColorText.CheckerboardBrush, pen: null, bar, 4, 4);

        var gradient = new LinearGradientBrush()
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(1, 0, RelativeUnit.Relative)
        };
        foreach (var stop in stops)
        {
            gradient.GradientStops.Add(new GradientStop(stop.Color, stop.Offset));
        }

        context.DrawRectangle(gradient, new Pen(RibbonTheme.TabLine, thickness: 1), bar, 4, 4);

        // the selected handle last, so that it is on top of a neighbor it overlaps
        foreach (var stop in stops.OrderBy(s => s == SelectedStop))
        {
            var center = new Point(ToX(stop.Offset), HandleCenterY);
            bool isSelected = stop == SelectedStop;
            double opacity = stop == dragged && isRemoving ? 0.3 : 1;
            using (context.PushOpacity(opacity))
            {
                var tick = new Pen(isSelected ? RibbonTheme.ButtonCheckedBorder : RibbonTheme.TabLine, thickness: isSelected ? 2 : 1);
                context.DrawLine(tick, new Point(center.X, BarHeight), new Point(center.X, HandleCenterY - HandleRadius));
                context.DrawEllipse(ColorText.CheckerboardBrush, pen: null, center, HandleRadius, HandleRadius);
                context.DrawEllipse(
                    new SolidColorBrush(stop.Color),
                    new Pen(isSelected ? RibbonTheme.ButtonCheckedBorder : RibbonTheme.TabLine, thickness: isSelected ? 2.5 : 1),
                    center,
                    HandleRadius,
                    HandleRadius);
            }
        }
    }

    protected override void OnPointerPressed(PointerPressedEventArgs e)
    {
        base.OnPointerPressed(e);
        if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            return;
        }

        var position = e.GetPosition(this);
        var hit = stops
            .Where(s => System.Math.Abs(ToX(s.Offset) - position.X) <= HandleRadius + 2)
            .OrderBy(s => System.Math.Abs(ToX(s.Offset) - position.X))
            .FirstOrDefault();

        if (hit == null)
        {
            double offset = ToOffset(position.X);
            hit = new ColorStop(offset, ColorAt(offset));
            stops.Add(hit);
            Sort();
            StopsChanged?.Invoke();
        }

        dragged = hit;
        isRemoving = false;
        e.Pointer.Capture(this);
        e.Handled = true;
        Select(hit);
        InvalidateVisual();
    }

    protected override void OnPointerMoved(PointerEventArgs e)
    {
        base.OnPointerMoved(e);
        if (dragged == null)
        {
            return;
        }

        var position = e.GetPosition(this);
        isRemoving = stops.Count > MinimumStops && System.Math.Abs(position.Y - HandleCenterY) > RemoveDistance;
        dragged.Offset = ToOffset(position.X);
        Sort();
        InvalidateVisual();
        StopsChanged?.Invoke();
    }

    protected override void OnPointerReleased(PointerReleasedEventArgs e)
    {
        base.OnPointerReleased(e);
        if (dragged == null)
        {
            return;
        }

        e.Pointer.Capture(null);
        if (isRemoving)
        {
            stops.Remove(dragged);
            dragged = null;
            isRemoving = false;
            Select(stops.First());
            StopsChanged?.Invoke();
        }

        dragged = null;
        InvalidateVisual();
    }

    protected override void OnPointerCaptureLost(PointerCaptureLostEventArgs e)
    {
        base.OnPointerCaptureLost(e);
        dragged = null;
        isRemoving = false;
        InvalidateVisual();
    }

    void Select(ColorStop stop)
    {
        if (SelectedStop != stop)
        {
            SelectedStop = stop;
            SelectionChanged?.Invoke();
        }
    }

    void Sort()
    {
        stops.Sort((left, right) => left.Offset.CompareTo(right.Offset));
    }

    /// <summary>The color the gradient currently has at the offset: what a new stop there starts as.</summary>
    Color ColorAt(double offset)
    {
        if (stops.Count == 0)
        {
            return Colors.White;
        }

        var before = stops.LastOrDefault(s => s.Offset <= offset) ?? stops.First();
        var after = stops.FirstOrDefault(s => s.Offset >= offset) ?? stops.Last();
        double span = after.Offset - before.Offset;
        double ratio = span <= 0 ? 0 : (offset - before.Offset) / span;
        return Color.FromArgb(
            Mix(before.Color.A, after.Color.A, ratio),
            Mix(before.Color.R, after.Color.R, ratio),
            Mix(before.Color.G, after.Color.G, ratio),
            Mix(before.Color.B, after.Color.B, ratio));
    }

    static byte Mix(byte from, byte to, double ratio) => (byte)System.Math.Round(from + (to - from) * ratio);
}
