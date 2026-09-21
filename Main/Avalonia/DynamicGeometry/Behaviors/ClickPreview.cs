using System.Collections.Generic;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using AvaloniaShapes = Avalonia.Controls.Shapes;

namespace DynamicGeometry;

/// <summary>
/// Shows, while the mouse hovers, what a click would make of the point under it
/// (a <see cref="PointPlacement"/>): a faint point where the real one is going to be, a halo on
/// the figures it will depend on and, for a midpoint, a tick on each half of the segment.
/// These are plain visuals on the canvas, not figures: they can't be hit, saved or undone.
/// </summary>
public class ClickPreview
{
    public static double GhostOpacity = 0.4;
    public static double HaloWidth = 6;
    public static double TickLength = 10;

    // shorter than this on screen and the ticks would crowd the ghost point
    public static double MinSegmentLengthForTicks = 40;

    public static IBrush HaloBrush = new SolidColorBrush(Color.FromArgb(0x60, 0x3B, 0x8E, 0xEA));
    public static IBrush TickBrush = new SolidColorBrush(Color.FromRgb(0x2F, 0x7B, 0xD6));

    readonly List<Control> visuals = new List<Control>();
    Canvas canvas;
    AvaloniaShapes.Shape ghost;
    PointPlacement shown;
    Point shownOrigin;
    Point shownUnit;

    public void Show(Drawing drawing, PointPlacement placement, IFigureStyle pointStyle)
    {
        if (drawing == null || drawing.Canvas == null || placement == null || !placement.IsDependent)
        {
            Clear();
            return;
        }

        // the halos and ticks are in pixels, so they are only good for the view they were made in
        var coordinateSystem = drawing.CoordinateSystem;
        var origin = coordinateSystem.ToPhysical(new Point(0, 0));
        var unit = coordinateSystem.ToPhysical(new Point(1, 1));
        bool canReuse = canvas == drawing.Canvas
            && placement.HasSameSources(shown)
            && origin == shownOrigin
            && unit == shownUnit;

        if (!canReuse)
        {
            Clear();
            canvas = drawing.Canvas;
            shownOrigin = origin;
            shownUnit = unit;

            foreach (var source in placement.Sources)
            {
                Add(CreateHalo(source));
            }

            if (placement.Kind == PointPlacementKind.Midpoint)
            {
                AddTicks(coordinateSystem, (Segment)placement.Sources[0]);
            }

            ghost = CreateGhost(pointStyle);
            Add(ghost);
        }

        shown = placement;
        ghost.CenterAt(coordinateSystem.ToPhysical(placement.Coordinates));
    }

    public void Clear()
    {
        if (canvas != null)
        {
            foreach (var visual in visuals)
            {
                canvas.Children.Remove(visual);
            }
        }

        visuals.Clear();
        canvas = null;
        ghost = null;
        shown = null;
    }

    void Add(Control visual)
    {
        if (visual == null)
        {
            return;
        }

        visual.IsHitTestVisible = false;
        visuals.Add(visual);
        canvas.Children.Add(visual);
    }

    static AvaloniaShapes.Shape CreateGhost(IFigureStyle pointStyle)
    {
        var result = Factory.CreatePointShape();
        if (pointStyle != null)
        {
            result.Apply(pointStyle.GetWpfStyle());
        }

        result.Opacity = GhostOpacity;
        result.ZIndex = (int)ZOrder.Points + 1;
        return result;
    }

    /// <summary>
    /// A wide translucent copy of the figure's own shape, right under it.
    /// </summary>
    static Control CreateHalo(IFigure figure)
    {
        AvaloniaShapes.Shape halo = null;

        if (figure is LineBase line)
        {
            halo = new AvaloniaShapes.Line()
            {
                StartPoint = line.Shape.StartPoint,
                EndPoint = line.Shape.EndPoint
            };
            SetHaloStroke(halo, line.Shape);
        }
        else if (figure is EllipseBase ellipse)
        {
            var source = ellipse.Shape;
            halo = new AvaloniaShapes.Ellipse()
            {
                Width = source.Width,
                Height = source.Height,
                RenderTransform = source.RenderTransform,
                RenderTransformOrigin = source.RenderTransformOrigin
            };
            Canvas.SetLeft(halo, Canvas.GetLeft(source));
            Canvas.SetTop(halo, Canvas.GetTop(source));
            SetHaloStroke(halo, source);
        }
        else if (figure is EllipseArcBase arc)
        {
            halo = new AvaloniaShapes.Path()
            {
                Data = arc.Shape.Data
            };
            SetHaloStroke(halo, arc.Shape);
        }

        return halo;
    }

    static void SetHaloStroke(AvaloniaShapes.Shape halo, AvaloniaShapes.Shape source)
    {
        halo.Stroke = HaloBrush;
        halo.StrokeThickness = source.StrokeThickness + HaloWidth;
        halo.StrokeLineCap = PenLineCap.Round;
        halo.ZIndex = source.ZIndex - 1;
    }

    /// <summary>
    /// The school notation for "these two are equal": one tick across each half.
    /// </summary>
    void AddTicks(CoordinateSystem coordinateSystem, Segment segment)
    {
        var start = coordinateSystem.ToPhysical(segment.Coordinates.P1);
        var end = coordinateSystem.ToPhysical(segment.Coordinates.P2);
        var length = start.Distance(end);
        if (length < MinSegmentLengthForTicks)
        {
            return;
        }

        var across = new Point(-(end.Y - start.Y) / length, (end.X - start.X) / length) * (TickLength / 2);
        foreach (var ratio in new[] { 0.25, 0.75 })
        {
            var center = start + (end - start) * ratio;
            Add(new AvaloniaShapes.Line()
            {
                StartPoint = center - across,
                EndPoint = center + across,
                Stroke = TickBrush,
                StrokeThickness = 1.5,
                ZIndex = (int)ZOrder.Points
            });
        }
    }
}
