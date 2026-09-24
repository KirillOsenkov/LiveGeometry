using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using AvaloniaShapes = Avalonia.Controls.Shapes;

namespace DynamicGeometry;

/// <summary>
/// Shows, while the mouse hovers, what a click would do. For a point
/// (a <see cref="PointPlacement"/>): a faint point where the real one is going to be, a halo on
/// the figures it will depend on and, for a midpoint, a tick on each half of the segment.
/// For a tool that needs a figure: a halo on the figure the click would pick.
/// These are plain visuals on the canvas, not figures: they can't be hit, saved or undone.
/// </summary>
public class ClickPreview
{
    public static double GhostOpacity = 0.4;
    public static double HaloWidth = 6;
    public static double TickLength = 10;

    // shorter than this on screen and the ticks would crowd the ghost point
    public static double MinSegmentLengthForTicks = 40;

    public static IBrush HaloBrush = new SolidColorBrush(Color.FromArgb(0x20, 0x3B, 0x8E, 0xEA));
    // a point is small, so its halo is a little wider and stronger than a line's
    public static double PointHaloWidth = 5;
    public static IBrush PointHaloBrush = new SolidColorBrush(Color.FromArgb(0x48, 0x3B, 0x8E, 0xEA));
    public static IBrush TickBrush = new SolidColorBrush(Color.FromRgb(0x2F, 0x7B, 0xD6));

    readonly List<Control> visuals = new List<Control>();
    Canvas canvas;
    AvaloniaShapes.Shape ghost;
    IReadOnlyList<IFigure> shownSources;
    PointPlacementKind? shownKind;
    Point shownOrigin;
    Point shownUnit;

    /// <param name="placement">The point a click would create, or null</param>
    /// <param name="pickedFigure">The figure a click would pick (a tool that needs a line
    /// over a line), or null. Only looked at when there is no point to show.</param>
    public void Show(
        Drawing drawing,
        PointPlacement placement,
        IFigure pickedFigure,
        IFigureStyle pointStyle)
    {
        if (placement != null && !placement.IsDependent)
        {
            placement = null;
        }

        if (drawing == null || drawing.Canvas == null || (placement == null && pickedFigure == null))
        {
            Clear();
            return;
        }

        IReadOnlyList<IFigure> sources = placement != null ? placement.Sources : new[] { pickedFigure };

        // the halos and ticks are in pixels, so they are only good for the view they were made in
        var coordinateSystem = drawing.CoordinateSystem;
        var origin = coordinateSystem.ToPhysical(new Point(0, 0));
        var unit = coordinateSystem.ToPhysical(new Point(1, 1));
        var kind = placement != null ? placement.Kind : (PointPlacementKind?)null;
        bool canReuse = canvas == drawing.Canvas
            && kind == shownKind
            && shownSources != null
            && sources.SequenceEqual(shownSources)
            && origin == shownOrigin
            && unit == shownUnit;

        if (!canReuse)
        {
            Clear();
            canvas = drawing.Canvas;
            shownOrigin = origin;
            shownUnit = unit;
            shownKind = kind;
            shownSources = sources;

            foreach (var source in sources)
            {
                // a hidden figure's shape is never updated, so a halo made from it would be stale
                if (source.Visible)
                {
                    Add(CreateHalo(source));
                }
            }

            if (kind == PointPlacementKind.Midpoint)
            {
                AddTicks(coordinateSystem, (Segment)sources[0]);
            }

            if (placement != null)
            {
                ghost = CreateGhost(pointStyle);
                Add(ghost);
            }
        }

        if (ghost != null)
        {
            ghost.CenterAt(coordinateSystem.ToPhysical(placement.Coordinates));
        }
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
        shownSources = null;
        shownKind = null;
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
        else if (figure is PointBase point)
        {
            // a disc behind the point; an outline this thin and pale would not be seen
            var source = point.Shape;
            var diameter = source.Width + 2 * PointHaloWidth;
            var center = point.Drawing.CoordinateSystem.ToPhysical(point.Coordinates);
            halo = new AvaloniaShapes.Ellipse()
            {
                Width = diameter,
                Height = diameter,
                Fill = PointHaloBrush,
                ZIndex = source.ZIndex - 1
            };
            halo.CenterAt(center);
        }
        else if (figure is PolygonBase polygon)
        {
            halo = new AvaloniaShapes.Polygon()
            {
                Points = polygon.Shape.Points.ToList()
            };
            SetHaloStroke(halo, polygon.Shape);
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
