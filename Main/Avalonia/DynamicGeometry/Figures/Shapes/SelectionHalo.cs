using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Media.Immutable;
using AvaloniaShapes = Avalonia.Controls.Shapes;

namespace DynamicGeometry;

/// <summary>
/// What shows that a figure is selected: a wide striped band under it, along its outline
/// (around a point, a silhouette a little bigger than the point), while the figure itself is
/// drawn as its style says. A selected figure used to be drawn thicker, so the thickness and
/// size being edited in the side panel could not be seen. The hover halo of
/// <see cref="ClickPreview"/> works the same way, plain blue.
/// The halo draws the figure's own geometry, so it needs nothing per kind of figure, and it
/// is drawn again whenever its figure's shape changes or is laid out again.
/// </summary>
public class SelectionHalo : Control
{
    /// <summary>Added to the stroke of a line or an outline</summary>
    public static double BandWidth = 6;

    /// <summary>Added on each side of a point</summary>
    public static double PointBandWidth = 4;

    /// <summary>How far the theme's gradient goes, in pixels, before it turns back</summary>
    public static double StripeWidth = 4;

    static IBrush brush;

    /// <summary>
    /// The theme's <see cref="AppTheme.SelectionHalo"/>, a gradient repeated: its stops and
    /// its direction, from start to end in <see cref="StripeWidth"/> pixels, then back
    /// (reflected, so that two stops make smooth stripes and four make sharp ones)
    /// </summary>
    static IBrush Brush
    {
        get
        {
            if (brush == null)
            {
                brush = MakeRepeating(AppTheme.Current.SelectionHalo);
            }

            return brush;
        }
    }

    static IBrush MakeRepeating(IBrush themeBrush)
    {
        if (themeBrush is not ILinearGradientBrush gradient)
        {
            return themeBrush?.ToImmutable() ?? Brushes.Transparent;
        }

        // the theme's gradient spans the box of whatever it fills: only its direction counts
        double x = gradient.EndPoint.Point.X - gradient.StartPoint.Point.X;
        double y = gradient.EndPoint.Point.Y - gradient.StartPoint.Point.Y;
        double length = System.Math.Sqrt(x * x + y * y);
        if (length == 0)
        {
            x = y = length = 1;
        }

        return new ImmutableLinearGradientBrush(
            gradient.GradientStops.Select(stop => new ImmutableGradientStop(stop.Offset, stop.Color)).ToArray(),
            spreadMethod: GradientSpreadMethod.Reflect,
            startPoint: new RelativePoint(0, 0, RelativeUnit.Absolute),
            endPoint: new RelativePoint(x / length * StripeWidth, y / length * StripeWidth, RelativeUnit.Absolute));
    }

    /// <summary>The halos on screen, all drawn again when the brush changes</summary>
    static readonly HashSet<SelectionHalo> shown = new HashSet<SelectionHalo>();

    static SelectionHalo()
    {
        AppTheme.CurrentChanged += ThemeChanged;
        foreach (var theme in AppTheme.All)
        {
            theme.PropertyChanged += (sender, e) =>
            {
                if (e.PropertyName == nameof(AppTheme.SelectionHalo))
                {
                    ThemeChanged();
                }
            };
        }
    }

    static void ThemeChanged()
    {
        brush = null;
        foreach (var halo in shown)
        {
            halo.InvalidateVisual();
        }
    }

    readonly AvaloniaShapes.Shape source;
    readonly bool isPoint;

    public SelectionHalo(AvaloniaShapes.Shape source, bool isPoint)
    {
        this.source = source;
        this.isPoint = isPoint;
        IsHitTestVisible = false;
    }

    // What redraws the halo is what changes its shape's geometry: a property set (a point's
    // place, a line's ends, a path's geometry, a rotation, and the bounds after a layout),
    // and the two changes in place that the shape itself listens to - the segments of a
    // path's geometry (Geometry.Changed, as Avalonia's Path does) and the points of a
    // polygon (IChangesPointsInPlace). It listened to the shape's LayoutUpdated, which
    // Avalonia raises for every listener after any layout pass anywhere: every halo was
    // drawn again at every frame of a drag, and when a tooltip showed.

    Geometry watchedGeometry;

    protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnAttachedToVisualTree(e);
        source.PropertyChanged += Source_PropertyChanged;
        if (source is IChangesPointsInPlace points)
        {
            points.PointsChangedInPlace += Source_GeometryChanged;
        }

        WatchGeometry((source as AvaloniaShapes.Path)?.Data);
        shown.Add(this);
    }

    protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnDetachedFromVisualTree(e);
        source.PropertyChanged -= Source_PropertyChanged;
        if (source is IChangesPointsInPlace points)
        {
            points.PointsChangedInPlace -= Source_GeometryChanged;
        }

        WatchGeometry(null);
        shown.Remove(this);
    }

    void WatchGeometry(Geometry geometry)
    {
        if (watchedGeometry != null)
        {
            watchedGeometry.Changed -= Source_GeometryChanged;
        }

        watchedGeometry = geometry;
        if (watchedGeometry != null)
        {
            watchedGeometry.Changed += Source_GeometryChanged;
        }
    }

    void Source_PropertyChanged(object sender, AvaloniaPropertyChangedEventArgs e)
    {
        if (e.Property == AvaloniaShapes.Path.DataProperty)
        {
            WatchGeometry((source as AvaloniaShapes.Path)?.Data);
        }

        InvalidateVisual();
    }

    void Source_GeometryChanged(object sender, EventArgs e)
    {
        InvalidateVisual();
    }

    public override void Render(DrawingContext context)
    {
        // A copy: a geometry is a resource of the compositor, and drawn here as well as by
        // its shape it was shared by the two. A Bezier curve whose segments had changed in
        // place while it was selected vanished when it was unselected (the halo let go of
        // the geometry, and the shape drew nothing from then on).
        var geometry = source.RenderedGeometry?.Clone();
        var transform = source.TransformToVisual(this);
        if (geometry == null || transform == null || !source.IsEffectivelyVisible)
        {
            if (isPoint && transform != null && source.IsEffectivelyVisible)
            {
                // a character has no geometry: a disc as big as the character
                var size = source.Bounds.Size;
                var center = new Point(size.Width / 2, size.Height / 2).Transform(transform.Value);
                double radius = System.Math.Max(size.Width, size.Height) / 2 + PointBandWidth;
                context.DrawEllipse(Brush, pen: null, center, radius, radius);
            }

            return;
        }

        using (context.PushTransform(transform.Value))
        {
            if (isPoint)
            {
                var pen = new Pen(Brush, source.StrokeThickness + 2 * PointBandWidth, lineJoin: PenLineJoin.Round);
                context.DrawGeometry(Brush, pen, geometry);
            }
            else
            {
                var pen = new Pen(Brush, source.StrokeThickness + BandWidth, lineCap: PenLineCap.Round, lineJoin: PenLineJoin.Round);
                context.DrawGeometry(brush: null, pen, geometry);
            }
        }
    }

    /// <summary>
    /// All the halos on a canvas, together: they share the opacity, so where two overlap (a
    /// polygon and its sides, a vector's shaft and head) the band is not darker. Over every
    /// line, polygon and label, whatever its Z, under the points: a halo under its figure
    /// would be hidden by a figure brought to front over it.
    /// </summary>
    public class Layer : Canvas
    {
        public static double LayerOpacity = 0.5;

        public Layer()
        {
            Opacity = LayerOpacity;
            IsHitTestVisible = false;
            ZIndex = ZOrders.Default(ZOrder.SelectionHalos);
        }

        public static Layer Get(Canvas canvas)
        {
            var result = canvas.Children.OfType<Layer>().FirstOrDefault();
            if (result == null)
            {
                result = new Layer();
                canvas.Children.Add(result);
            }

            return result;
        }
    }

    public void AddTo(Canvas canvas)
    {
        Layer.Get(canvas).Children.Add(this);
    }

    public void Remove()
    {
        if (Parent is Layer layer)
        {
            layer.Children.Remove(this);
        }
    }
}
