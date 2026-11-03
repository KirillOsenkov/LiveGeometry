using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Media.Immutable;
using AvaloniaShapes = Avalonia.Controls.Shapes;

namespace LiveGeometry;

/// <summary>
/// What shows where a figure picked in the Figure List is on the canvas, which may be far from
/// the row: a band along it twice as wide as its <see cref="SelectionHalo"/>, in sharp black
/// and white stripes, faded out from the click on while the halo stays. Over every
/// figure and point (the halo is under them, and a filled shape over the figure hides it),
/// but clear of the figure's own stroke and fill, which show through the middle. A figure
/// whose shape has no geometry (a label, an emoji, a show/hide box) gets a frame around its
/// box. A figure picked on the canvas is under the cursor and gets none.
/// </summary>
public class SelectionFlash : Control
{
    /// <summary>On each side of a stroke or an outline: twice the halo's</summary>
    public static double BandWidth = 6;

    /// <summary>Around a point: twice the halo's</summary>
    public static double PointBandWidth = 8;

    /// <summary>Around the box of a control</summary>
    public static double BoxBandWidth = 6;

    /// <summary>A stripe of either color, in pixels</summary>
    public static double StripeWidth = 4;

    /// <summary>Fading out from full strength over this long, from the click on</summary>
    public static TimeSpan Fade = TimeSpan.FromMilliseconds(700);

    /// <summary>Black and white stripes across the band, repeated along a diagonal</summary>
    static readonly IBrush brush = new ImmutableLinearGradientBrush(
        new[]
        {
            new ImmutableGradientStop(0, Colors.Black),
            new ImmutableGradientStop(0.5, Colors.Black),
            new ImmutableGradientStop(0.5, Colors.White),
            new ImmutableGradientStop(1, Colors.White)
        },
        spreadMethod: GradientSpreadMethod.Repeat,
        startPoint: new RelativePoint(0, 0, RelativeUnit.Absolute),
        endPoint: new RelativePoint(StripeWidth * System.Math.Sqrt(2), StripeWidth * System.Math.Sqrt(2), RelativeUnit.Absolute));

    /// <summary>A shape of the figure and the band drawn around it, kept while the geometry is the same</summary>
    class Flashed
    {
        public Control Shape;
        public bool IsPoint;
        public Geometry Source;
        public double Thickness;
        public Geometry Band;
    }

    readonly List<Flashed> shapes;
    Glide glide;

    SelectionFlash(List<Flashed> shapes)
    {
        this.shapes = shapes;
        IsHitTestVisible = false;
    }

    /// <summary>Flashes the figure, one at a time: a new flash replaces the one before</summary>
    public static void Show(IFigure figure)
    {
        Show([figure]);
    }

    /// <summary>Flashes the figures together (a step of the construction and what it is built on)</summary>
    public static void Show(IEnumerable<IFigure> figures)
    {
        var canvas = figures.FirstOrDefault(figure => figure != null)?.Drawing?.Canvas;
        if (canvas == null)
        {
            return;
        }

        var shapes = new List<Flashed>();
        foreach (var figure in figures)
        {
            Collect(figure, shapes);
        }

        if (shapes.Count == 0)
        {
            return;
        }

        var layer = Layer.Get(canvas);
        foreach (var old in layer.Children.OfType<SelectionFlash>().ToList())
        {
            old.Stop();
        }

        var flash = new SelectionFlash(shapes);
        layer.Children.Add(flash);
        flash.Start();
    }

    /// <summary>Takes away whatever is flashing on the canvas (the stepping left)</summary>
    public static void StopAll(Canvas canvas)
    {
        if (canvas == null)
        {
            return;
        }

        foreach (var flash in Layer.Get(canvas).Children.OfType<SelectionFlash>().ToList())
        {
            flash.Stop();
        }
    }

    /// <summary>The controls that draw the figure: its shape, or the shapes of its parts (a slider, a vector, a regular polygon)</summary>
    static void Collect(IFigure figure, List<Flashed> shapes)
    {
        if (figure is IHasShape owner && owner.Shape != null)
        {
            shapes.Add(new Flashed() { Shape = owner.Shape, IsPoint = figure is IPoint });
        }

        if (figure is CompositeFigure composite)
        {
            foreach (var child in composite.Children)
            {
                if (!(child is BezierPath.BezierPathHandle))
                {
                    Collect(child, shapes);
                }
            }
        }
    }

    void Start()
    {
        glide = new Glide(Fade, t =>
        {
            Opacity = 1 - t;
            if (t >= 1)
            {
                Remove();
            }
            else
            {
                // the shapes move while the view glides to the figure
                InvalidateVisual();
            }
        });
    }

    public void Stop()
    {
        glide?.Finish();
    }

    void Remove()
    {
        if (Parent is Layer layer)
        {
            layer.Children.Remove(this);
        }
    }

    public override void Render(DrawingContext context)
    {
        foreach (var flashed in shapes)
        {
            var shape = flashed.Shape;
            if (!shape.IsEffectivelyVisible)
            {
                continue;
            }

            var transform = shape.TransformToVisual(this);
            if (transform == null)
            {
                continue;
            }

            using (context.PushTransform(transform.Value))
            {
                RenderShape(context, flashed);
            }
        }
    }

    void RenderShape(DrawingContext context, Flashed flashed)
    {
        var shape = flashed.Shape;
        if (shape is AvaloniaShapes.Shape geometryShape && geometryShape.RenderedGeometry != null)
        {
            var band = GetBand(flashed, geometryShape);
            if (band != null)
            {
                context.DrawGeometry(brush, pen: null, band);
            }

            return;
        }

        // no geometry (a character, a label, a box): a frame around its box
        var size = shape.Bounds.Size;
        if (size.Width <= 0 || size.Height <= 0)
        {
            return;
        }

        if (flashed.IsPoint)
        {
            var center = new Point(size.Width / 2, size.Height / 2);
            double radius = System.Math.Max(size.Width, size.Height) / 2 + PointBandWidth / 2;
            context.DrawEllipse(brush: null, new Pen(brush, PointBandWidth), center, radius, radius);
        }
        else
        {
            var box = new Rect(size).Inflate(BoxBandWidth / 2);
            context.DrawRectangle(brush: null, new Pen(brush, BoxBandWidth), box, BoxBandWidth, BoxBandWidth);
        }
    }

    /// <summary>
    /// The band along the shape's outline, clear of its stroke and fill: the outline widened
    /// by the band on each side of the stroke, less the stroke and the fill. Made again only
    /// when the geometry or the thickness changed. A copy of the geometry (as the halo takes):
    /// a geometry is a resource of the compositor, shared with the shape.
    /// </summary>
    static Geometry GetBand(Flashed flashed, AvaloniaShapes.Shape shape)
    {
        var geometry = shape.RenderedGeometry;
        double thickness = shape.StrokeThickness;
        if (ReferenceEquals(geometry, flashed.Source) && thickness == flashed.Thickness)
        {
            return flashed.Band;
        }

        flashed.Source = geometry;
        flashed.Thickness = thickness;
        flashed.Band = null;

        bool stroked = SelectionRegion.IsPainted(shape.Stroke) && thickness > 0;
        bool filled = SelectionRegion.IsPainted(shape.Fill);
        if (!stroked && !filled)
        {
            return null;
        }

        var copy = geometry.Clone();
        double band = flashed.IsPoint ? PointBandWidth : BandWidth;
        double strokeWidth = stroked ? thickness : 0;
        var outer = copy.GetWidenedGeometry(new Pen(Brushes.Black, strokeWidth + 2 * band, lineCap: PenLineCap.Round, lineJoin: PenLineJoin.Round));
        Geometry inner = stroked
            ? copy.GetWidenedGeometry(new Pen(Brushes.Black, thickness, lineCap: shape.StrokeLineCap, lineJoin: shape.StrokeJoin))
            : null;
        if (filled)
        {
            inner = inner == null ? copy : new CombinedGeometry(GeometryCombineMode.Union, copy, inner);
        }

        flashed.Band = new CombinedGeometry(GeometryCombineMode.Exclude, outer, inner);
        return flashed.Band;
    }

    /// <summary>The flashes on a canvas: over every figure and point, whatever its Z</summary>
    public class Layer : Canvas
    {
        public Layer()
        {
            IsHitTestVisible = false;
            ZIndex = ZOrders.Default(ZOrder.SelectionFlash);
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
}
