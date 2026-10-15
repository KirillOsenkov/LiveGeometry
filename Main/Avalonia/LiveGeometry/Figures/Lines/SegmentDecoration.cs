using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using AvaloniaShapes = Avalonia.Controls.Shapes;

namespace DynamicGeometry;

/// <summary>
/// The school marks on a segment: ticks across it for "these are equal", chevrons along it
/// for "these are parallel", a wave for a segment that stands for something longer.
/// </summary>
public enum SegmentDecoration
{
    None,
    OneTick,
    TwoTicks,
    ThreeTicks,
    OneArrow,
    TwoArrows,
    ThreeArrows,
    Wave
}

/// <summary>
/// The mark a <see cref="Segment"/> wears at its middle (<see cref="Segment.Decoration"/>): a
/// passive visual, like the <see cref="RightAngleMark"/>, drawn in the segment's own stroke at
/// a fixed size in pixels, so it looks the same at every zoom. Not a figure: nothing to hit,
/// select or list.
/// </summary>
public class SegmentDecorationMark
{
    /// <summary>Across the segment, the same as the midpoint preview's ticks</summary>
    public static double TickLength = ClickPreview.TickLength;

    /// <summary>Between ticks and between chevrons, for a hairline; grows with the stroke</summary>
    public static double Spacing = 4;

    public static double ArrowLength = 7;
    public static double ArrowHalfWidth = 5;
    public static double WaveHalfLength = 12.5;
    public static double WaveHeight = 5.5;

    readonly AvaloniaShapes.Path path = new AvaloniaShapes.Path()
    {
        StrokeLineCap = PenLineCap.Round,
        StrokeJoin = PenLineJoin.Round,
        IsHitTestVisible = false,
        IsVisible = false
    };

    public void OnAddingToCanvas(Canvas canvas)
    {
        if (!canvas.Children.Contains(path))
        {
            canvas.Children.Add(path);
        }
    }

    public void OnRemovingFromCanvas(Canvas canvas)
    {
        canvas.Children.Remove(path);
    }

    public void Hide()
    {
        path.IsVisible = false;
    }

    /// <param name="start">The segment's first end, in pixels</param>
    /// <param name="end">The segment's second end, in pixels; chevrons point this way</param>
    /// <param name="stroke">The segment's stroke, which the mark shares</param>
    /// <param name="thickness">The segment's stroke width, in pixels</param>
    /// <param name="zIndex">The segment's, which the mark shares (brought to front with it)</param>
    public void Show(
        Point start,
        Point end,
        IBrush stroke,
        double thickness,
        int zIndex,
        SegmentDecoration decoration)
    {
        var along = RightAngleMark.Direction(start, end);
        if (decoration == SegmentDecoration.None || along == null)
        {
            Hide();
            return;
        }

        var middle = new Point((start.X + end.X) / 2, (start.Y + end.Y) / 2);
        path.Data = CreateGeometry(decoration, middle, along.Value, thickness);
        path.Stroke = stroke;
        path.StrokeThickness = thickness;
        path.ZIndex = zIndex;
        path.IsVisible = true;
    }

    /// <summary>
    /// The mark's lines, in pixels: at the middle of a segment running along the unit vector,
    /// sized for a stroke of the given width (a thick stroke gets a bigger mark, as an
    /// arrow's head grows with its shaft).
    /// </summary>
    public static PathGeometry CreateGeometry(SegmentDecoration decoration, Point middle, Point along, double thickness)
    {
        var across = new Point(-along.Y, along.X);
        var figures = new PathFigures();
        double grow = System.Math.Max(0, thickness - 1);
        double spacing = Spacing + thickness;
        switch (decoration)
        {
            case SegmentDecoration.OneTick:
            case SegmentDecoration.TwoTicks:
            case SegmentDecoration.ThreeTicks:
                {
                    int count = decoration - SegmentDecoration.OneTick + 1;
                    double half = TickLength / 2 + grow / 2;
                    for (int i = 0; i < count; i++)
                    {
                        var center = middle + along * ((i - (count - 1) / 2.0) * spacing);
                        figures.Add(Line(center - across * half, center + across * half));
                    }

                    break;
                }

            case SegmentDecoration.OneArrow:
            case SegmentDecoration.TwoArrows:
            case SegmentDecoration.ThreeArrows:
                {
                    int count = decoration - SegmentDecoration.OneArrow + 1;
                    double length = ArrowLength + grow;
                    double half = ArrowHalfWidth + grow / 2;
                    for (int i = 0; i < count; i++)
                    {
                        // the chevrons point the segment's way, the group centered on the middle
                        var tip = middle + along * ((i - (count - 1) / 2.0) * spacing + length / 2);
                        var back = tip - along * length;
                        figures.Add(Polyline(back + across * half, tip, back - across * half));
                    }

                    break;
                }

            case SegmentDecoration.Wave:
                {
                    // one S through the middle: a full turn of a sine along the segment
                    double halfLength = WaveHalfLength + grow;
                    double height = WaveHeight + grow / 2;
                    const int steps = 16;
                    var points = new Point[steps + 1];
                    for (int i = 0; i <= steps; i++)
                    {
                        double t = -1 + 2.0 * i / steps;
                        points[i] = middle + along * (t * halfLength) + across * (height * System.Math.Sin(System.Math.PI * t));
                    }

                    figures.Add(Polyline(points));
                    break;
                }
        }

        return new PathGeometry() { Figures = figures };
    }

    static PathFigure Line(Point from, Point to)
    {
        return Polyline(from, to);
    }

    static PathFigure Polyline(params Point[] points)
    {
        // Avalonia closes and fills a figure unless told not to
        var segment = new PolyLineSegment();
        for (int i = 1; i < points.Length; i++)
        {
            segment.Points.Add(points[i]);
        }

        return new PathFigure()
        {
            StartPoint = points[0],
            IsClosed = false,
            IsFilled = false,
            Segments = new PathSegments() { segment }
        };
    }

    /// <summary>The number GeoGebra saves in a segment's decoration element, 1 to 6, as ours; it has no wave</summary>
    public static SegmentDecoration FromGeoGebra(int type)
    {
        return type >= 1 && type <= 6 ? (SegmentDecoration)type : SegmentDecoration.None;
    }
}
