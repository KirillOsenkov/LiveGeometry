using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// Background of a tab header - of a ribbon group, or of a segment of a
/// <see cref="SegmentSwitcher"/>. For the selected tab it draws the tab shape: the line
/// running along the bottom of the header row swings up, around the header and back down,
/// so the header reads as one surface with the content below it. The feet of the tab flare
/// outwards by <see cref="Flare"/>; the tab body is inset by that much on both sides, so
/// that everything is drawn within bounds (anything outside gets clipped).
/// </summary>
public class TabOutline : Control
{
    public const double DefaultFlare = 6;

    /// <summary>How far the feet of the tab curve outwards; also the inset of the tab body</summary>
    public double Flare { get; set; } = DefaultFlare;

    /// <summary>Radius of the two top corners</summary>
    public double TopRadius { get; set; } = 7;

    /// <summary>The surface the tab opens into: the selected tab is filled with it.</summary>
    public IBrush Surface { get; set; } = RibbonTheme.Background;

    bool isSelected;
    public bool IsSelected
    {
        get => isSelected;
        set
        {
            isSelected = value;
            InvalidateVisual();
        }
    }

    bool isHovered;
    public bool IsHovered
    {
        get => isHovered;
        set
        {
            isHovered = value;
            InvalidateVisual();
        }
    }

    public override void Render(DrawingContext context)
    {
        double width = Bounds.Width;
        double height = Bounds.Height;
        if (width <= 2 * (Flare + TopRadius) || height <= Flare + TopRadius)
        {
            return;
        }

        if (!isSelected)
        {
            if (isHovered)
            {
                var plate = new Rect(Flare + 2, 2, width - 2 * Flare - 4, height - 5);
                context.DrawRectangle(RibbonTheme.ButtonHover, pen: null, plate, TopRadius - 2, TopRadius - 2);
            }

            return;
        }

        context.DrawGeometry(Surface, pen: null, CreateTab(width, height, isClosed: true));
        context.DrawGeometry(brush: null, new Pen(RibbonTheme.TabLine, thickness: 1), CreateTab(width, height, isClosed: false));
    }

    StreamGeometry CreateTab(double width, double height, bool isClosed)
    {
        // on pixel centers, so that the 1px stroke is crisp and lines up with the row's bottom line
        double left = Flare + 0.5;
        double right = width - Flare - 0.5;
        double top = 0.5;
        double bottom = height - 0.5;
        var flare = new Size(Flare, Flare);
        var corner = new Size(TopRadius, TopRadius);

        var geometry = new StreamGeometry();
        using (var figure = geometry.Open())
        {
            figure.BeginFigure(new Point(0, bottom), isFilled: isClosed);
            figure.ArcTo(new Point(left, bottom - Flare), flare, rotationAngle: 0, isLargeArc: false, SweepDirection.CounterClockwise);
            figure.LineTo(new Point(left, top + TopRadius));
            figure.ArcTo(new Point(left + TopRadius, top), corner, rotationAngle: 0, isLargeArc: false, SweepDirection.Clockwise);
            figure.LineTo(new Point(right - TopRadius, top));
            figure.ArcTo(new Point(right, top + TopRadius), corner, rotationAngle: 0, isLargeArc: false, SweepDirection.Clockwise);
            figure.LineTo(new Point(right, bottom - Flare));
            figure.ArcTo(new Point(width, bottom), flare, rotationAngle: 0, isLargeArc: false, SweepDirection.CounterClockwise);
            if (isClosed)
            {
                // down over the row's bottom line, so that the tab opens into what is below
                figure.LineTo(new Point(width, height));
                figure.LineTo(new Point(0, height));
            }

            figure.EndFigure(isClosed);
        }

        return geometry;
    }
}
