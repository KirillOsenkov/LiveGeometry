using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// Background of a ribbon group header. For the selected group it draws the tab: the line
/// running along the bottom of the header row swings up, around the header and back down,
/// so the header reads as one surface with the tools below it. The feet of the tab flare
/// outwards by <see cref="Flare"/>; the tab body is inset by that much on both sides, so
/// that everything is drawn within bounds (anything outside gets clipped).
/// </summary>
public class TabOutline : Control
{
    public const double Flare = 6;
    public const double CornerRadius = 7;

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
        if (width <= 2 * (Flare + CornerRadius) || height <= Flare + CornerRadius)
        {
            return;
        }

        if (!isSelected)
        {
            if (isHovered)
            {
                var plate = new Rect(Flare + 2, 3, width - 2 * Flare - 4, height - 7);
                context.DrawRectangle(RibbonTheme.ButtonHover, pen: null, plate, CornerRadius - 2, CornerRadius - 2);
            }

            return;
        }

        context.DrawGeometry(RibbonTheme.Background, pen: null, CreateTab(width, height, isClosed: true));
        context.DrawGeometry(brush: null, new Pen(RibbonTheme.TabLine, thickness: 1), CreateTab(width, height, isClosed: false));
    }

    static StreamGeometry CreateTab(double width, double height, bool isClosed)
    {
        // on pixel centers, so that the 1px stroke is crisp and lines up with the row's bottom line
        double left = Flare + 0.5;
        double right = width - Flare - 0.5;
        double top = 0.5;
        double bottom = height - 0.5;
        var flare = new Size(Flare, Flare);
        var corner = new Size(CornerRadius, CornerRadius);

        var geometry = new StreamGeometry();
        using (var figure = geometry.Open())
        {
            figure.BeginFigure(new Point(0, bottom), isFilled: isClosed);
            figure.ArcTo(new Point(left, bottom - Flare), flare, rotationAngle: 0, isLargeArc: false, SweepDirection.CounterClockwise);
            figure.LineTo(new Point(left, top + CornerRadius));
            figure.ArcTo(new Point(left + CornerRadius, top), corner, rotationAngle: 0, isLargeArc: false, SweepDirection.Clockwise);
            figure.LineTo(new Point(right - CornerRadius, top));
            figure.ArcTo(new Point(right, top + CornerRadius), corner, rotationAngle: 0, isLargeArc: false, SweepDirection.Clockwise);
            figure.LineTo(new Point(right, bottom - Flare));
            figure.ArcTo(new Point(width, bottom), flare, rotationAngle: 0, isLargeArc: false, SweepDirection.CounterClockwise);
            if (isClosed)
            {
                // down over the row's bottom line, so that the tab opens into the tools below
                figure.LineTo(new Point(width, height));
                figure.LineTo(new Point(0, height));
            }

            figure.EndFigure(isClosed);
        }

        return geometry;
    }
}
