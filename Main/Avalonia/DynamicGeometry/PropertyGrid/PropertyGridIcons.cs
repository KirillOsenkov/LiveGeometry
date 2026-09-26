using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Layout;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// The small icons on property grid buttons, drawn on a 14 x 14 grid like the trash can
/// of a destructive button.
/// </summary>
public static class PropertyGridIcons
{
    public const double Size = 14;

    static readonly IBrush outline = new SolidColorBrush(Color.FromRgb(0x3A, 0x42, 0x50));
    static readonly IBrush wood = new SolidColorBrush(Color.FromRgb(0xF2, 0xB6, 0x32));
    static readonly IBrush woodOutline = new SolidColorBrush(Color.FromRgb(0xA8, 0x7B, 0x05));
    static readonly IBrush sharpened = new SolidColorBrush(Color.FromRgb(0xF7, 0xE6, 0xC4));
    static readonly IBrush eraser = new SolidColorBrush(Color.FromRgb(0xFF, 0xB3, 0xC6));
    static readonly IBrush green = new SolidColorBrush(Color.FromRgb(0x2E, 0x9E, 0x4F));

    public static Control Create(PropertyGridIcon icon)
    {
        switch (icon)
        {
            case PropertyGridIcon.Pencil:
                return Pencil();
            case PropertyGridIcon.Plus:
                return Plus();
            case PropertyGridIcon.Lock:
                return Padlock(open: false);
            case PropertyGridIcon.Unlock:
                return Padlock(open: true);
            case PropertyGridIcon.Check:
                return Icon(Shape("M2.8,7.6 L5.8,10.6 L11.4,3.8", null, green, thickness: 1.8));
            case PropertyGridIcon.Cross:
                return Icon(Shape("M3.5,3.5 L10.5,10.5 M10.5,3.5 L3.5,10.5", null, outline, thickness: 1.6));
            case PropertyGridIcon.Segment:
                return Icon(Shape("M3,11 L11,3", null, outline, thickness: 1.2), Dot(3, 11), Dot(11, 3));
            case PropertyGridIcon.Ray:
                return Icon(Shape("M3,11 L11.5,2.5", null, outline, thickness: 1.2), Arrowhead(11.5, 2.5), Dot(3, 11));
            case PropertyGridIcon.Line:
                return Icon(Shape("M2.5,11.5 L11.5,2.5", null, outline, thickness: 1.2), Arrowhead(11.5, 2.5), Arrowhead(2.5, 11.5, back: true));
            case PropertyGridIcon.Reverse:
                return Icon(Shape("M2.5,4.5 H11.5 M9,2 L11.5,4.5 L9,7 M11.5,9.5 H2.5 M5,7 L2.5,9.5 L5,12", null, outline, thickness: 1.4));
            case PropertyGridIcon.Angle:
                return Icon(
                    Shape("M2.5,11.5 H12.5 M2.5,11.5 L10,2.5", null, outline, thickness: 1.2),
                    Shape("M8,11.5 A5.5,5.5 0 0 0 5.9,7.2", null, green, thickness: 1.4));
            case PropertyGridIcon.Arc:
                return Icon(Shape("M2,11 A6,6 0 0 1 12,11", null, outline, thickness: 1.4), Dot(2, 11), Dot(12, 11));
            case PropertyGridIcon.CircleSegment:
                return Icon(Shape("M2,9.5 A6.3,6.3 0 0 1 12,9.5 Z", sky, outline, thickness: 1.2));
            case PropertyGridIcon.Sector:
                return Icon(Shape("M7,12 L3,3.8 A7.5,7.5 0 0 1 11,3.8 Z", sky, outline, thickness: 1.2));
            case PropertyGridIcon.Polyline:
                return Icon(Shape("M2,11 L5,4 L8.5,10 L12,3", null, outline, thickness: 1.4));
            default:
                return null;
        }
    }

    static readonly IBrush point = new SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0x64));
    static readonly IBrush sky = new SolidColorBrush(Color.FromRgb(0xC8, 0xE0, 0xFF));

    /// <summary>A yellow point, like a free point on the canvas</summary>
    static Path Dot(double x, double y)
    {
        const double radius = 1.9;
        var data = string.Format(
            System.Globalization.CultureInfo.InvariantCulture,
            "M{0},{1} m-{2},0 a{2},{2} 0 1 0 {3},0 a{2},{2} 0 1 0 -{3},0",
            x,
            y,
            radius,
            2 * radius);
        return Shape(data, point, outline, thickness: 0.8);
    }

    /// <summary>The two barbs of an arrow pointing up-right at (x, y), or down-left when it points back</summary>
    static Path Arrowhead(double x, double y, bool back = false)
    {
        double sign = back ? -1 : 1;
        var data = string.Format(
            System.Globalization.CultureInfo.InvariantCulture,
            "M{0},{1} L{2},{3} M{0},{1} L{4},{5}",
            x,
            y,
            x - 4 * sign,
            y + 0.8 * sign,
            x - 0.8 * sign,
            y + 4 * sign);
        return Shape(data, null, outline, thickness: 1.2);
    }

    /// <summary>
    /// A gold padlock: the shackle closed into the body, or lifted and swung open to the
    /// right with its left leg in the air.
    /// </summary>
    static Control Padlock(bool open)
    {
        // open: the shackle is lifted so that its left leg ends well above the body
        string shackle = open
            ? "M9.3,6 V2.4 A2.4,2.4 0 0 0 4.5,2.4 V3.4"
            : "M4.5,6 V3.6 A2.4,2.4 0 0 1 9.3,3.6 V6";
        return Icon(
            Shape(shackle, null, outline, thickness: 1.5),
            Shape("M2.5,6 H11.5 V12.5 H2.5 Z", wood, woodOutline),
            Shape("M7,8.3 V10.3", null, outline, thickness: 1.5));
    }

    /// <summary>A yellow pencil lying at 45°, point down-left, eraser up-right.</summary>
    static Control Pencil()
    {
        return Icon(
            Shape("M3.03,8.43 L9.93,1.53 L12.47,4.07 L5.57,10.97 Z", wood, woodOutline),
            Shape("M8.8,2.66 L9.93,1.53 L12.47,4.07 L11.34,5.2 Z", eraser, woodOutline),
            Shape("M3.03,8.43 L5.57,10.97 L2.2,11.8 Z", sharpened, woodOutline),
            Shape("M2.2,11.8 L2.53,10.45 L3.55,11.47 Z", outline, null));
    }

    static Control Plus()
    {
        return Icon(Shape("M7,2.5 V11.5 M2.5,7 H11.5", null, green, thickness: 1.8));
    }

    static Control Icon(params Path[] shapes)
    {
        var canvas = new Canvas()
        {
            Width = Size,
            Height = Size,
            VerticalAlignment = VerticalAlignment.Center
        };
        canvas.Children.AddRange(shapes);
        return canvas;
    }

    static Path Shape(string data, IBrush fill, IBrush stroke, double thickness = 1)
    {
        return new Path()
        {
            Data = Geometry.Parse(data),
            Fill = fill,
            Stroke = stroke,
            StrokeThickness = stroke != null ? thickness : 0,
            StrokeJoin = PenLineJoin.Round,
            StrokeLineCap = PenLineCap.Round
        };
    }
}
