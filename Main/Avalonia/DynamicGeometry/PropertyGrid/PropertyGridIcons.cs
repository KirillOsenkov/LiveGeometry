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
            default:
                return null;
        }
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
