using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// The continuous picker: a saturation/value field for the current hue, a hue strip and an
/// opacity strip. Everything is painted with gradient brushes, there are no bitmaps.
/// </summary>
public class SpectrumPage : ColorPage
{
    const double StripHeight = 14;
    const double MarkerSize = 12;

    HsvColor hsv = HsvColor.FromColor(Colors.Black);

    readonly DragSurface field = new DragSurface() { Height = 132 };
    readonly DragSurface hueStrip = new DragSurface() { Height = StripHeight, Margin = new Thickness(0, 8, 0, 0) };
    readonly DragSurface alphaStrip = new DragSurface() { Height = StripHeight, Margin = new Thickness(0, 8, 0, 0) };

    readonly Border fieldHue = new Border();
    readonly Border alphaGradient = new Border();

    readonly Control fieldMarker = CreateRingMarker();
    readonly Control hueMarker = CreateRingMarker();
    readonly Control alphaMarker = CreateRingMarker();

    public SpectrumPage()
    {
        // saturation: white on the left fading out; value: black at the bottom fading in
        field.Children.Add(fieldHue);
        field.Children.Add(new Border() { Background = Gradient(horizontal: true, Colors.White, Color.FromArgb(0, 255, 255, 255)) });
        field.Children.Add(new Border() { Background = Gradient(horizontal: false, Color.FromArgb(0, 0, 0, 0), Colors.Black) });
        AddMarker(field, fieldMarker);

        var rainbow = new LinearGradientBrush()
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(1, 0, RelativeUnit.Relative)
        };
        for (int i = 0; i <= 6; i++)
        {
            rainbow.GradientStops.Add(new GradientStop(new HsvColor(i * 60, 1, 1).ToColor(), i / 6.0));
        }

        hueStrip.Children.Add(new Border() { Background = rainbow, CornerRadius = new CornerRadius(StripHeight / 2) });
        AddMarker(hueStrip, hueMarker);

        alphaStrip.Children.Add(new Border() { Background = ColorText.CheckerboardBrush, CornerRadius = new CornerRadius(StripHeight / 2) });
        alphaGradient.CornerRadius = new CornerRadius(StripHeight / 2);
        alphaStrip.Children.Add(alphaGradient);
        AddMarker(alphaStrip, alphaMarker);

        field.Dragged += (position, isStillDragging) => Apply(hsv.WithSaturationAndValue(position.X, 1 - position.Y));
        hueStrip.Dragged += (position, isStillDragging) => Apply(hsv.WithHue(System.Math.Min(position.X * 360, 359.999)));
        alphaStrip.Dragged += (position, isStillDragging) => Apply(hsv.WithAlpha(position.X));

        var panel = new StackPanel();
        panel.Children.Add(field);
        panel.Children.Add(hueStrip);
        panel.Children.Add(alphaStrip);
        Child = panel;

        foreach (var surface in new[] { field, hueStrip, alphaStrip })
        {
            surface.SizeChanged += (s, e) => ArrangeMarkers();
        }

        UpdateVisuals();
    }

    public override string Title => "Spectrum";

    void Apply(HsvColor newColor)
    {
        hsv = newColor;
        UpdateVisuals();
        Pick(hsv.ToColor());
    }

    protected override void OnColorSet(Color newColor)
    {
        var converted = HsvColor.FromColor(newColor);

        // grays and black have no hue of their own: keep the one the user was on
        if (converted.Saturation == 0 || converted.Value == 0)
        {
            converted = new HsvColor(
                hsv.Hue,
                converted.Value == 0 ? hsv.Saturation : 0,
                converted.Value,
                converted.Alpha);
        }

        hsv = converted;
        UpdateVisuals();
    }

    void UpdateVisuals()
    {
        fieldHue.Background = new SolidColorBrush(hsv.PureHue);
        var opaque = hsv.WithAlpha(1).ToColor();
        alphaGradient.Background = Gradient(horizontal: true, Color.FromArgb(0, opaque.R, opaque.G, opaque.B), opaque);
        ArrangeMarkers();
    }

    void ArrangeMarkers()
    {
        Place(fieldMarker, field, hsv.Saturation, 1 - hsv.Value);
        Place(hueMarker, hueStrip, hsv.Hue / 360, 0.5);
        Place(alphaMarker, alphaStrip, hsv.Alpha, 0.5);
    }

    static void Place(Control marker, DragSurface surface, double x, double y)
    {
        Canvas.SetLeft(marker, x * surface.Bounds.Width - MarkerSize / 2);
        Canvas.SetTop(marker, y * surface.Bounds.Height - MarkerSize / 2);
    }

    static void AddMarker(DragSurface surface, Control marker)
    {
        var canvas = new Canvas() { IsHitTestVisible = false };
        canvas.Children.Add(marker);
        surface.Children.Add(canvas);
    }

    /// <summary>White ring inside a dark ring: visible on any color.</summary>
    static Control CreateRingMarker()
    {
        var marker = new Panel() { Width = MarkerSize, Height = MarkerSize };
        marker.Children.Add(new Avalonia.Controls.Shapes.Ellipse() { Stroke = new SolidColorBrush(Color.FromArgb(200, 0, 0, 0)), StrokeThickness = 1 });
        marker.Children.Add(new Avalonia.Controls.Shapes.Ellipse() { Stroke = Brushes.White, StrokeThickness = 2, Margin = new Thickness(1) });
        return marker;
    }

    static LinearGradientBrush Gradient(bool horizontal, Color from, Color to)
    {
        var brush = new LinearGradientBrush()
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint = horizontal
                ? new RelativePoint(1, 0, RelativeUnit.Relative)
                : new RelativePoint(0, 1, RelativeUnit.Relative)
        };
        brush.GradientStops.Add(new GradientStop(from, 0));
        brush.GradientStops.Add(new GradientStop(to, 1));
        return brush;
    }
}
