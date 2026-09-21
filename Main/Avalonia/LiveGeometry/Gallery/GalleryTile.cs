using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using DynamicGeometry;
using Math = System.Math;

namespace LiveGeometry;

/// <summary>
/// A big button of the gallery: a near-white pastel plate with a picture and a caption under it.
/// </summary>
public class GalleryTile : Border
{
    readonly Action action;
    readonly IBrush background;
    readonly IBrush hoverBackground;
    readonly IBrush border;
    readonly IBrush hoverBorder;
    bool isPressed;

    /// <param name="hue">0-360; the plate is a very pale tint of it</param>
    public GalleryTile(Control picture, string caption, double hue, Action action)
    {
        this.action = action;
        background = new SolidColorBrush(FromHsv(hue, saturation: 0.07, value: 0.995));
        hoverBackground = new SolidColorBrush(FromHsv(hue, saturation: 0.13, value: 0.99));
        border = new SolidColorBrush(FromHsv(hue, saturation: 0.16, value: 0.90));
        hoverBorder = new SolidColorBrush(FromHsv(hue, saturation: 0.45, value: 0.78));

        CornerRadius = new CornerRadius(10);
        BorderThickness = new Thickness(1.5);
        Cursor = new Cursor(StandardCursorType.Hand);
        ClipToBounds = true;

        var text = new TextBlock()
        {
            Text = caption,
            FontSize = 14,
            FontWeight = FontWeight.SemiBold,
            Foreground = RibbonTheme.Text,
            HorizontalAlignment = HorizontalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Margin = new Thickness(10, 2, 10, 10)
        };
        DockPanel.SetDock(text, Dock.Bottom);

        var layout = new DockPanel();
        layout.Children.Add(text);
        layout.Children.Add(picture);
        Child = layout;
        Update(isOver: false);

        PointerEntered += (s, e) => Update(isOver: true);
        PointerExited += (s, e) =>
        {
            isPressed = false;
            Update(isOver: false);
        };
        PointerPressed += (s, e) =>
        {
            if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            {
                isPressed = true;
            }
        };
        PointerReleased += (s, e) =>
        {
            bool wasPressed = isPressed;
            isPressed = false;
            if (wasPressed && e.InitialPressMouseButton == MouseButton.Left)
            {
                this.action();
            }
        };
    }

    void Update(bool isOver)
    {
        Background = isOver ? hoverBackground : background;
        BorderBrush = isOver ? hoverBorder : border;
    }

    public static Color FromHsv(double hue, double saturation, double value)
    {
        hue = ((hue % 360) + 360) % 360;
        double chroma = value * saturation;
        double x = chroma * (1 - Math.Abs(hue / 60 % 2 - 1));
        double m = value - chroma;
        (double r, double g, double b) = ((int)(hue / 60)) switch
        {
            0 => (chroma, x, 0.0),
            1 => (x, chroma, 0.0),
            2 => (0.0, chroma, x),
            3 => (0.0, x, chroma),
            4 => (x, 0.0, chroma),
            _ => (chroma, 0.0, x)
        };

        return Color.FromRgb(
            (byte)Math.Round((r + m) * 255),
            (byte)Math.Round((g + m) * 255),
            (byte)Math.Round((b + m) * 255));
    }
}
