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
    static readonly IBrush DarkCaption = new SolidColorBrush(Color.FromRgb(0x2B, 0x30, 0x38));

    readonly Action action;

    // what the plate was given: a pastel (darkened under the dark theme) or a drawing's paper
    // (shown as it is, a gradient kept whole)
    Color plate;
    bool plateIsPaper;
    IBrush paperGradient;

    IBrush background;
    IBrush hoverBackground;
    IBrush border;
    IBrush hoverBorder;
    bool isPressed;
    bool isOver;

    /// <param name="plate">A near-white tint (see <see cref="Pastels"/>); the border and the
    /// hover states are the same hue, a little deeper</param>
    readonly TextBlock caption;

    public GalleryTile(Control picture, string text, Color plate, Action action)
    {
        this.action = action;
        Picture = picture;
        SetPlate(plate);
        AppTheme.CurrentChanged += Derive;

        CornerRadius = new CornerRadius(10);
        BorderThickness = new Thickness(1.5);
        Cursor = new Cursor(StandardCursorType.Hand);
        ClipToBounds = true;

        caption = new TextBlock()
        {
            Text = text,
            FontSize = 14,
            FontWeight = FontWeight.SemiBold,
            HorizontalAlignment = HorizontalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Margin = new Thickness(10, 2, 10, 10)
        };
        DockPanel.SetDock(caption, Dock.Bottom);

        var layout = new DockPanel();
        layout.Children.Add(caption);
        layout.Children.Add(picture);
        Child = layout;
        Update();

        PointerEntered += (s, e) =>
        {
            isOver = true;
            Update();
        };
        PointerExited += (s, e) =>
        {
            isPressed = false;
            isOver = false;
            Update();
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

    /// <summary>What the tile shows above its caption (a <see cref="DrawingThumbnail"/> for a drawing)</summary>
    public Control Picture { get; }

    bool isCompact;

    /// <summary>A small tile (those of the start row): the caption smaller and closer to the edges</summary>
    public bool IsCompact
    {
        get => isCompact;
        set
        {
            isCompact = value;
            caption.FontSize = value ? 13 : 14;
            caption.Margin = value ? new Thickness(4, 0, 4, 8) : new Thickness(10, 2, 10, 10);
        }
    }

    /// <summary>A pastel plate, as the light theme shows it; the dark theme shows it deepened</summary>
    public void SetPlate(Color pastel)
    {
        plate = pastel;
        plateIsPaper = false;
        paperGradient = null;
        Derive();
    }

    /// <summary>
    /// The plate is a drawing's paper (as the theme on screen resolves it): a solid color as
    /// it is, a gradient as it is (hovering only changes the border, which follows the
    /// gradient's end).
    /// </summary>
    public void SetPlate(IBrush paper)
    {
        if (paper is ISolidColorBrush solid)
        {
            plate = solid.Color;
            plateIsPaper = true;
            paperGradient = null;
            Derive();
        }
        else if (paper is ILinearGradientBrush gradient && gradient.GradientStops.Count > 0)
        {
            plate = gradient.GradientStops[gradient.GradientStops.Count - 1].Color;
            plateIsPaper = true;
            paperGradient = paper;
            Derive();
        }
    }

    /// <summary>The border and the hover states from the plate: deeper than a light plate, lighter than a dark one</summary>
    void Derive()
    {
        var color = plateIsPaper || !IsDark(AppTheme.Current.Page) ? plate : Deepen(plate);
        ToHsv(color, out double hue, out double saturation, out double value);

        // a gray plate keeps a hint of color in its border so that it still reads as a plate
        double tint = System.Math.Max(saturation, 0.03);
        Color hover;
        if (IsDark(color))
        {
            hover = FromHsv(hue, System.Math.Min(tint * 1.4, 1), System.Math.Min(value * 1.18 + 0.02, 1));
            border = new SolidColorBrush(FromHsv(hue, System.Math.Min(tint * 1.6, 1), System.Math.Min(value * 1.5 + 0.04, 1)));
            hoverBorder = new SolidColorBrush(FromHsv(hue, System.Math.Min(tint * 2.5, 1), System.Math.Min(value * 1.9 + 0.06, 1)));
        }
        else
        {
            hover = FromHsv(hue, System.Math.Min(tint * 1.8, 1), value);
            border = new SolidColorBrush(FromHsv(hue, System.Math.Min(tint * 2.2, 1), value * 0.91));
            hoverBorder = new SolidColorBrush(FromHsv(hue, System.Math.Min(tint * 5, 1), value * 0.8));
        }

        background = paperGradient ?? new SolidColorBrush(color);
        hoverBackground = paperGradient ?? new SolidColorBrush(hover);
        Update();
    }

    /// <summary>
    /// A pastel for the dark theme: the same hue, deep instead of pale, with a little more
    /// saturation so that it still reads as a color (a gray stays gray)
    /// </summary>
    static Color Deepen(Color pastel)
    {
        ToHsv(pastel, out double hue, out double saturation, out _);
        double deepSaturation = saturation < 0.05 ? saturation : System.Math.Min(System.Math.Max(saturation * 2.2, 0.16), 1);
        return FromHsv(hue, deepSaturation, 0.24);
    }

    void Update()
    {
        Background = isOver ? hoverBackground : background;
        BorderBrush = isOver ? hoverBorder : border;

        // the caption sits on the bottom of the plate, where a gradient has ended: it goes by
        // the plate (a drawing's own paper is what it is under any theme)
        if (caption != null)
        {
            caption.Foreground = IsDark(BottomColor(background)) ? Brushes.White : DarkCaption;
        }
    }

    static Color BottomColor(IBrush brush)
    {
        if (brush is ILinearGradientBrush gradient && gradient.GradientStops.Count > 0)
        {
            return gradient.GradientStops[gradient.GradientStops.Count - 1].Color;
        }

        return brush is ISolidColorBrush solid ? solid.Color : Colors.White;
    }

    static bool IsDark(Color color)
    {
        return 0.2126 * color.R + 0.7152 * color.G + 0.0722 * color.B < 128;
    }

    public static void ToHsv(Color color, out double hue, out double saturation, out double value)
    {
        double r = color.R / 255.0;
        double g = color.G / 255.0;
        double b = color.B / 255.0;
        double max = Math.Max(r, Math.Max(g, b));
        double min = Math.Min(r, Math.Min(g, b));
        double delta = max - min;
        value = max;
        saturation = max == 0 ? 0 : delta / max;
        if (delta == 0)
        {
            hue = 0;
        }
        else if (max == r)
        {
            hue = 60 * (((g - b) / delta) % 6);
        }
        else if (max == g)
        {
            hue = 60 * ((b - r) / delta + 2);
        }
        else
        {
            hue = 60 * ((r - g) / delta + 4);
        }

        if (hue < 0)
        {
            hue += 360;
        }
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
