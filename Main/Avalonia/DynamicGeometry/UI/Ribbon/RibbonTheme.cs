using Avalonia;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// The handful of colors and metrics that make up the look of the toolbar, in one place.
/// </summary>
public static class RibbonTheme
{
    public static readonly IBrush Background = new SolidColorBrush(Color.FromRgb(0xF6, 0xF7, 0xF9));
    public static readonly IBrush BottomBorder = new SolidColorBrush(Color.FromRgb(0xD5, 0xD9, 0xE0));
    public static readonly IBrush HeaderRowBackground = new SolidColorBrush(Color.FromRgb(0xE9, 0xEC, 0xF1));
    public static readonly IBrush TabLine = new SolidColorBrush(Color.FromRgb(0xA9, 0xB1, 0xBE));
    public static readonly IBrush Separator = new SolidColorBrush(Color.FromRgb(0xD5, 0xD9, 0xE0));

    public static readonly IBrush ButtonHover = new SolidColorBrush(Color.FromRgb(0xE6, 0xEB, 0xF2));
    public static readonly IBrush ButtonPressed = new SolidColorBrush(Color.FromRgb(0xD3, 0xDC, 0xE8));
    public static readonly IBrush ButtonChecked = new SolidColorBrush(Color.FromRgb(0xD2, 0xE7, 0xFF));
    public static readonly IBrush ButtonCheckedBorder = new SolidColorBrush(Color.FromRgb(0x6F, 0xAE, 0xEC));

    public static readonly IBrush Text = new SolidColorBrush(Color.FromRgb(0x2B, 0x30, 0x38));
    public static readonly IBrush TabHeaderText = new SolidColorBrush(Color.FromRgb(0x2B, 0x30, 0x38));
    public static readonly IBrush TabHeaderTextSelected = new SolidColorBrush(Color.FromRgb(0x00, 0x00, 0x00));

    public static readonly IBrush GroupBackground = new SolidColorBrush(Color.FromRgb(0xFC, 0xFC, 0xFD));
    public static readonly IBrush Destructive = new SolidColorBrush(Color.FromRgb(0xB3, 0x26, 0x1E));

    public static readonly IBrush HintBackground = new SolidColorBrush(Color.FromRgb(0xFF, 0xFD, 0xE8));
    public static readonly IBrush HintBorder = new SolidColorBrush(Color.FromRgb(0xD9, 0xD2, 0x9A));

    public static readonly CornerRadius ButtonCornerRadius = new CornerRadius(5);
    public const double ButtonMinWidth = 52;
}
