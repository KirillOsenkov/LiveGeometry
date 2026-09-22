using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using DynamicGeometry;

namespace LiveGeometry;

/// <summary>
/// The strip at the very top: the few things that are about the document and not about geometry
/// (new, open, save, undo, redo). Icons only, the names and shortcuts are in the tooltips.
/// Three parts, laid out by hand: the buttons at the left, something small at the right (the
/// build stamp, dropped when there is no room for it) and the centered group in between. When
/// the window is too narrow for the group beside the buttons, it wraps to a row of its own.
/// </summary>
public class MainToolbar : Panel
{
    readonly StackPanel buttons = new StackPanel()
    {
        Orientation = Orientation.Horizontal,
        Spacing = 2,
        Margin = new Thickness(8, 4, 6, 0)
    };

    // where buttons are added: the strip itself or the group that is open
    Panel current;

    Control rightControl;
    double rightWidth; // as last measured; kept while it is hidden

    // the group between the buttons and the right control
    MainToolbarGroup centered;

    // decided by the measure pass for the arrange pass
    bool wrapped;
    double firstRowHeight;

    // between a wrapped group and the ribbon under it
    const double WrappedRowGap = 4;

    public MainToolbar()
    {
        current = buttons;

        // The same gray as the ribbon's header row right under it, and no line between them:
        // the two read as one band, so the window has two grays and not three.
        Background = RibbonTheme.HeaderRowBackground;
        Children.Add(buttons);
    }

    /// <summary>Something small at the far right (the build stamp)</summary>
    public void AddAtRight(Control control)
    {
        rightControl = control;
        Children.Add(control);
    }

    protected override Size MeasureOverride(Size availableSize)
    {
        var unbounded = new Size(double.PositiveInfinity, double.PositiveInfinity);
        buttons.Measure(unbounded);
        double width = buttons.DesiredSize.Width;
        double height = buttons.DesiredSize.Height;

        // the group beside the buttons if all of it fits there (the stamp may have to go),
        // else on a row of its own, where the trailing text still gets cut short if it must
        bool hasCentered = centered != null && centered.IsVisible;
        double centeredWidth = 0;
        wrapped = false;
        if (hasCentered)
        {
            centered.Measure(unbounded);
            centeredWidth = centered.DesiredSize.Width;
            wrapped = width + centeredWidth > availableSize.Width;
        }

        if (!wrapped)
        {
            width += centeredWidth;
            height = System.Math.Max(height, centered?.DesiredSize.Height ?? 0);
        }

        double rightRoom = 0;
        if (rightControl != null)
        {
            // a hidden control measures as nothing, so its width is remembered from when it
            // was visible; it never changes
            if (rightControl.IsVisible)
            {
                rightControl.Measure(unbounded);
                rightWidth = rightControl.DesiredSize.Width;
            }

            bool fits = width + rightWidth <= availableSize.Width;
            rightControl.IsVisible = fits;
            if (fits)
            {
                rightRoom = rightWidth;
                width += rightWidth;
                height = System.Math.Max(height, rightControl.DesiredSize.Height);
            }
        }

        firstRowHeight = height;

        // once more with the room it will really get, so that the trailing text lays itself
        // out to fit (measured unbounded it would just run off the edge)
        if (hasCentered)
        {
            double room = wrapped
                ? availableSize.Width
                : System.Math.Max(0, availableSize.Width - buttons.DesiredSize.Width - rightRoom);
            centered.IsLeftAligned = wrapped;
            centered.Measure(new Size(room, double.PositiveInfinity));
            if (wrapped)
            {
                width = System.Math.Max(width, centered.DesiredSize.Width);
                height += centered.DesiredSize.Height + WrappedRowGap;
            }
        }

        return new Size(width, height);
    }

    protected override Size ArrangeOverride(Size finalSize)
    {
        double left = buttons.DesiredSize.Width;
        buttons.Arrange(new Rect(0, 0, left, firstRowHeight));

        double right = finalSize.Width;
        if (rightControl != null && rightControl.IsVisible)
        {
            right -= rightWidth;
            rightControl.Arrange(new Rect(right, 0, rightWidth, firstRowHeight));
        }

        if (centered != null)
        {
            if (wrapped)
            {
                centered.Arrange(new Rect(0, firstRowHeight, finalSize.Width, centered.DesiredSize.Height));
            }
            else
            {
                centered.Arrange(new Rect(left, 0, System.Math.Max(0, right - left), firstRowHeight));
            }
        }

        return finalSize;
    }

    public MainToolbarButton AddButton(
        Control icon,
        string name,
        string shortcut,
        Action action,
        double iconSize = MainToolbarButton.IconSize,
        double inset = MainToolbarButton.DefaultInset)
    {
        var button = new MainToolbarButton(icon, action, iconSize, inset);
        ToolTip.SetTip(button, shortcut != null ? name + " (" + shortcut + ")" : name);
        current.Children.Add(button);
        return button;
    }

    public void AddSeparator()
    {
        current.Children.Add(new Border()
        {
            Width = 1,
            Margin = new Thickness(5, 5, 5, 5),
            Background = RibbonTheme.TabLine
        });
    }

    public TextBlock AddText(FontWeight weight, double minWidth = 0, double fontSize = 13)
    {
        var text = new TextBlock()
        {
            FontSize = fontSize,
            FontWeight = weight,
            Foreground = RibbonTheme.Text,
            MinWidth = minWidth,
            TextAlignment = TextAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
        current.Children.Add(text);
        return text;
    }

    /// <summary>
    /// What is added from now on goes into a group of its own, to be shown and hidden together,
    /// in the room between the buttons and whatever is at the right (see
    /// <see cref="MainToolbarGroup"/> for where exactly).
    /// </summary>
    public Panel BeginCenteredGroup()
    {
        centered = new MainToolbarGroup(buttons.Spacing);
        Children.Add(centered);
        current = centered.Middle;
        return centered;
    }

    /// <summary>Text to the right of the centered group, cut short when there is no room</summary>
    public TextBlock AddTrailingText(FontWeight weight, double fontSize)
    {
        var text = new TextBlock()
        {
            FontSize = fontSize,
            FontWeight = weight,
            Foreground = RibbonTheme.Text,
            TextTrimming = TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(6, 0, 6, 0)
        };
        centered.Trailing.Children.Add(text);
        return text;
    }

    public void EndGroup()
    {
        current = buttons;
    }

    /// <summary>A button that follows a command: runs it, and is disabled when it is</summary>
    public MainToolbarButton AddButton(Control icon, string shortcut, Command command)
    {
        var button = AddButton(icon, command.Name, shortcut, command.Execute);
        command.AddObserver(button);
        button.EnabledChanged(command.Enabled);
        return button;
    }
}

/// <summary>
/// The toolbar's middle group: a few buttons (<see cref="Middle"/>) with text trailing them
/// (<see cref="Trailing"/>). The buttons are centered in the width of the group as long as the
/// text still fits whole to their right; then they move left just as far as the text needs,
/// and only when even that is not enough is the text cut short. So the buttons stay put from
/// one text to the next until room really runs out. <see cref="IsLeftAligned"/> packs
/// everything to the left instead (for a row of its own).
/// </summary>
public class MainToolbarGroup : Panel
{
    public StackPanel Middle { get; }
    public Panel Trailing { get; }

    bool isLeftAligned;
    double trailingNaturalWidth;

    public MainToolbarGroup(double spacing)
    {
        Middle = new StackPanel()
        {
            Orientation = Orientation.Horizontal,
            Spacing = spacing,
            Margin = new Thickness(6, 2, 6, 0)
        };
        // not a StackPanel: that would hand the text unlimited width and it could never trim
        Trailing = new Panel()
        {
            Margin = new Thickness(0, 2, 6, 0)
        };
        Children.Add(Middle);
        Children.Add(Trailing);
    }

    public bool IsLeftAligned
    {
        get => isLeftAligned;
        set
        {
            if (isLeftAligned != value)
            {
                isLeftAligned = value;
                InvalidateMeasure();
            }
        }
    }

    protected override Size MeasureOverride(Size availableSize)
    {
        var unbounded = new Size(double.PositiveInfinity, double.PositiveInfinity);
        Middle.Measure(unbounded);
        Trailing.Measure(unbounded);
        trailingNaturalWidth = Trailing.DesiredSize.Width;

        // the text is measured again with what it will get, so that it trims itself
        if (!double.IsPositiveInfinity(availableSize.Width))
        {
            Place(availableSize.Width, out _, out double trailingWidth);
            Trailing.Measure(new Size(trailingWidth, double.PositiveInfinity));
        }

        return new Size(
            Middle.DesiredSize.Width + trailingNaturalWidth,
            System.Math.Max(Middle.DesiredSize.Height, Trailing.DesiredSize.Height));
    }

    protected override Size ArrangeOverride(Size finalSize)
    {
        Place(finalSize.Width, out double left, out double trailingWidth);
        double middleWidth = Middle.DesiredSize.Width;
        Middle.Arrange(new Rect(left, 0, middleWidth, finalSize.Height));
        Trailing.Arrange(new Rect(left + middleWidth, 0, trailingWidth, finalSize.Height));
        return finalSize;
    }

    /// <summary>Where the middle starts, and how much the text after it gets</summary>
    void Place(double width, out double left, out double trailingWidth)
    {
        double middleWidth = Middle.DesiredSize.Width;
        double centeredLeft = (width - middleWidth) / 2;
        double leftForWholeText = width - middleWidth - trailingNaturalWidth;
        left = isLeftAligned ? 0 : System.Math.Max(0, System.Math.Min(centeredLeft, leftForWholeText));
        trailingWidth = System.Math.Max(0, width - left - middleWidth);
    }
}

public class MainToolbarButton : Border, ICommandObserver
{
    /// <summary>How big the icons are drawn (their grid is <see cref="MainToolbarIcons.Size"/>)</summary>
    public const double IconSize = 24;

    /// <summary>The plate around the icon (a little more at the sides)</summary>
    public const double DefaultInset = 5;

    readonly Action action;
    bool isPressed;

    public MainToolbarButton(
        Control icon,
        Action action,
        double iconSize = IconSize,
        double inset = DefaultInset)
    {
        this.action = action;
        Width = iconSize + 2 * inset + 4;
        Height = iconSize + 2 * inset;
        CornerRadius = RibbonTheme.ButtonCornerRadius;
        Background = Brushes.Transparent;
        VerticalAlignment = VerticalAlignment.Center;
        Child = new Viewbox()
        {
            Width = iconSize,
            Height = iconSize,
            Child = icon
        };

        PointerEntered += (s, e) => UpdateBackground(isOver: true);
        PointerExited += (s, e) =>
        {
            isPressed = false;
            UpdateBackground(isOver: false);
        };
        PointerPressed += (s, e) =>
        {
            if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            {
                isPressed = true;
                UpdateBackground(isOver: true);
            }
        };
        PointerReleased += (s, e) =>
        {
            bool wasPressed = isPressed;
            isPressed = false;
            UpdateBackground(isOver: IsPointerOver);
            if (wasPressed && e.InitialPressMouseButton == MouseButton.Left)
            {
                this.action();
            }
        };
    }

    void UpdateBackground(bool isOver)
    {
        // on the header row, which is darker than the tools strip: the hover plate is lighter
        Background = isPressed ? RibbonTheme.ButtonPressed : (isOver ? RibbonTheme.GroupBackground : Brushes.Transparent);
    }

    public void EnabledChanged(bool newEnabledState)
    {
        IsEnabled = newEnabledState;
        Opacity = newEnabledState ? 1 : 0.35;
        if (!newEnabledState)
        {
            isPressed = false;
            UpdateBackground(isOver: false);
        }
    }

    public void CommandRemoved()
    {
    }

    public void IconChanged(Control icon)
    {
    }
}

/// <summary>
/// Drawn, not loaded: no image files to ship to the browser, crisp at any scaling.
/// All on a 20 x 20 grid.
/// </summary>
public static class MainToolbarIcons
{
    public const double Size = 20;

    static readonly IBrush outline = new SolidColorBrush(Color.FromRgb(0x3A, 0x42, 0x50));
    static readonly IBrush paper = Brushes.White;
    static readonly IBrush green = new SolidColorBrush(Color.FromRgb(0x2E, 0x9E, 0x4F));
    static readonly IBrush folderBack = new SolidColorBrush(Color.FromRgb(0xF2, 0xB6, 0x32));
    static readonly IBrush folderFront = new SolidColorBrush(Color.FromRgb(0xFA, 0xD5, 0x65));
    static readonly IBrush folderOutline = new SolidColorBrush(Color.FromRgb(0xA8, 0x7B, 0x05));
    static readonly IBrush diskBody = new SolidColorBrush(Color.FromRgb(0x4C, 0x8B, 0xF5));
    static readonly IBrush diskOutline = new SolidColorBrush(Color.FromRgb(0x2A, 0x5D, 0xB0));
    static readonly IBrush diskLabel = new SolidColorBrush(Color.FromRgb(0xE8, 0xEE, 0xF9));
    static readonly IBrush arrow = new SolidColorBrush(Color.FromRgb(0x2F, 0x7B, 0xD6));
    static readonly IBrush tileBlue = new SolidColorBrush(Color.FromRgb(0xBF, 0xDC, 0xFF));
    static readonly IBrush tileYellow = new SolidColorBrush(Color.FromRgb(0xFF, 0xE7, 0xA3));
    static readonly IBrush tileGreen = new SolidColorBrush(Color.FromRgb(0xC4, 0xEB, 0xC8));
    static readonly IBrush tilePink = new SolidColorBrush(Color.FromRgb(0xFF, 0xCF, 0xDD));

    public static Control New()
    {
        return Icon(
            Shape("M4.5,2.5 H11.5 L15.5,6.5 V17.5 H4.5 Z", paper, outline),
            Shape("M11.5,2.5 V6.5 H15.5", null, outline),
            Shape("M14.5,10.5 A4,4 0 1 1 14.49,10.5 Z", green, null),
            Shape("M14.5,12.3 V16.7 M12.3,14.5 H16.7", null, Brushes.White, thickness: 1.6));
    }

    public static Control Open()
    {
        return Icon(
            Shape("M2.5,4.5 H8 L9.5,6.5 H16.5 V16 H2.5 Z", folderBack, folderOutline),
            Shape("M2.5,16 L4.8,9 H18.5 L16.5,16 Z", folderFront, folderOutline));
    }

    public static Control Save()
    {
        return Icon(
            Shape("M3,3 H14.5 L17,5.5 V17 H3 Z", diskBody, diskOutline),
            Shape("M6,3 H13 V7.5 H6 Z", paper, diskOutline),
            Shape("M10.8,4 V6.5", null, diskOutline, thickness: 1.4),
            Shape("M5.5,11 H14.5 V17 H5.5 Z", diskLabel, diskOutline));
    }

    /// <summary>Tiles, in the pastels of the gallery</summary>
    public static Control Gallery()
    {
        return Icon(
            Shape("M3.5,3.5 H9 V9 H3.5 Z", tileBlue, outline),
            Shape("M11,3.5 H16.5 V9 H11 Z", tileYellow, outline),
            Shape("M3.5,11 H9 V16.5 H3.5 Z", tileGreen, outline),
            Shape("M11,11 H16.5 V16.5 H11 Z", tilePink, outline));
    }

    // the chevrons sit a little towards the count between them
    public static Control Previous()
    {
        return Icon(Shape("M13.5,4.5 L8,10 L13.5,15.5", null, arrow, thickness: 2.2));
    }

    public static Control Next()
    {
        return Icon(Shape("M6.5,4.5 L12,10 L6.5,15.5", null, arrow, thickness: 2.2));
    }

    public static Control Undo()
    {
        return Icon(
            Shape("M4.5,8 H12.5 A4.25,4.25 0 0 1 12.5,16.5 H8", null, arrow, thickness: 2),
            Shape("M8,4 L4,8 L8,12", null, arrow, thickness: 2));
    }

    public static Control Redo()
    {
        return Icon(
            Shape("M15.5,8 H7.5 A4.25,4.25 0 0 0 7.5,16.5 H12", null, arrow, thickness: 2),
            Shape("M12,4 L16,8 L12,12", null, arrow, thickness: 2));
    }

    static Control Icon(params Path[] shapes)
    {
        var canvas = new Canvas() { Width = Size, Height = Size };
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
