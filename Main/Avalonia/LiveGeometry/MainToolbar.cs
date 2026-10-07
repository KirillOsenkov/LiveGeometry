using System;
using System.Linq;
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
        Spacing = NormalSpacing,
        Margin = NormalMargin
    };

    // where buttons are added: the strip itself or the group that is open
    Panel current;

    // the small things at the far right (the theme button, the build stamp), as one
    readonly StackPanel rightControls = new StackPanel()
    {
        Orientation = Orientation.Horizontal,
        Spacing = 2
    };
    Control rightControl;
    double rightWidth; // as last measured; kept while it is hidden

    // the group between the buttons and the right control
    MainToolbarGroup centered;

    // decided by the measure pass for the arrange pass
    bool wrapped;
    double firstRowHeight;

    // between a wrapped group and the ribbon under it
    const double WrappedRowGap = 4;

    // along the bottom edge, like the line under the ribbon's group headers: the tab of the
    // open ribbon (below) swings up out of it, and it is what the strip ends in when the ribbon
    // is folded and the canvas is right under it
    readonly Border bottomLine = new Border()
    {
        Height = 1
    };

    // The button that stands for the ribbon: while the ribbon is open it is drawn as a tab
    // opening into the ribbon's header row, in the same shape as the selected group header.
    // Its fill is the tools strip's light gray, turning into the header row's gray along the
    // last stretch, so that the tab seems to open into the row it sits on.
    MainToolbarButton tabButton;
    bool isTabOpen;
    bool showTab; // decided by the measure pass: not when the group has wrapped
    const double TabGap = 2; // between the button's plate and the tab's outline

    readonly TabOutline tab = new TabOutline()
    {
        IsSelected = true
    };

    public MainToolbar()
    {
        current = buttons;

        // A notch darker and more neutral than the ribbon's header row right under it; the line
        // between the two is what the tab of the open ribbon opens through.
        this.BindTheme(BackgroundProperty, nameof(AppTheme.Strip));
        bottomLine.BindTheme(Border.BackgroundProperty, nameof(AppTheme.TabLine));
        Children.Add(bottomLine);
        Children.Add(tab);
        Children.Add(buttons);

        // the tab's fill is a gradient of two theme colors: built again whenever either changes
        tab.BindTheme(TabOutline.SurfaceProperty, key: null);
        this.ObserveTheme(nameof(AppTheme.Background), color =>
        {
            tabBackground = color;
            UpdateTabSurface();
        });
        this.ObserveTheme(nameof(AppTheme.HeaderRow), color =>
        {
            tabHeaderRow = color;
            UpdateTabSurface();
        });
    }

    Color? tabBackground;
    Color? tabHeaderRow;

    void UpdateTabSurface()
    {
        if (tabBackground == null || tabHeaderRow == null)
        {
            return;
        }

        tab.Surface = new LinearGradientBrush()
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(0, 1, RelativeUnit.Relative),
            GradientStops =
            {
                new GradientStop(tabBackground.Value, 0),
                new GradientStop(tabBackground.Value, 0.9),
                new GradientStop(tabHeaderRow.Value, 1)
            }
        };
    }

    /// <summary>The button that is drawn as a tab while <see cref="IsTabOpen"/></summary>
    public void SetTabButton(MainToolbarButton button)
    {
        tabButton = button;
        InvalidateMeasure();
    }

    /// <summary>
    /// Whether the tab button opens into what is under the strip. When the strip has wrapped
    /// to two rows the tab would have to span both, so the button shows its checked plate
    /// instead.
    /// </summary>
    public bool IsTabOpen
    {
        get => isTabOpen;
        set
        {
            if (isTabOpen != value)
            {
                isTabOpen = value;
                InvalidateMeasure();
            }
        }
    }

    /// <summary>
    /// The centered group and the right control are inset from the top by this much, so that
    /// they are centered on the same line as each other (the buttons have their own margin).
    /// </summary>
    public const double RowTopInset = 2;

    /// <summary>Something small at the far right (the theme button, the build stamp), after what is there</summary>
    public void AddAtRight(Control control)
    {
        if (rightControl == null)
        {
            rightControl = new Border()
            {
                Padding = new Thickness(0, RowTopInset, 0, 0),
                Child = rightControls
            };
            Children.Add(rightControl);
        }

        rightControls.Children.Add(control);
    }

    // the buttons closer together, for a narrow window
    bool isCompact;

    const double NormalSpacing = 2;
    const double NormalSeparatorMargin = 5;
    const double CompactSeparatorMargin = 2;
    const double SeparatorInset = 5; // above and below, within the row of buttons
    static readonly Thickness NormalMargin = new Thickness(8, 4, 6, 2);
    // (the same at the left: the tab drawn around the first button reaches past it)
    static readonly Thickness CompactMargin = new Thickness(8, 4, 2, 2);

    /// <summary>
    /// On a phone the buttons as they are need more than the width of the screen (about
    /// 410 pixels), and the last of them - the settings - was cut off with no way to reach
    /// it. Closer together, square and without room around the separators, they fit in 330.
    /// </summary>
    void SetCompact(bool compact)
    {
        isCompact = compact;
        buttons.Spacing = compact ? 0 : NormalSpacing;
        buttons.Margin = compact ? CompactMargin : NormalMargin;
        foreach (var child in buttons.Children)
        {
            if (child is MainToolbarButton button)
            {
                button.Width = compact ? button.Height : button.Height + MainToolbarButton.ExtraWidth;
            }
            else if (child is Border separator && separator.Width == 1)
            {
                double margin = compact ? CompactSeparatorMargin : NormalSeparatorMargin;
                separator.Margin = new Thickness(margin, SeparatorInset, margin, SeparatorInset);
            }
        }
    }

    /// <summary>How much wider the buttons are when they are not compact</summary>
    double CompactSavings()
    {
        int buttonCount = buttons.Children.OfType<MainToolbarButton>().Count();
        int separatorCount = buttons.Children.Count(child => !(child is MainToolbarButton) && child is Border border && border.Width == 1);
        return buttonCount * MainToolbarButton.ExtraWidth
            + separatorCount * 2 * (NormalSeparatorMargin - CompactSeparatorMargin)
            + System.Math.Max(0, buttons.Children.Count - 1) * NormalSpacing
            + NormalMargin.Left + NormalMargin.Right - CompactMargin.Left - CompactMargin.Right;
    }

    protected override Size MeasureOverride(Size availableSize)
    {
        var unbounded = new Size(double.PositiveInfinity, double.PositiveInfinity);
        buttons.Measure(unbounded);

        // changed only when the width asks for the other way, or every pass would undo the last
        if (!isCompact && buttons.DesiredSize.Width > availableSize.Width)
        {
            SetCompact(true);
            buttons.Measure(unbounded);
        }
        else if (isCompact && buttons.DesiredSize.Width + CompactSavings() <= availableSize.Width)
        {
            SetCompact(false);
            buttons.Measure(unbounded);
        }

        double width = buttons.DesiredSize.Width;
        double height = buttons.DesiredSize.Height;

        // the group beside the buttons if all of it fits there (the stamp may have to go),
        // else on a row of its own, where the trailing text still gets cut short if it must
        bool hasCentered = centered != null && centered.IsVisible;
        double centeredWidth = 0;
        wrapped = false;
        if (hasCentered)
        {
            centered.HasSeparator = IsGroupLeftAligned;
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
            centered.IsLeftAligned = IsGroupLeftAligned || wrapped;
            centered.HasSeparator = IsGroupLeftAligned && !wrapped;
            centered.Measure(new Size(room, double.PositiveInfinity));
            if (wrapped)
            {
                width = System.Math.Max(width, centered.DesiredSize.Width);
                height += centered.DesiredSize.Height + WrappedRowGap;
            }
        }

        bottomLine.Measure(unbounded);
        height += bottomLine.Height;

        showTab = isTabOpen && tabButton != null && !wrapped;
        if (tabButton != null)
        {
            tabButton.IsChecked = isTabOpen && !showTab;
        }

        tab.Measure(unbounded);
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

        bottomLine.Arrange(new Rect(0, finalSize.Height - bottomLine.Height, finalSize.Width, bottomLine.Height));

        // a little bigger than the button (its hover plate would paint over the outline), wider
        // still by the flare of its feet, and down over the bottom line. Not when the strip is
        // shorter than that: a page with no height (a tab opened in the background) gets a
        // strip of none, and a rectangle of negative height throws
        var tabRect = new Rect();
        if (showTab)
        {
            var button = tabButton.Bounds;
            double top = buttons.Bounds.Y + button.Y - TabGap;
            if (finalSize.Height > top)
            {
                tabRect = new Rect(
                    buttons.Bounds.X + button.X - TabGap - tab.Flare,
                    top,
                    button.Width + 2 * (TabGap + tab.Flare),
                    finalSize.Height - top);
            }
        }

        tab.Arrange(tabRect);

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
        current.Children.Add(CreateSeparator());
    }

    Border CreateSeparator()
    {
        var separator = new Border()
        {
            Width = 1,
            Margin = new Thickness(NormalSeparatorMargin, SeparatorInset, NormalSeparatorMargin, SeparatorInset)
        };
        separator.BindTheme(Border.BackgroundProperty, nameof(AppTheme.TabLine));
        return separator;
    }

    public TextBlock AddText(FontWeight weight, double minWidth = 0, double fontSize = 13)
    {
        var text = new TextBlock()
        {
            FontSize = fontSize,
            FontWeight = weight,
            MinWidth = minWidth,
            TextAlignment = TextAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
        text.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Text));
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
        // in line with the separators among the buttons: as far from the button before it and
        // the one after it (past the strip's margin and the group's), and as tall
        var separator = CreateSeparator();
        separator.Margin = new Thickness(
            NormalSpacing + NormalSeparatorMargin - NormalMargin.Right,
            NormalMargin.Top + SeparatorInset,
            NormalSpacing + NormalSeparatorMargin - MainToolbarGroup.MiddleMargin,
            NormalMargin.Bottom + SeparatorInset);
        centered = new MainToolbarGroup(buttons.Spacing, separator);
        Children.Add(centered);
        current = centered.Middle;
        return centered;
    }

    /// <summary>
    /// Whether the group sits right after the buttons, behind a separator, rather than in the
    /// middle of the room the buttons leave
    /// </summary>
    public bool IsGroupLeftAligned { get; set; } = true;

    /// <summary>Text to the right of the centered group, cut short when there is no room</summary>
    public TextBlock AddTrailingText(FontWeight weight, double fontSize)
    {
        var text = new TextBlock()
        {
            FontSize = fontSize,
            FontWeight = weight,
            TextTrimming = TextTrimming.CharacterEllipsis,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(6, 0, 6, 0)
        };
        text.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Text));
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
/// everything to the left instead (beside the buttons, or on a row of its own), and
/// <see cref="HasSeparator"/> sets it off from the buttons.
/// </summary>
public class MainToolbarGroup : Panel
{
    /// <summary>Around <see cref="Middle"/> at the left and the right</summary>
    public const double MiddleMargin = 6;

    public StackPanel Middle { get; }
    public Panel Trailing { get; }

    readonly Control separator;
    bool isLeftAligned;
    double trailingNaturalWidth;

    public MainToolbarGroup(double spacing, Control separator)
    {
        this.separator = separator;
        Middle = new StackPanel()
        {
            Orientation = Orientation.Horizontal,
            Spacing = spacing,
            Margin = new Thickness(MiddleMargin, MainToolbar.RowTopInset, MiddleMargin, 0)
        };
        // not a StackPanel: that would hand the text unlimited width and it could never trim
        Trailing = new Panel()
        {
            Margin = new Thickness(0, MainToolbar.RowTopInset, 6, 0)
        };
        Children.Add(separator);
        Children.Add(Middle);
        Children.Add(Trailing);
    }

    /// <summary>Whether the separator shows at the left, before everything else</summary>
    public bool HasSeparator
    {
        get => separator.IsVisible;
        set
        {
            if (separator.IsVisible != value)
            {
                separator.IsVisible = value;
                InvalidateMeasure();
            }
        }
    }

    // a hidden separator measures as nothing
    double SeparatorWidth => separator.DesiredSize.Width;

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
        separator.Measure(unbounded);
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
            SeparatorWidth + Middle.DesiredSize.Width + trailingNaturalWidth,
            System.Math.Max(Middle.DesiredSize.Height, Trailing.DesiredSize.Height));
    }

    protected override Size ArrangeOverride(Size finalSize)
    {
        // its margin is what places it in the height of the row
        separator.Arrange(new Rect(0, 0, SeparatorWidth, finalSize.Height));
        Place(finalSize.Width, out double left, out double trailingWidth);
        double middleWidth = Middle.DesiredSize.Width;
        Middle.Arrange(new Rect(left, 0, middleWidth, finalSize.Height));
        Trailing.Arrange(new Rect(left + middleWidth, 0, trailingWidth, finalSize.Height));
        return finalSize;
    }

    /// <summary>
    /// Where the middle starts, and how much the text after it gets; with the middle hidden
    /// (a file's name alone) the text itself is centered
    /// </summary>
    void Place(double width, out double left, out double trailingWidth)
    {
        double start = SeparatorWidth;
        width -= start;
        double middleWidth = Middle.DesiredSize.Width;
        double centeredLeft = middleWidth > 0 ? (width - middleWidth) / 2 : (width - trailingNaturalWidth) / 2;
        double leftForWholeText = width - middleWidth - trailingNaturalWidth;
        left = isLeftAligned ? 0 : System.Math.Max(0, System.Math.Min(centeredLeft, leftForWholeText));
        trailingWidth = System.Math.Max(0, width - left - middleWidth);
        left += start;
    }
}

public class MainToolbarButton : Border, ICommandObserver
{
    /// <summary>How big the icons are drawn (their grid is <see cref="MainToolbarIcons.Size"/>)</summary>
    public const double IconSize = 24;

    /// <summary>The plate around the icon (a little more at the sides)</summary>
    public const double DefaultInset = 5;

    /// <summary>A button is this much wider than high (none in a narrow window, see MainToolbar.SetCompact)</summary>
    public const double ExtraWidth = 4;

    readonly Action action;
    bool isPressed;
    bool isChecked;

    public MainToolbarButton(
        Control icon,
        Action action,
        double iconSize = IconSize,
        double inset = DefaultInset)
    {
        this.action = action;
        Width = iconSize + 2 * inset + ExtraWidth;
        Height = iconSize + 2 * inset;
        CornerRadius = AppTheme.ButtonCornerRadius;
        Background = Brushes.Transparent;

        // always there (transparent), so that the icon doesn't shift when the button is checked
        BorderThickness = new Thickness(1);
        BorderBrush = Brushes.Transparent;
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

    /// <summary>For an on/off button: the same blue plate as a checked toggle in the ribbon</summary>
    public bool IsChecked
    {
        get => isChecked;
        set
        {
            if (isChecked != value)
            {
                isChecked = value;
                UpdateBackground(isOver: IsPointerOver);
            }
        }
    }

    void UpdateBackground(bool isOver)
    {
        if (isChecked && !isPressed)
        {
            this.BindTheme(BackgroundProperty, nameof(AppTheme.ButtonChecked));
            this.BindTheme(BorderBrushProperty, nameof(AppTheme.ButtonCheckedBorder));
            return;
        }

        // on the header row, which is darker than the tools strip: the hover plate is lighter
        string plate = isPressed ? nameof(AppTheme.ButtonPressed) : (isOver ? nameof(AppTheme.GroupBackground) : null);
        this.BindTheme(BackgroundProperty, plate, whenNone: Brushes.Transparent);
        this.BindTheme(BorderBrushProperty, key: null, whenNone: Brushes.Transparent);
    }

    /// <summary>A new picture in the button (the theme button turns from moon to sun)</summary>
    public void SetIcon(Control icon)
    {
        ((Viewbox)Child).Child = icon;
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

    /// <summary>Stands for the theme's <see cref="AppTheme.IconOutline"/>: <see cref="Shape"/> binds it</summary>
    static readonly IBrush outline = new SolidColorBrush(Colors.Black);

    /// <summary>Stands for the theme's <see cref="AppTheme.Accent"/></summary>
    static readonly IBrush arrow = new SolidColorBrush(Colors.Blue);

    static readonly IBrush paper = Brushes.White;
    static readonly IBrush green = new SolidColorBrush(Color.FromRgb(0x2E, 0x9E, 0x4F));
    static readonly IBrush folderBack = new SolidColorBrush(Color.FromRgb(0xF2, 0xB6, 0x32));
    static readonly IBrush folderFront = new SolidColorBrush(Color.FromRgb(0xFA, 0xD5, 0x65));
    static readonly IBrush folderOutline = new SolidColorBrush(Color.FromRgb(0xA8, 0x7B, 0x05));
    static readonly IBrush diskBody = new SolidColorBrush(Color.FromRgb(0x4C, 0x8B, 0xF5));
    static readonly IBrush diskOutline = new SolidColorBrush(Color.FromRgb(0x2A, 0x5D, 0xB0));
    static readonly IBrush diskLabel = new SolidColorBrush(Color.FromRgb(0xE8, 0xEE, 0xF9));
    static readonly IBrush steel = new LinearGradientBrush()
    {
        StartPoint = new RelativePoint(0.5, 0, RelativeUnit.Relative),
        EndPoint = new RelativePoint(0.5, 1, RelativeUnit.Relative),
        GradientStops =
        {
            new GradientStop(Color.FromRgb(0xF4, 0xF7, 0xFB), 0),
            new GradientStop(Color.FromRgb(0x8A, 0x96, 0xA8), 1)
        }
    };

    static readonly IBrush steelOutline = new SolidColorBrush(Color.FromRgb(0x4F, 0x59, 0x6B));
    static readonly IBrush steelHighlight = new SolidColorBrush(Color.FromArgb(0x70, 0xFF, 0xFF, 0xFF));
    static readonly IBrush tileBlue =new SolidColorBrush(Color.FromRgb(0xBF, 0xDC, 0xFF));
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

    /// <summary>A picture on its way out: the canvas as an image, to a file or the clipboard</summary>
    public static Control Export()
    {
        return Icon(
            Shape("M2.5,3.5 H15.5 V14.5 H2.5 Z", paper, null),
            Shape("M6.2,5.6 A1.5,1.5 0 1 1 6.19,5.6 Z", folderBack, null),
            Shape("M2.5,14.5 V12.5 L6.5,8.5 L9.5,11.5 L11.5,9.5 L15.5,13.5 V14.5 Z", green, null),
            Shape("M2.5,3.5 H15.5 V14.5 H2.5 Z", null, outline),
            Shape("M14.5,10.5 A4,4 0 1 1 14.49,10.5 Z", diskBody, null),
            Shape("M12.3,14.5 H16.5 M14.7,12.6 L16.6,14.5 L14.7,16.4", null, Brushes.White, thickness: 1.6));
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

    /// <summary>
    /// A gear of steel, lit from above: the settings page. Its hub is a hole, with a faint
    /// light ring around it.
    /// </summary>
    public static Control Settings()
    {
        var gear = Gear(
            teeth: 8,
            outer: 8.7,
            root: 6.5,
            topHalf: 0.17,
            baseHalf: 0.25);
        return Icon(
            Shape(gear + Ring(radius: 3), steel, steelOutline, thickness: 1.1),
            Shape(Ring(radius: 4.4), null, steelHighlight, thickness: 0.8));
    }

    /// <summary>
    /// The outline of a gear around the middle of the grid, a tooth pointing up: tops and
    /// roots are arcs of their circles, the flanks straight
    /// </summary>
    /// <param name="topHalf">Half of the angle a tooth takes of the outer circle</param>
    /// <param name="baseHalf">Half of the angle it takes of the root circle</param>
    static string Gear(
        int teeth,
        double outer,
        double root,
        double topHalf,
        double baseHalf)
    {
        const double center = 10;
        double step = 2 * System.Math.PI / teeth;

        // even-odd, so that a ring added to the outline is a hole
        var sb = new System.Text.StringBuilder("F0 ");
        for (int i = 0; i < teeth; i++)
        {
            double angle = i * step - System.Math.PI / 2;
            sb.Append(i == 0 ? "M" : "L").Append(PolarPoint(center, root, angle - baseHalf));
            sb.Append(" L").Append(PolarPoint(center, outer, angle - topHalf));
            sb.Append(Arc(outer)).Append(PolarPoint(center, outer, angle + topHalf));
            sb.Append(" L").Append(PolarPoint(center, root, angle + baseHalf));
            sb.Append(Arc(root)).Append(PolarPoint(center, root, angle + step - baseHalf)).Append(' ');
        }

        sb.Append("Z ");
        return sb.ToString();
    }

    static string Arc(double radius)
    {
        return string.Format(System.Globalization.CultureInfo.InvariantCulture, " A{0},{0} 0 0 1 ", radius);
    }

    /// <summary>A circle around the middle of the grid</summary>
    static string Ring(double radius)
    {
        return string.Format(
            System.Globalization.CultureInfo.InvariantCulture,
            "M{0},10 a{1},{1} 0 1 0 {2},0 a{1},{1} 0 1 0 -{2},0 Z ",
            10 - radius,
            radius,
            2 * radius);
    }

    /// <summary>A crescent moon: the dark theme is a click away</summary>
    public static Control Moon()
    {
        return Icon(Shape("M12.5,3.2 A7,7 0 1 0 16.8,12.6 A5.6,5.6 0 0 1 12.5,3.2 Z", null, outline, thickness: 1.5));
    }

    /// <summary>The sun: the light theme is a click away</summary>
    public static Control Sun()
    {
        var rays = new System.Text.StringBuilder();
        for (int i = 0; i < 8; i++)
        {
            double angle = i * System.Math.PI / 4;
            rays.Append("M").Append(PolarPoint(10, 6.2, angle)).Append(" L").Append(PolarPoint(10, 8.6, angle)).Append(' ');
        }

        return Icon(
            Shape("M10,10 m-3.6,0 a3.6,3.6 0 1 0 7.2,0 a3.6,3.6 0 1 0 -7.2,0", null, outline, thickness: 1.5),
            Shape(rays.ToString(), null, outline, thickness: 1.5));
    }

    static string PolarPoint(double center, double radius, double angle)
    {
        return string.Format(
            System.Globalization.CultureInfo.InvariantCulture,
            "{0:0.##},{1:0.##}",
            center + radius * System.Math.Cos(angle),
            center + radius * System.Math.Sin(angle));
    }

    static Control Icon(params Path[] shapes)
    {
        var canvas = new Canvas() { Width = Size, Height = Size };
        canvas.Children.AddRange(shapes);
        return canvas;
    }

    static Path Shape(string data, IBrush fill, IBrush stroke, double thickness = 1)
    {
        var path = new Path()
        {
            Data = Geometry.Parse(data),
            Fill = fill,
            Stroke = stroke,
            StrokeThickness = stroke != null ? thickness : 0,
            StrokeJoin = PenLineJoin.Round,
            StrokeLineCap = PenLineCap.Round
        };

        // the outlines and the arrows follow the theme; the fills are the colors of the
        // things drawn (a folder, a disk)
        if (stroke == outline)
        {
            path.BindTheme(Avalonia.Controls.Shapes.Shape.StrokeProperty, nameof(AppTheme.IconOutline));
        }
        else if (stroke == arrow)
        {
            path.BindTheme(Avalonia.Controls.Shapes.Shape.StrokeProperty, nameof(AppTheme.Accent));
        }

        return path;
    }
}
