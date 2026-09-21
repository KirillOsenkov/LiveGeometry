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
/// </summary>
public class MainToolbar : DockPanel
{
    readonly StackPanel buttons = new StackPanel()
    {
        Orientation = Orientation.Horizontal,
        Spacing = 2,
        Margin = new Thickness(8, 4, 6, 0)
    };

    // where buttons are added: the strip itself or the group that is open
    Panel current;

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
        SetDock(control, Dock.Right);
        Children.Insert(0, control);
    }

    public MainToolbarButton AddButton(Control icon, string name, string shortcut, Action action)
    {
        var button = new MainToolbarButton(icon, action);
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

    public TextBlock AddText(FontWeight weight, double minWidth = 0)
    {
        var text = new TextBlock()
        {
            FontSize = 13,
            FontWeight = weight,
            Foreground = RibbonTheme.Text,
            MinWidth = minWidth,
            TextAlignment = TextAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(6, 0, 6, 0)
        };
        current.Children.Add(text);
        return text;
    }

    /// <summary>
    /// What is added from now on goes into a group of its own, to be shown and hidden together
    /// </summary>
    public Panel BeginGroup()
    {
        var group = new StackPanel()
        {
            Orientation = Orientation.Horizontal,
            Spacing = buttons.Spacing
        };
        buttons.Children.Add(group);
        current = group;
        return group;
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

public class MainToolbarButton : Border, ICommandObserver
{
    readonly Action action;
    bool isPressed;

    public MainToolbarButton(Control icon, Action action)
    {
        this.action = action;
        Width = 34;
        Height = 30;
        CornerRadius = RibbonTheme.ButtonCornerRadius;
        Background = Brushes.Transparent;
        Child = new Viewbox()
        {
            Width = MainToolbarIcons.Size,
            Height = MainToolbarIcons.Size,
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

    public static Control Previous()
    {
        return Icon(Shape("M12.5,4.5 L7,10 L12.5,15.5", null, arrow, thickness: 2.2));
    }

    public static Control Next()
    {
        return Icon(Shape("M7.5,4.5 L13,10 L7.5,15.5", null, arrow, thickness: 2.2));
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
