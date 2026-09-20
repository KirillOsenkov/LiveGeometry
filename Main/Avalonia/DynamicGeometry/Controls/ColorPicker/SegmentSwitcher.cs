using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// A row of captions of which one is current ("Swatches | Spectrum", "Solid | Gradient").
/// Hides itself while there is nothing to choose between.
/// </summary>
public class SegmentSwitcher : StackPanel
{
    public SegmentSwitcher()
    {
        Orientation = Orientation.Horizontal;
        Spacing = 2;
        IsVisible = false;
    }

    /// <summary>The user clicked a segment; the argument is what was passed to <see cref="Add"/>.</summary>
    public event Action<object> Selected;

    public void Add(string caption, object value)
    {
        var segment = new Border()
        {
            Padding = new Thickness(10, 3, 10, 3),
            CornerRadius = new CornerRadius(4),
            Background = Brushes.Transparent,
            Cursor = new Cursor(StandardCursorType.Hand),
            Child = new TextBlock() { Text = caption, FontSize = 11 },
            Tag = value
        };
        segment.PointerPressed += (s, e) =>
        {
            Current = value;
            Selected?.Invoke(value);
        };
        Children.Add(segment);
        IsVisible = Children.Count > 1;
    }

    object current;

    /// <summary>Setting it does not raise <see cref="Selected"/>.</summary>
    public object Current
    {
        get => current;
        set
        {
            current = value;
            foreach (var child in Children)
            {
                var segment = (Border)child;
                segment.Background = Equals(segment.Tag, value) ? RibbonTheme.ButtonChecked : Brushes.Transparent;
            }
        }
    }
}
