using System;
using System.Collections.Generic;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// A small tab strip ("Swatches | Spectrum", "Solid | Gradient") in the visual language of
/// the ribbon: a line along the bottom, and the current caption drawn as a tab
/// (<see cref="TabOutline"/>) that opens into whatever is placed below the strip.
/// Hides itself while there is nothing to choose between.
/// </summary>
public class SegmentSwitcher : Panel
{
    const double Flare = 4;

    readonly StackPanel segments = new StackPanel() { Orientation = Orientation.Horizontal };
    readonly List<(object Value, TabOutline Outline, TextBlock Caption)> entries = new List<(object, TabOutline, TextBlock)>();

    public SegmentSwitcher()
    {
        // behind the tabs, so that the selected one paints over it
        Children.Add(new Border()
        {
            BorderBrush = RibbonTheme.TabLine,
            BorderThickness = new Thickness(0, 0, 0, 1)
        });
        Children.Add(segments);
        IsVisible = false;
    }

    IBrush surface = RibbonTheme.Background;

    /// <summary>The background of the area below the strip; the selected tab is filled with it.</summary>
    public IBrush Surface
    {
        get => surface;
        set
        {
            surface = value;
            foreach (var entry in entries)
            {
                entry.Outline.Surface = value;
                entry.Outline.InvalidateVisual();
            }
        }
    }

    /// <summary>The user clicked a segment; the argument is what was passed to <see cref="Add"/>.</summary>
    public event Action<object> Selected;

    public void Add(string caption, object value)
    {
        var outline = new TabOutline() { Flare = Flare, TopRadius = 5, Surface = surface };
        var text = new TextBlock()
        {
            Text = caption,
            FontSize = 11,
            Margin = new Thickness(Flare + 9, 4, Flare + 9, 5),
            VerticalAlignment = VerticalAlignment.Center
        };

        var segment = new Panel()
        {
            Background = Brushes.Transparent,
            Cursor = new Cursor(StandardCursorType.Hand)
        };
        segment.Children.Add(outline);
        segment.Children.Add(text);
        segment.PointerEntered += (s, e) => outline.IsHovered = true;
        segment.PointerExited += (s, e) => outline.IsHovered = false;
        segment.PointerPressed += (s, e) =>
        {
            Current = value;
            Selected?.Invoke(value);
        };

        entries.Add((value, outline, text));
        segments.Children.Add(segment);
        IsVisible = entries.Count > 1;
        Current = current;
    }

    object current;

    /// <summary>Setting it does not raise <see cref="Selected"/>.</summary>
    public object Current
    {
        get => current;
        set
        {
            current = value;
            foreach (var entry in entries)
            {
                bool isCurrent = Equals(entry.Value, value);
                entry.Outline.IsSelected = isCurrent;
                entry.Caption.Foreground = isCurrent ? RibbonTheme.TabHeaderTextSelected : RibbonTheme.TabHeaderText;
            }
        }
    }
}
