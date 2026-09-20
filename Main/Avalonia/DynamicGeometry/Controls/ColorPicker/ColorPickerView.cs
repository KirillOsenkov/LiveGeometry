using System;
using System.Collections.Generic;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// The color picker: a switcher over any number of <see cref="ColorPage"/>s (swatches and
/// spectrum by default), plus what they share - a sample of the current color and a box to
/// type a color name or hex code into. All pages always show the same color.
/// </summary>
public class ColorPickerView : Decorator
{
    public const double ContentWidth = 224;

    readonly List<ColorPage> pages = new List<ColorPage>();
    readonly SegmentSwitcher switcher = new SegmentSwitcher();
    readonly Decorator pageHost = new Decorator() { Margin = new Thickness(0, 6, 0, 6) };
    readonly Border sample = new Border();
    readonly TextBox textBox = new TextBox() { MinWidth = 0 };

    public ColorPickerView()
        : this(new SwatchPage(), new SpectrumPage())
    {
    }

    public ColorPickerView(params ColorPage[] initialPages)
    {
        var sampleHolder = new Border()
        {
            Width = 40,
            Background = ColorText.CheckerboardBrush,
            BorderBrush = RibbonTheme.TabLine,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            ClipToBounds = true,
            Margin = new Thickness(0, 0, 6, 0),
            Child = sample
        };

        var footer = new DockPanel();
        DockPanel.SetDock(sampleHolder, Dock.Left);
        footer.Children.Add(sampleHolder);
        footer.Children.Add(textBox);

        var layout = new StackPanel() { Width = ContentWidth, HorizontalAlignment = HorizontalAlignment.Left };
        layout.Children.Add(switcher);
        layout.Children.Add(pageHost);
        layout.Children.Add(footer);
        Child = layout;

        textBox.KeyDown += (s, e) =>
        {
            if (e.Key == Key.Enter)
            {
                CommitText();
                e.Handled = true;
            }
        };
        textBox.LostFocus += (s, e) => CommitText();
        switcher.Selected += page => SelectedPage = (ColorPage)page;

        foreach (var page in initialPages)
        {
            AddPage(page);
        }

        UpdateShared();
    }

    public IReadOnlyList<ColorPage> Pages => pages;

    public void AddPage(ColorPage page)
    {
        pages.Add(page);
        page.Color = color;
        page.ColorChanged += picked => SetColor(picked, source: page, notify: true);

        switcher.Add(page.Title, page);
        if (selectedPage == null)
        {
            SelectedPage = page;
        }
    }

    ColorPage selectedPage;
    public ColorPage SelectedPage
    {
        get => selectedPage;
        set
        {
            selectedPage = value;
            pageHost.Child = value;
            switcher.Current = value;
        }
    }

    Color color = Colors.Black;

    /// <summary>Setting it from outside does not raise <see cref="ColorChanged"/>.</summary>
    public Color Color
    {
        get => color;
        set => SetColor(value, source: null, notify: false);
    }

    /// <summary>The user chose a color, on any page or by typing.</summary>
    public event Action<Color> ColorChanged;

    void SetColor(Color newColor, ColorPage source, bool notify)
    {
        color = newColor;
        foreach (var page in pages)
        {
            if (page != source)
            {
                page.Color = newColor;
            }
        }

        UpdateShared();
        if (notify)
        {
            ColorChanged?.Invoke(newColor);
        }
    }

    void UpdateShared()
    {
        sample.Background = new SolidColorBrush(color);

        // don't retype what the user is typing if it already means this color
        if (!ColorText.TryParse(textBox.Text, out var typed) || typed != color)
        {
            textBox.Text = ColorText.Describe(color);
        }
    }

    void CommitText()
    {
        if (ColorText.TryParse(textBox.Text, out var typed))
        {
            if (typed != color)
            {
                SetColor(typed, source: null, notify: true);
            }
        }
        else
        {
            textBox.Text = ColorText.Describe(color);
        }
    }
}
