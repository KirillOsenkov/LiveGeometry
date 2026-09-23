using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.Controls.Templates;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Styling;
using AvaloniaSetter = Avalonia.Styling.Setter;
using AvaloniaStyle = Avalonia.Styling.Style;

namespace DynamicGeometry;

/// <summary>
/// Makes the editors inside a property grid (text boxes, buttons, the style picker, labels)
/// look like they belong to the toolbar: same colors, same compact text size. Done with
/// styles scoped to the grid, so the editors themselves stay plain controls.
/// </summary>
public static class PropertyGridTheme
{
    public const double FontSize = 12;

    static readonly IBrush inputBackground = Brushes.White;
    static readonly CornerRadius inputCornerRadius = new CornerRadius(4);

    public static void Apply(PropertyGrid propertyGrid)
    {
        var styles = propertyGrid.Styles;

        styles.Add(new AvaloniaStyle(x => x.OfType<TextBlock>())
        {
            Setters =
            {
                new AvaloniaSetter(TextBlock.FontSizeProperty, FontSize),
                new AvaloniaSetter(TextBlock.ForegroundProperty, RibbonTheme.Text)
            }
        });

        styles.Add(new AvaloniaStyle(x => x.OfType<TextBox>())
        {
            Setters =
            {
                new AvaloniaSetter(TextBox.FontSizeProperty, FontSize),
                new AvaloniaSetter(TextBox.MinHeightProperty, 26.0),

                // a long text (a caption, an error) wraps instead of stretching the panel
                // across the window
                new AvaloniaSetter(TextBox.MaxWidthProperty, 480.0),
                new AvaloniaSetter(TextBox.TextWrappingProperty, TextWrapping.Wrap),
                new AvaloniaSetter(TextBox.PaddingProperty, new Thickness(6, 4, 6, 3)),
                new AvaloniaSetter(TextBox.MarginProperty, new Thickness(0, 2, 0, 2)),
                new AvaloniaSetter(TextBox.BackgroundProperty, inputBackground),
                new AvaloniaSetter(TextBox.BorderBrushProperty, RibbonTheme.TabLine),
                new AvaloniaSetter(TextBox.CornerRadiusProperty, inputCornerRadius),

                // selected text: the same light blue as a checked plate, and the text stays
                // dark on it (the theme's default is a saturated blue with white text)
                new AvaloniaSetter(TextBox.SelectionBrushProperty, RibbonTheme.ButtonChecked),
                new AvaloniaSetter(TextBox.SelectionForegroundBrushProperty, RibbonTheme.Text)
            }
        });

        styles.Add(new AvaloniaStyle(x => x.OfType<Button>())
        {
            Setters =
            {
                new AvaloniaSetter(Button.FontSizeProperty, FontSize),
                new AvaloniaSetter(Button.PaddingProperty, new Thickness(12, 5, 12, 5)),
                new AvaloniaSetter(Button.BackgroundProperty, RibbonTheme.ButtonHover),
                new AvaloniaSetter(Button.BorderBrushProperty, RibbonTheme.TabLine),
                new AvaloniaSetter(Button.BorderThicknessProperty, new Thickness(1)),
                new AvaloniaSetter(Button.CornerRadiusProperty, RibbonTheme.ButtonCornerRadius)
            }
        });
        styles.Add(TemplatePartBackground<Button>(":pointerover", RibbonTheme.ButtonPressed));
        styles.Add(TemplatePartBackground<Button>(":pressed", RibbonTheme.ButtonChecked));

        styles.Add(new AvaloniaStyle(x => x.OfType<CheckBox>())
        {
            Setters =
            {
                new AvaloniaSetter(CheckBox.FontSizeProperty, FontSize),
                new AvaloniaSetter(CheckBox.MinHeightProperty, 26.0)
            }
        });

        styles.Add(new AvaloniaStyle(x => x.OfType<ComboBox>())
        {
            Setters =
            {
                new AvaloniaSetter(ComboBox.FontSizeProperty, FontSize),
                new AvaloniaSetter(ComboBox.MinHeightProperty, 26.0),
                new AvaloniaSetter(ComboBox.BackgroundProperty, inputBackground),
                new AvaloniaSetter(ComboBox.BorderBrushProperty, RibbonTheme.TabLine),
                new AvaloniaSetter(ComboBox.CornerRadiusProperty, inputCornerRadius),
                new AvaloniaSetter(ComboBox.HorizontalAlignmentProperty, HorizontalAlignment.Stretch)
            }
        });

        // the style picker: rows of swatches instead of a tall list. The panel's scroll viewer
        // measures with unlimited width, so without a cap the row would never wrap and a
        // drawing with many styles made the panel run across the whole window.
        styles.Add(new AvaloniaStyle(x => x.OfType<ListBox>())
        {
            Setters =
            {
                new AvaloniaSetter(ListBox.BackgroundProperty, Brushes.Transparent),
                new AvaloniaSetter(ListBox.MarginProperty, new Thickness(0, 2, 0, 2)),
                new AvaloniaSetter(ListBox.MaxWidthProperty, 340.0),
                new AvaloniaSetter(ListBox.ItemsPanelProperty, new FuncTemplate<Panel>(() => new WrapPanel()))
            }
        });
        styles.Add(new AvaloniaStyle(x => x.OfType<ListBoxItem>())
        {
            Setters =
            {
                new AvaloniaSetter(ListBoxItem.PaddingProperty, new Thickness(0)),
                new AvaloniaSetter(ListBoxItem.MarginProperty, new Thickness(0, 0, 3, 3)),
                new AvaloniaSetter(ListBoxItem.CornerRadiusProperty, RibbonTheme.ButtonCornerRadius)
            }
        });
        styles.Add(TemplatePartBackground<ListBoxItem>(":pointerover", RibbonTheme.ButtonHover));
        styles.Add(TemplatePartBackground<ListBoxItem>(":selected", RibbonTheme.ButtonChecked));
        styles.Add(TemplatePartBackground<ListBoxItem>(":selected:pointerover", RibbonTheme.ButtonChecked));
    }

    /// <summary>
    /// The Fluent theme paints hover/pressed/selected states on the ContentPresenter inside
    /// the control's template, so that is where they have to be overridden.
    /// </summary>
    static AvaloniaStyle TemplatePartBackground<T>(string pseudoClasses, IBrush background) where T : Control
    {
        return new AvaloniaStyle(x =>
        {
            var selector = x.OfType<T>();
            foreach (var pseudoClass in pseudoClasses.Split(':', System.StringSplitOptions.RemoveEmptyEntries))
            {
                selector = selector.Class(":" + pseudoClass);
            }

            return selector.Template().OfType<ContentPresenter>();
        })
        {
            Setters =
            {
                new AvaloniaSetter(ContentPresenter.BackgroundProperty, background)
            }
        };
    }
}
