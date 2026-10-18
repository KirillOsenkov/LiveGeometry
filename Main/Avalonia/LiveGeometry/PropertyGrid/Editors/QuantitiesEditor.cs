using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace LiveGeometry;

/// <summary>
/// What can be measured of the figure on its page (<see cref="FigureQuantities"/>): a row
/// per quantity, the caption in the grid's caption column and the value read-only - but
/// selectable, and copied whole by the row's context menu - with a small "+" after it while
/// the drawing has no measurement of it yet, which puts one on the drawing. A quantity an
/// editable row says already (a segment's Length) gets no row of its own; its "+" goes on
/// that row. The values follow the figures as they move (<see cref="PropertyGrid.RefreshNumbers"/>).
/// </summary>
public class QuantitiesEditor : StackPanel
{
    readonly List<(FigureQuantity Quantity, SelectableTextBlock Value, Control Plus)> rows =
        new List<(FigureQuantity Quantity, SelectableTextBlock Value, Control Plus)>();

    public QuantitiesEditor(IEnumerable<FigureQuantity> quantities, IEnumerable<LabeledValueEditor> editors)
    {
        var editorList = editors.ToList();
        foreach (var quantity in quantities)
        {
            var plus = quantity.Measure != null ? CreatePlus(quantity) : null;
            var row = quantity.RowName == null
                ? null
                : editorList.FirstOrDefault(editor => editor.Value?.Name == quantity.RowName || editor.Value?.DisplayName == quantity.RowName);
            if (row != null)
            {
                if (plus != null)
                {
                    plus.Margin = new Thickness(4, 0, 0, 0);
                    row.AddTrailing(plus);
                }

                rows.Add((quantity, null, plus));
                continue;
            }

            rows.Add((quantity, AddRow(quantity, plus), plus));
        }

        Refresh();
    }

    /// <summary>A row like the grid's own: the caption in the shared caption column, then the value, then the "+"</summary>
    SelectableTextBlock AddRow(FigureQuantity quantity, Control plus)
    {
        var grid = new Grid() { MinHeight = 28 };
        grid.ColumnDefinitions.Add(new ColumnDefinition()
        {
            Width = GridLength.Auto,
            MinWidth = 60,
            SharedSizeGroup = "PropertyGridLabel"
        });
        grid.ColumnDefinitions.Add(new ColumnDefinition() { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition() { Width = GridLength.Auto });

        var caption = new TextBlock()
        {
            Text = quantity.Caption,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(2, 0, 12, 0)
        };
        var value = new SelectableTextBlock()
        {
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 6, 0)
        };
        value.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Text));

        // the whole number, not the word up to its decimal point
        value.DoubleTapped += (s, e) => value.SelectAll();

        // the whole value, not the part selected, as a quick way to copy it
        var copy = new MenuItem() { Header = "Copy" };
        copy.Click += (s, e) => Clipboard.SetText(quantity.Text);
        value.ContextMenu = new ContextMenu() { Items = { copy } };

        Grid.SetColumn(value, 1);
        grid.Children.Add(caption);
        grid.Children.Add(value);
        if (plus != null)
        {
            Grid.SetColumn(plus, 2);
            grid.Children.Add(plus);
        }

        Children.Add(grid);
        return value;
    }

    Control CreatePlus(FigureQuantity quantity)
    {
        var button = new Button()
        {
            Width = 22,
            Height = 22,
            Padding = new Thickness(0),
            Background = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            HorizontalContentAlignment = HorizontalAlignment.Center,
            VerticalContentAlignment = VerticalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            Content = PropertyGridIcons.Create(PropertyGridIcon.Plus)
        };
        ToolTip.SetTip(button, quantity.MeasureHint);
        button.Click += (s, e) =>
        {
            quantity.Measure();
            Refresh();
        };
        return button;
    }

    /// <summary>The values as the figures are now, and a "+" only while there is nothing on the drawing for it</summary>
    public void Refresh()
    {
        foreach (var (quantity, value, plus) in rows)
        {
            if (value != null)
            {
                value.Text = quantity.Text;
            }

            if (plus != null)
            {
                plus.IsVisible = quantity.CanMeasure;
            }
        }
    }
}
