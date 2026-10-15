using System.Collections.Generic;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// A page of color swatches from a <see cref="ColorPalette"/>.
/// Meant to be played with: <see cref="SwatchSize"/>, <see cref="Spacing"/> and
/// <see cref="SwatchCornerRadius"/> change the look without code, and a subclass can
/// replace the swatch visual (<see cref="CreateSwatch"/>, <see cref="ShowSelected"/>) or the
/// whole arrangement (<see cref="CreateLayout"/>) - stars in a spiral, if that's more fun.
/// </summary>
public class SwatchPage : ColorPage
{
    readonly Dictionary<Control, NamedColor> swatches = new Dictionary<Control, NamedColor>();

    /// <summary>The swatches in palette order, row by row: what the arrow keys move through</summary>
    readonly List<Control> swatchesInOrder = new List<Control>();

    Control selectedSwatch;

    public SwatchPage()
        : this(ColorPalette.WebColors)
    {
    }

    public SwatchPage(ColorPalette palette)
    {
        this.palette = palette;

        // takes the keyboard when a swatch is clicked, for the arrow keys (OnKeyDown)
        Focusable = true;
    }

    public override string Title => "Swatches";

    ColorPalette palette;
    public ColorPalette Palette
    {
        get => palette;
        set
        {
            palette = value;
            Rebuild();
        }
    }

    // 15 + 1 of spacing: the 14 columns come to the same 224 as the rest of the picker
    double swatchSize = 15;
    public double SwatchSize
    {
        get => swatchSize;
        set
        {
            swatchSize = value;
            Rebuild();
        }
    }

    double spacing = 1;

    /// <summary>Gap between neighboring swatches</summary>
    public double Spacing
    {
        get => spacing;
        set
        {
            spacing = value;
            Rebuild();
        }
    }

    double swatchCornerRadius = 2.5;
    public double SwatchCornerRadius
    {
        get => swatchCornerRadius;
        set
        {
            swatchCornerRadius = value;
            Rebuild();
        }
    }

    IBrush swatchBorderBrush = null;

    /// <summary>Outline of every swatch (1 px); null for none.</summary>
    public IBrush SwatchBorderBrush
    {
        get => swatchBorderBrush;
        set
        {
            swatchBorderBrush = value;
            Rebuild();
        }
    }

    protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnAttachedToVisualTree(e);
        if (Child == null)
        {
            Rebuild();
        }
    }

    void Rebuild()
    {
        swatches.Clear();
        swatchesInOrder.Clear();
        selectedSwatch = null;

        foreach (var named in palette.Colors)
        {
            var swatch = CreateSwatch(named);
            swatch.PointerPressed += Swatch_PointerPressed;
            if (named.Name.Length > 0)
            {
                ToolTip.SetTip(swatch, named.Name + "\n" + ColorText.ToHex(named.Color));
            }

            swatches.Add(swatch, named);
            swatchesInOrder.Add(swatch);
        }

        Child = CreateLayout(swatchesInOrder.ToArray());
        OnColorSet(Color);
    }

    void Swatch_PointerPressed(object sender, PointerPressedEventArgs e)
    {
        var swatch = (Control)sender;
        Select(swatch);
        Pick(swatches[swatch].Color);

        // the arrow keys go on from here
        Focus(NavigationMethod.Pointer);
    }

    /// <summary>
    /// An arrow key moves the current swatch to the next one that way, in its row or its
    /// column, and picks that color as a click would (the editor makes a run of them one
    /// undo step, as it does a run of clicks). Without a current swatch - a color the palette
    /// doesn't have - the first press goes to the swatch nearest the color.
    /// </summary>
    protected override void OnKeyDown(KeyEventArgs e)
    {
        base.OnKeyDown(e);
        if (e.Handled || e.KeyModifiers != KeyModifiers.None)
        {
            return;
        }

        (int Columns, int Rows) step = e.Key switch
        {
            Key.Left => (-1, 0),
            Key.Right => (1, 0),
            Key.Up => (0, -1),
            Key.Down => (0, 1),
            _ => (0, 0)
        };
        if (step == (0, 0))
        {
            return;
        }

        // at the edge too: the side panel's scroll viewer would take the key and scroll
        e.Handled = true;
        var target = selectedSwatch == null
            ? FindNearest(Color)
            : FindNeighbor(selectedSwatch, step.Columns, step.Rows);
        if (target != null && target != selectedSwatch)
        {
            Select(target);
            Pick(swatches[target].Color);
        }
    }

    /// <summary>
    /// The next swatch with a color from the given one, stepping through the grid of
    /// <see cref="CreateLayout"/>, past holes in the palette; null at the edge
    /// </summary>
    Control FindNeighbor(Control swatch, int columnStep, int rowStep)
    {
        int columns = System.Math.Max(palette.Columns, 1);
        int index = swatchesInOrder.IndexOf(swatch);
        int column = index % columns;
        int row = index / columns;
        while (true)
        {
            column += columnStep;
            row += rowStep;
            int next = row * columns + column;
            if (column < 0 || column >= columns || row < 0 || next >= swatchesInOrder.Count)
            {
                return null;
            }

            var candidate = swatchesInOrder[next];
            if (swatches[candidate].Name.Length > 0)
            {
                return candidate;
            }
        }
    }

    /// <summary>The swatch whose color is nearest the given one</summary>
    Control FindNearest(Color color)
    {
        Control nearest = null;
        double nearestDistance = double.MaxValue;
        foreach (var swatch in swatchesInOrder)
        {
            var named = swatches[swatch];
            if (named.Name.Length == 0)
            {
                continue;
            }

            var swatchColor = named.Color;
            double distance = Square(swatchColor.R - color.R)
                + Square(swatchColor.G - color.G)
                + Square(swatchColor.B - color.B)
                + Square(swatchColor.A - color.A);
            if (distance < nearestDistance)
            {
                nearestDistance = distance;
                nearest = swatch;
            }
        }

        return nearest;

        static double Square(double value) => value * value;
    }

    /// <summary>The visual of one swatch. Whatever it returns gets the click and the tooltip.</summary>
    protected virtual Control CreateSwatch(NamedColor named)
    {
        return new Border()
        {
            Width = SwatchSize,
            Height = SwatchSize,
            // the whole gap on one side: half a pixel on each would blur at 100% scaling
            Margin = new Thickness(0, 0, Spacing, Spacing),
            CornerRadius = new CornerRadius(SwatchCornerRadius),
            BorderBrush = SwatchBorderBrush,
            BorderThickness = new Thickness(SwatchBorderBrush != null ? 1 : 0),
            Background = named.Color.A == 0 ? ColorText.CheckerboardBrush : new SolidColorBrush(named.Color),
            Cursor = new Cursor(StandardCursorType.Hand)
        };
    }

    /// <summary>Marks or unmarks a swatch created by <see cref="CreateSwatch"/> as the current color.</summary>
    protected virtual void ShowSelected(Control swatch, NamedColor named, bool isSelected)
    {
        var border = (Border)swatch;

        // a dark ring with a white one inside it: reads on any swatch color and on any neighbor
        border.BorderBrush = isSelected ? Brushes.Black : SwatchBorderBrush;
        border.BorderThickness = new Thickness(isSelected ? 2 : (SwatchBorderBrush != null ? 1 : 0));
        border.BoxShadow = isSelected
            ? new BoxShadows(new BoxShadow() { IsInset = true, Spread = 3.5, Color = Colors.White })
            : default;
        border.ZIndex = isSelected ? 1 : 0;
    }

    /// <summary>Arranges the swatches (given in palette order). The default is a plain grid.</summary>
    protected virtual Control CreateLayout(IReadOnlyList<Control> swatchControls)
    {
        var grid = new UniformGrid()
        {
            Columns = palette.Columns,
            HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Left
        };
        grid.Children.AddRange(swatchControls);
        return grid;
    }

    protected override void OnColorSet(Color newColor)
    {
        Control match = null;
        foreach (var pair in swatches)
        {
            if (pair.Value.Color == newColor && pair.Value.Name.Length > 0)
            {
                match = pair.Key;
                break;
            }
        }

        Select(match);
    }

    void Select(Control swatch)
    {
        if (selectedSwatch == swatch)
        {
            return;
        }

        if (selectedSwatch != null)
        {
            ShowSelected(selectedSwatch, swatches[selectedSwatch], isSelected: false);
        }

        selectedSwatch = swatch;
        if (selectedSwatch != null)
        {
            ShowSelected(selectedSwatch, swatches[selectedSwatch], isSelected: true);
        }
    }
}
