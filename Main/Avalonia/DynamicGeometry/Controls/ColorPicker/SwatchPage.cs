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
    Control selectedSwatch;

    public SwatchPage()
        : this(ColorPalette.WebColors)
    {
    }

    public SwatchPage(ColorPalette palette)
    {
        this.palette = palette;
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

    double swatchSize = 16;
    public double SwatchSize
    {
        get => swatchSize;
        set
        {
            swatchSize = value;
            Rebuild();
        }
    }

    double spacing = 0;

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

    double swatchCornerRadius = 0;
    public double SwatchCornerRadius
    {
        get => swatchCornerRadius;
        set
        {
            swatchCornerRadius = value;
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
        selectedSwatch = null;

        var controls = new List<Control>();
        foreach (var named in palette.Colors)
        {
            var swatch = CreateSwatch(named);
            swatch.PointerPressed += Swatch_PointerPressed;
            if (named.Name.Length > 0)
            {
                ToolTip.SetTip(swatch, named.Name + "\n" + ColorText.ToHex(named.Color));
            }

            swatches.Add(swatch, named);
            controls.Add(swatch);
        }

        Child = CreateLayout(controls);
        OnColorSet(Color);
    }

    void Swatch_PointerPressed(object sender, PointerPressedEventArgs e)
    {
        var swatch = (Control)sender;
        Select(swatch);
        Pick(swatches[swatch].Color);
    }

    /// <summary>The visual of one swatch. Whatever it returns gets the click and the tooltip.</summary>
    protected virtual Control CreateSwatch(NamedColor named)
    {
        return new Border()
        {
            Width = SwatchSize,
            Height = SwatchSize,
            Margin = new Thickness(Spacing / 2),
            CornerRadius = new CornerRadius(SwatchCornerRadius),
            Background = named.Color.A == 0 ? ColorText.CheckerboardBrush : new SolidColorBrush(named.Color),
            Cursor = new Cursor(StandardCursorType.Hand)
        };
    }

    /// <summary>Marks or unmarks a swatch created by <see cref="CreateSwatch"/> as the current color.</summary>
    protected virtual void ShowSelected(Control swatch, NamedColor named, bool isSelected)
    {
        var border = (Border)swatch;

        // a dark ring with a white one inside it: reads on any swatch color and on any neighbor
        border.BorderBrush = isSelected ? Brushes.Black : null;
        border.BorderThickness = new Thickness(isSelected ? 2 : 0);
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
