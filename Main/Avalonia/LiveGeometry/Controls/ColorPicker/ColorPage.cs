using System;
using Avalonia.Controls;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// One way of choosing a color: a grid of swatches, a continuous spectrum, whatever comes
/// next. A <see cref="ColorPickerView"/> hosts any number of pages and lets the user flip
/// between them; they all show and edit the same color.
/// To add a new kind of picker, derive from this and add it to <see cref="ColorPickerView.Pages"/>.
/// </summary>
public abstract class ColorPage : Decorator
{
    /// <summary>Caption of the page in the page switcher</summary>
    public abstract string Title { get; }

    Color color = Colors.Black;

    /// <summary>
    /// The color the page shows. Setting it from outside updates the visuals and does not
    /// raise <see cref="ColorChanged"/>.
    /// </summary>
    public Color Color
    {
        get => color;
        set
        {
            if (color == value)
            {
                return;
            }

            color = value;
            OnColorSet(value);
        }
    }

    /// <summary>The user picked a color on this page.</summary>
    public event Action<Color> ColorChanged;

    /// <summary>Bring the visuals (selection, markers) in line with the color.</summary>
    protected abstract void OnColorSet(Color newColor);

    /// <summary>To be called by the page when the user picks a color.</summary>
    protected void Pick(Color picked)
    {
        color = picked;
        ColorChanged?.Invoke(picked);
    }
}
