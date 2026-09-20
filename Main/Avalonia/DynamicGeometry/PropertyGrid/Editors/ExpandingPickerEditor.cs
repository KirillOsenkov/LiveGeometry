using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// A property grid row whose editor is too big to live in the row: the row shows a chip
/// with the current value, and clicking the chip unfolds the actual picker underneath,
/// across the full width of the grid. Only one picker is unfolded at a time.
/// </summary>
public abstract class ExpandingPickerEditor : LabeledValueEditor
{
    static ExpandingPickerEditor expanded;

    readonly Border chipFill = new Border();
    Control picker;

    protected override UIElement CreateEditor()
    {
        var chip = new Border()
        {
            Width = 56,
            Height = 22,
            Margin = new Thickness(0, 3, 0, 3),
            HorizontalAlignment = HorizontalAlignment.Left,
            Background = ColorText.CheckerboardBrush,
            BorderBrush = RibbonTheme.TabLine,
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(4),
            ClipToBounds = true,
            Cursor = new Cursor(StandardCursorType.Hand),
            Child = chipFill
        };
        chip.PointerPressed += (s, e) => IsExpanded = !IsExpanded;

        RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
        RowDefinitions.Add(new RowDefinition() { Height = GridLength.Auto });
        return chip;
    }

    /// <summary>The big editor, created the first time the row is unfolded.</summary>
    protected abstract Control CreatePicker();

    /// <summary>Push the current value into the picker created by <see cref="CreatePicker"/>.</summary>
    protected abstract void UpdatePicker();

    /// <summary>What the chip shows</summary>
    protected abstract IBrush GetChipBrush();

    public bool IsExpanded
    {
        get => picker != null && picker.IsVisible;
        set
        {
            if (value && (Value == null || !Value.CanSetValue))
            {
                return;
            }

            if (value && picker == null)
            {
                picker = CreatePicker();
                picker.Margin = new Thickness(0, 4, 0, 8);
                Grid.SetRow(picker, 1);
                Grid.SetColumnSpan(picker, 2);
                Children.Add(picker);
            }

            if (picker == null)
            {
                return;
            }

            if (value)
            {
                if (expanded != null && expanded != this)
                {
                    expanded.IsExpanded = false;
                }

                expanded = this;
                UpdatePicker();
            }

            picker.IsVisible = value;
        }
    }

    public override void UpdateEditor()
    {
        chipFill.Background = GetChipBrush();
        if (IsExpanded)
        {
            UpdatePicker();
        }
    }

    /// <summary>To be called by the subclass when the user changes the value in the picker.</summary>
    protected void Commit(object newValue)
    {
        SetValue(newValue);
        chipFill.Background = GetChipBrush();
    }
}
