using System.Globalization;
using Avalonia.Controls;
using Avalonia.Layout;

namespace DynamicGeometry;

/// <summary>
/// Picked by <c>[PropertyGridPreferredEditor("UpDown")]</c> on a double: a number box with
/// up/down buttons (also the Up/Down keys and the wheel) and no slider, for a value with no
/// natural range, such as a length.
/// </summary>
public class UpDownEditorFactory : BaseValueEditorFactory<UpDownEditor, double>
{
    public UpDownEditorFactory()
    {
        // after the plain double editor: only the attribute chooses this one
        LoadOrder = 3;
    }
}

public class UpDownEditor : LabeledValueEditor, IValueEditor
{
    public const double StepSize = 1;

    public TextBox TextBox { get; set; }
    public UpDownControl UpDown { get; set; }

    protected override UIElement CreateEditor()
    {
        TextBox = new TextBox();
        TextBox.Width = 80;
        TextBox.VerticalAlignment = VerticalAlignment.Center;
        TextBox.CornerRadius = new Avalonia.CornerRadius(4, 0, 0, 4);
        TextBox.TextChanged += TextBox_TextChanged;

        UpDown = new UpDownControl();
        UpDown.VerticalAlignment = VerticalAlignment.Center;
        TextBox.SizeChanged += (s, e) => UpDown.Height = TextBox.Bounds.Height;
        UpDown.Up += () => StepValue(up: true);
        UpDown.Down += () => StepValue(up: false);
        UpDown.AttachTo(TextBox);

        var panel = new StackPanel() { Orientation = Orientation.Horizontal };
        panel.Children.Add(TextBox);
        panel.Children.Add(UpDown);
        return panel;
    }

    protected override void Focus()
    {
        TextBox.Focus();
        TextBox.SelectAll();
    }

    void StepValue(bool up)
    {
        if (Value == null || !Value.CanSetValue)
        {
            return;
        }

        var next = UpDownControl.Step(GetValue<double>(), up, StepSize, double.MinValue, double.MaxValue);
        SetValue(next);
        ShowValue(next);
    }

    bool guard;

    void TextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (guard)
        {
            return;
        }

        if (double.TryParse(TextBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var result))
        {
            SetValue(result);
        }
    }

    void ShowValue(double value)
    {
        guard = true;
        TextBox.Text = System.Math.Round(value, 4).ToStringInvariant();
        guard = false;
    }

    public override void UpdateEditor()
    {
        ShowValue(GetValue<double>());
        TextBox.IsEnabled = Value.CanSetValue;
    }
}
