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

        // a value typed or stepped to, then another after Enter or after coming back to the
        // box: two undo steps
        TextBox.LostFocus += (s, e) => EndEditRun();
        TextBox.KeyDown += (s, e) =>
        {
            if (e.Key == Avalonia.Input.Key.Enter)
            {
                EndEditRun();
            }
        };

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

    // what the editor itself put in the box last: TextChanged arrives late, through the
    // dispatcher, and parsing a rounded display back would change the value
    string shownText;

    void TextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
        if (TextBox.Text == shownText)
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
        shownText = System.Math.Round(value, Settings.DisplayDecimals).ToStringInvariant();
        TextBox.Text = shownText;
    }

    public override void UpdateEditor()
    {
        ShowValue(GetValue<double>());
        bool canSet = Value.CanSetValue;
        TextBox.IsEnabled = canSet;

        // a value that can't be set has nothing to step: no buttons, and the box gets its
        // right corners back
        UpDown.IsVisible = canSet;
        TextBox.CornerRadius = canSet ? new Avalonia.CornerRadius(4, 0, 0, 4) : new Avalonia.CornerRadius(4);
    }
}
