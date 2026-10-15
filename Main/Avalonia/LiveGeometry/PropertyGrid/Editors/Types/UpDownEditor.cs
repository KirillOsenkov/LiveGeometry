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
        TextBox.LostFocus += (s, e) =>
        {
            CommitText();
            EndEditRun();
        };
        TextBox.KeyDown += (s, e) =>
        {
            if (e.Key == Avalonia.Input.Key.Enter)
            {
                CommitText();
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
        // (no value: several figures that differ - there is nothing to step from)
        if (Value == null || !Value.CanSetValue || GetValue() == null)
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

        // The user's from here on: the text the box showed, typed again, is theirs too.
        // (The box said 12; Backspace made the length 1; the 2 typed back was taken for
        // the editor's own text and dropped - a segment 1 long under a box that said 12.)
        shownText = null;
        Apply(TextBox.Text);
    }

    // The typed text that has gone into the property already: it is not set again when the
    // box is left. The value read back need not equal it (a length is worked out from the
    // points: 5 comes back as 4.999999999999999), and a second set would be an undo step
    // of its own.
    string appliedText;

    bool Apply(string text)
    {
        if (!TryParse(text, out var result))
        {
            return false;
        }

        if (text != appliedText)
        {
            appliedText = text;
            SetValue(result);
        }

        return true;
    }

    // a number there is: "Infinity" and 1e999 parse too
    static bool TryParse(string text, out double result)
    {
        return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out result) && result.IsValidValue();
    }

    /// <summary>
    /// The user is done with the box (Enter, leaving it): text that is no number gives way
    /// to the value, which is what the box then says (it used to keep saying "abc")
    /// </summary>
    void CommitText()
    {
        if (Value == null || !Value.CanSetValue || TextBox.Text == shownText)
        {
            return;
        }

        // (a number is set here too: its TextChanged may still be on its way)
        if (!Apply(TextBox.Text))
        {
            ShowCurrentValue();
        }
    }

    void ShowValue(double value)
    {
        ShowText(Math.Round(value, Settings.DisplayDecimals).ToStringInvariant());
    }

    void ShowText(string text)
    {
        shownText = text;
        appliedText = null;
        TextBox.Text = shownText;
    }

    /// <summary>
    /// The value, or an empty box when there is none: several figures whose values
    /// differ (it said 0, which none of them had)
    /// </summary>
    void ShowCurrentValue()
    {
        if (GetValue() is double value)
        {
            ShowValue(value);
        }
        else
        {
            ShowText("");
        }
    }

    public override void UpdateEditor()
    {
        ShowCurrentValue();
        bool canSet = Value.CanSetValue;
        TextBox.IsEnabled = canSet;

        // a value that can't be set has nothing to step: no buttons, and the box gets its
        // right corners back
        UpDown.IsVisible = canSet;
        TextBox.CornerRadius = canSet ? new Avalonia.CornerRadius(4, 0, 0, 4) : new Avalonia.CornerRadius(4);
    }
}
