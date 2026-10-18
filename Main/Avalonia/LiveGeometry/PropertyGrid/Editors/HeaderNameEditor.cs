using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace LiveGeometry;

/// <summary>
/// The figure's name in the header of its page, in the title's place
/// (<see cref="PropertyGridSection.Header"/>), after what the figure is: a
/// <see cref="NameEditor"/> that looks like the title - no box until the cursor is over it or
/// it has the keyboard, which the theme's states put on, as wide as the name - and shows the
/// name as it is drawn (A₁) while it isn't being edited, and as it is typed (A_1) while it is
/// (<see cref="NameDisplay"/>).
/// </summary>
public class HeaderNameEditor : NameEditor
{
    /// <summary>The size of a page's title, the name's included</summary>
    public const double TitleFontSize = 15;

    // the box's border and padding: its text starts this far in, and the box reaches out by
    // as much, so that the name keeps the gap after the kind of figure before it
    const double Inset = 7;

    // a short name (A) still has a box to click into
    const double MinimumWidth = 60;

    // how far the name goes below the middle of its box, to share the kind's baseline
    const double BaselineNudge = 0.5;

    // the name as drawn, over the box while it isn't being edited
    TextBlock display;

    public HeaderNameEditor()
    {
        // no caption: the header says what the box is
        Label.IsVisible = false;
        ColumnDefinitions[0].MinWidth = 0;
        ColumnDefinitions[0].SharedSizeGroup = null;

        // as wide as the name, and wider as it grows; in the middle of the line, as the kind
        // before it is (the toggles after it make the line taller than the box)
        HorizontalAlignment = HorizontalAlignment.Left;
        VerticalAlignment = VerticalAlignment.Center;
    }

    protected override UIElement CreateEditor()
    {
        CreateTextBox();
        TextBox.FontSize = TitleFontSize;
        TextBox.FontWeight = FontWeight.SemiBold;
        TextBox.MinHeight = 0;
        TextBox.MinWidth = MinimumWidth;
        // half a pixel lower than centered (a pixel on a screen at 200%): the box's text sat
        // that much over the baseline of the kind before it
        TextBox.Padding = new Thickness(Inset - 1, 2 + BaselineNudge, Inset - 1, 2 - BaselineNudge);
        TextBox.Margin = new Thickness(-Inset, 0, 0, 0);

        // the theme styles its template on hover and focus, under these: a box shows then
        TextBox.Background = Brushes.Transparent;
        TextBox.BorderBrush = Brushes.Transparent;

        // selected as the grid's boxes are (PropertyGridTheme): the header is outside the
        // grid, where its styles don't reach
        TextBox.BindTheme(TextBox.SelectionBrushProperty, nameof(AppTheme.ButtonChecked));
        TextBox.BindTheme(TextBox.SelectionForegroundBrushProperty, nameof(AppTheme.Text));
        TextBox.GotFocus += (s, e) => UpdateDisplay();
        TextBox.LostFocus += (s, e) => UpdateDisplay();

        display = new TextBlock()
        {
            FontSize = TitleFontSize,
            FontWeight = FontWeight.SemiBold,
            VerticalAlignment = VerticalAlignment.Center,
            // centered, so twice the nudge: the same place as the box's text
            Margin = new Thickness(0, 2 * BaselineNudge, 0, 0),
            IsHitTestVisible = false,
            IsVisible = false
        };
        display.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Text));
        var cell = new Grid() { Children = { TextBox, display } };
        return WithErrorBox(cell);
    }

    public override void UpdateEditor()
    {
        base.UpdateEditor();
        UpdateDisplay();
    }

    protected override void OnBoxTextChanged(string text)
    {
        base.OnBoxTextChanged(text);
        UpdateDisplay();
    }

    void UpdateDisplay()
    {
        if (display == null)
        {
            return;
        }

        string text = TextBox.Text ?? "";
        string shown = NameDisplay.Format(text);
        bool showsDrawn = !TextBox.IsFocused && shown != text;
        display.Text = shown;
        display.IsVisible = showsDrawn;
        if (showsDrawn)
        {
            TextBox.Foreground = Brushes.Transparent;
        }
        else
        {
            TextBox.ClearValue(TextBox.ForegroundProperty);
        }
    }
}
