using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;

namespace LiveGeometry;

public class StringEditorFactory
    : BaseValueEditorFactory<StringEditor, string> { }

public class StringEditor : LabeledValueEditor, IValueEditor
{
    public TextBox TextBox { get; set; }

    protected override UIElement CreateEditor()
    {
        CreateTextBox();
        return WithErrorBox(TextBox);
    }

    protected void CreateTextBox()
    {
        TextBox = CreateTextBoxCore();
        TextBox.TextChanged += StringPropertyEditor_TextChanged;
        TextBox.KeyDown += TextBox_KeyDown;
        TextBox.LostFocus += TextBox_LostFocus;
    }

    /// <summary>The box itself: a plain one, or an <see cref="ExpressionTextBox"/> for a row that holds an expression</summary>
    protected virtual TextBox CreateTextBoxCore()
    {
        return new TextBox();
    }

    protected override void Focus()
    {
        TextBox.Focus();
        if (!string.IsNullOrEmpty(TextBox.Text))
        {
            TextBox.SelectAll();
        }
    }

    /// <summary>
    /// The text the editor itself put in the box, until the user edits it (then null).
    /// Avalonia raises TextChanged through the dispatcher, after the guard around
    /// UpdateEditor is gone, so that is how a programmatic update is told from the user's
    /// typing - and it matters once the box shows a rounded number: parsing it back would
    /// change the value. Only until the first edit, though: typing a text and deleting it
    /// again is the user's empty text, and must reach the property.
    /// </summary>
    protected string ShownText { get; private set; }

    protected void Show(string text)
    {
        ShownText = text;
        appliedText = null;
        TextBox.Text = text;
    }

    // The text of the user's that has gone into the property already. A commit (Enter,
    // then leaving the box) doesn't set it a second and a third time: the value read
    // back need not equal what was typed (a vector's direction is worked out from its
    // points, 27.57 comes back as 27.570000000000004), and each such set was an undo
    // step of its own that undid nothing.
    string appliedText;

    void StringPropertyEditor_TextChanged(object sender, TextChangedEventArgs e)
    {
        TextEdited();
    }

    // what the box says now goes into the property; again on commit, in case the
    // TextChanged is still on its way
    void TextEdited()
    {
        string text = TextBox.Text ?? "";
        if (text != errorForText)
        {
            ErrorText = null;
        }

        OnBoxTextChanged(text);
        if (text == ShownText)
        {
            return;
        }

        ShownText = null;
        if (text == appliedText)
        {
            return;
        }

        appliedText = text;
        SetValue(text);
    }

    /// <summary>
    /// The box says this now, the editor's own text or the user's: an expression editor
    /// shows the value under the box (<see cref="ValueText"/>)
    /// </summary>
    protected virtual void OnBoxTextChanged(string text)
    {
    }

    #region Commit and errors

    // Enter in a one-line box, leaving the box
    void TextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter && !TextBox.AcceptsReturn)
        {
            // not handled: a tool's panel acts on Enter too (Add point), after this
            Commit();
            OnEnter();
        }
    }

    /// <summary>Enter in a one-line box, after the commit: an expression editor of a figure's row is done with the keyboard</summary>
    protected virtual void OnEnter()
    {
    }

    /// <summary>What the text is worth, worked out anew: the figures it names have moved (<see cref="PropertyGrid.RefreshNumbers"/>)</summary>
    public void RefreshLiveValue()
    {
        OnBoxTextChanged(TextBox.Text ?? "");
    }

    void TextBox_LostFocus(object sender, RoutedEventArgs e)
    {
        Commit();
    }

    /// <summary>
    /// The user is done with the text: it goes into the property even if its TextChanged
    /// is still on its way (a panel's Enter acts on the property right after this), and
    /// what is wrong with it shows under the box. Errors don't show while typing.
    /// </summary>
    void Commit()
    {
        TextEdited();

        // what is typed after this is another undo step
        EndEditRun();

        // what the editor showed needs no checking (an empty box never touched, say)
        string text = TextBox.Text ?? "";
        if (text == ShownText)
        {
            return;
        }

        var validation = Validate(text);
        if (!validation.IsValid && !string.IsNullOrEmpty(validation.Error))
        {
            ErrorText = validation.Error;
        }
    }

    Border errorBox;
    TextBlock errorBlock;
    Border valueBox;
    TextBlock valueBlock;
    CornerRadius? cornersWithoutError;

    // the text the error is about: editing it takes the error away
    string errorForText;

    /// <summary>
    /// What the text is worth right now, in a plate under the box ("= 18.85"), as VB6's
    /// calculator said it while one typed; null or empty hides it. An error takes its place.
    /// </summary>
    public string ValueText
    {
        get => valueBlock?.Text;
        set
        {
            if (valueBox == null)
            {
                return;
            }

            bool show = !string.IsNullOrEmpty(value);
            valueBlock.Text = show ? value : null;
            UpdatePlates(show);
        }
    }

    // one plate at a time: the error over the value, and square bottom corners on the text
    // box while either shows, so the two read as one piece
    void UpdatePlates(bool valueWanted)
    {
        bool error = !string.IsNullOrEmpty(errorBlock.Text);
        bool showValue = valueWanted && !error;
        errorBox.IsVisible = error;
        valueBox.IsVisible = showValue;
        bool show = error || showValue;
        if (show && cornersWithoutError == null)
        {
            var corners = TextBox.CornerRadius;
            cornersWithoutError = corners;
            TextBox.CornerRadius = new CornerRadius(corners.TopLeft, corners.TopRight, 0, 0);
        }
        else if (!show && cornersWithoutError != null)
        {
            TextBox.CornerRadius = cornersWithoutError.Value;
            cornersWithoutError = null;
        }
    }

    /// <summary>
    /// What is wrong with the text, in a box attached under the text box (it can wrap and
    /// make the row wider); null or empty hides it. Set by a commit or by a tool's panel
    /// (<see cref="ToolPanel"/>), and taken away when the user edits the text.
    /// </summary>
    public string ErrorText
    {
        get => errorBlock?.Text;
        set
        {
            if (errorBox == null)
            {
                return;
            }

            bool show = !string.IsNullOrEmpty(value);
            errorBlock.Text = show ? value : null;
            errorForText = show ? TextBox.Text : null;
            UpdatePlates(!string.IsNullOrEmpty(valueBlock.Text));
        }
    }

    /// <summary>The editor with room for <see cref="ErrorText"/> and <see cref="ValueText"/> under it</summary>
    protected UIElement WithErrorBox(Control editor)
    {
        errorBlock = new TextBlock()
        {
            TextWrapping = TextWrapping.Wrap
        };
        errorBlock.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.ErrorText));
        errorBox = CreatePlate(errorBlock);
        errorBox.BindTheme(Border.BackgroundProperty, nameof(AppTheme.ErrorBackground));
        errorBox.BindTheme(Border.BorderBrushProperty, nameof(AppTheme.ErrorBorder));

        valueBlock = new TextBlock()
        {
            TextWrapping = TextWrapping.Wrap
        };
        valueBlock.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.TextMuted));
        valueBox = CreatePlate(valueBlock);
        valueBox.BindTheme(Border.BackgroundProperty, nameof(AppTheme.InputBackground));
        valueBox.BindTheme(Border.BorderBrushProperty, nameof(AppTheme.Separator));

        var panel = new StackPanel();
        panel.Children.Add(editor);
        panel.Children.Add(errorBox);
        panel.Children.Add(valueBox);
        return panel;
    }

    /// <summary>A plate attached under the text box, hidden until it has something to say</summary>
    static Border CreatePlate(TextBlock block)
    {
        return new Border()
        {
            BorderThickness = new Thickness(1),
            CornerRadius = new CornerRadius(0, 0, 4, 4),
            Padding = new Thickness(6, 3, 6, 4),
            // flush with the text box: over its bottom margin and its bottom border line
            Margin = new Thickness(0, -3, 0, 2),
            // as wide as a text box may get (PropertyGridTheme), then it wraps
            MaxWidth = PropertyGridTheme.MaxTextWidth,
            IsVisible = false,
            Child = block
        };
    }

    #endregion

    public override void UpdateEditor()
    {
        // one line unless the property asks for more: a line break in an expression (a
        // function, a coordinate) only breaks it, and Enter belongs to the panel's button
        TextBox.AcceptsReturn = Value.GetAttribute<PropertyGridMultilineAttribute>() != null;
        Show((GetValue() ?? "").ToString());
        // grayed, like the other editors: a read-only box looks the same as a live one
        TextBox.IsEnabled = Value.CanSetValue;
    }
}
