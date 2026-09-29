using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;

namespace DynamicGeometry
{
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
            TextBox = new TextBox();
            TextBox.TextChanged += StringPropertyEditor_TextChanged;
            TextBox.KeyDown += TextBox_KeyDown;
            TextBox.LostFocus += TextBox_LostFocus;
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
            TextBox.Text = text;
        }

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

            if (text == ShownText)
            {
                return;
            }

            ShownText = null;
            SetValue(text);
        }

        #region Commit and errors

        // Enter in a one-line box, leaving the box
        void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && !TextBox.AcceptsReturn)
            {
                // not handled: a tool's panel acts on Enter too (Add point), after this
                Commit();
            }
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
        CornerRadius? cornersWithoutError;

        // the text the error is about: editing it takes the error away
        string errorForText;

        /// <summary>
        /// What is wrong with the text, in a box attached under the text box (it can wrap and
        /// make the row wider); null or empty hides it. Set by a commit or by a tool's panel
        /// (<see cref="ToolPanel"/>), and taken away when the user edits the text.
        /// </summary>
        public string ErrorText
        {
            get
            {
                return errorBlock?.Text;
            }
            set
            {
                if (errorBox == null)
                {
                    return;
                }

                bool show = !string.IsNullOrEmpty(value);
                errorBlock.Text = show ? value : null;
                errorForText = show ? TextBox.Text : null;
                if (show == errorBox.IsVisible)
                {
                    return;
                }

                errorBox.IsVisible = show;

                // square bottom corners on the text box, so the two read as one piece
                if (show)
                {
                    var corners = TextBox.CornerRadius;
                    cornersWithoutError = corners;
                    TextBox.CornerRadius = new CornerRadius(corners.TopLeft, corners.TopRight, 0, 0);
                }
                else if (cornersWithoutError != null)
                {
                    TextBox.CornerRadius = cornersWithoutError.Value;
                    cornersWithoutError = null;
                }
            }
        }

        /// <summary>The editor with room for <see cref="ErrorText"/> under it</summary>
        protected UIElement WithErrorBox(Control editor)
        {
            errorBlock = new TextBlock()
            {
                TextWrapping = TextWrapping.Wrap
            };
            errorBlock.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.ErrorText));
            errorBox = new Border()
            {
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(0, 0, 4, 4),
                Padding = new Thickness(6, 3, 6, 4),
                // flush with the text box: over its bottom margin and its bottom border line
                Margin = new Thickness(0, -3, 0, 2),
                // as wide as a text box may get (PropertyGridTheme), then it wraps
                MaxWidth = 480,
                IsVisible = false,
                Child = errorBlock
            };
            errorBox.BindTheme(Border.BackgroundProperty, nameof(AppTheme.ErrorBackground));
            errorBox.BindTheme(Border.BorderBrushProperty, nameof(AppTheme.ErrorBorder));

            var panel = new StackPanel();
            panel.Children.Add(editor);
            panel.Children.Add(errorBox);
            return panel;
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
}
