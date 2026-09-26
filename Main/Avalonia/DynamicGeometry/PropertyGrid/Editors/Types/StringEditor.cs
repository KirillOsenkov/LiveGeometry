using Avalonia.Controls;

namespace DynamicGeometry
{
    public class StringEditorFactory 
        : BaseValueEditorFactory<StringEditor, string> { }

    public class StringEditor : LabeledValueEditor, IValueEditor
    {
        public TextBox TextBox { get; set; }

        protected override UIElement CreateEditor()
        {
            TextBox = new TextBox();
            TextBox.TextChanged += StringPropertyEditor_TextChanged;
            TextBox.AcceptsReturn = true;
            return TextBox;
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
        /// The text the editor itself put in the box last. Avalonia raises TextChanged through
        /// the dispatcher, after the guard around UpdateEditor is gone, so that is how a
        /// programmatic update is told from the user's typing - and it matters once the box
        /// shows a rounded number: parsing it back would change the value.
        /// </summary>
        protected string ShownText { get; private set; }

        protected void Show(string text)
        {
            ShownText = text;
            TextBox.Text = text;
        }

        void StringPropertyEditor_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (TextBox.Text == ShownText)
            {
                return;
            }

            SetValue(TextBox.Text);
        }

        public override void UpdateEditor()
        {
            Show((GetValue() ?? "").ToString());
            // grayed, like the other editors: a read-only box looks the same as a live one
            TextBox.IsEnabled = Value.CanSetValue;
        }
    }
}
