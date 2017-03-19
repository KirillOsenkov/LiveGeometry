using System.Windows;
using System.Windows.Controls;

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

        void StringPropertyEditor_TextChanged(object sender, TextChangedEventArgs e)
        {
            SetValue(TextBox.Text);
        }

        public override void UpdateEditor()
        {
            TextBox.Text = (GetValue() ?? "").ToString();
            TextBox.IsReadOnly = !Value.CanSetValue;
        }
    }
}
