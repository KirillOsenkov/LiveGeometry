using System.Windows;
using System.Windows.Controls;

namespace DynamicGeometry
{
    public class ReadOnlyValuePropertyEditorFactory 
        : BaseValueEditorFactory<ReadOnlyValuePropertyEditor>
    {
        public ReadOnlyValuePropertyEditorFactory()
        {
            LoadOrder = 4;
        }

        public override bool SupportsValue(IValueProvider property)
        {
            return !property.CanSetValue;
        }
    }

    public class ReadOnlyValuePropertyEditor : LabeledValueEditor, IValueEditor
    {
        public TextBox TextBox { get; set; }

        protected override UIElement CreateEditor()
        {
            TextBox = new TextBox();
            TextBox.VerticalAlignment = VerticalAlignment.Center;
            TextBox.IsReadOnly = true;
            return TextBox;
        }

        public override void UpdateEditor()
        {
            TextBox.Text = (GetValue() ?? "").ToString();
        }
    }
}
