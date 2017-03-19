using System.Windows;
using System.Windows.Controls;

namespace DynamicGeometry
{
    public class BooleanEditorFactory 
        : BaseValueEditorFactory<BooleanEditor, bool> {}

    public class BooleanEditor : LabeledValueEditor, IValueEditor
    {
        public CheckBox CheckBox { get; set; }

        protected override UIElement CreateEditor()
        {
            CheckBox = new CheckBox();
            CheckBox.VerticalAlignment = VerticalAlignment.Center;
            CheckBox.Checked += CheckBox_CheckedChanged;
            CheckBox.Unchecked += CheckBox_CheckedChanged;
            return CheckBox;
        }

        void CheckBox_CheckedChanged(object sender, RoutedEventArgs e)
        {
            SetValue(CheckBox.IsChecked ?? true);
        }

        public override void UpdateEditor()
        {
            CheckBox.IsChecked = GetValue<bool>();
            CheckBox.IsEnabled = Value.CanSetValue;
        }
    }
}
