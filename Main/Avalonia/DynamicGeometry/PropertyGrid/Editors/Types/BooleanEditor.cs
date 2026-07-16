using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Interactivity;
using Avalonia.Controls;

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
            CheckBox.IsCheckedChanged += CheckBox_CheckedChanged;
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
