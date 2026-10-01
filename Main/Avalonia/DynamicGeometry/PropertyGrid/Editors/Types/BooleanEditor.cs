using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Interactivity;

namespace DynamicGeometry
{
    public class BooleanEditorFactory 
        : BaseValueEditorFactory<BooleanEditor, bool> {}

    public class BooleanEditor : LabeledValueEditor, IValueEditor
    {
        public CheckBox CheckBox { get; set; }

        /// <summary>
        /// Each tick is an undo step of its own: merged, hiding and showing again would
        /// leave a step that undoes nothing
        /// </summary>
        protected override bool CoalescesEdits
        {
            get { return false; }
        }

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
