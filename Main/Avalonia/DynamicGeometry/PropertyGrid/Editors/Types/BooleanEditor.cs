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

        // several figures, some of them ticked: the box shows neither
        bool mixed;

        void CheckBox_CheckedChanged(object sender, RoutedEventArgs e)
        {
            // a click on the mixed box ticks them all (the check box itself goes from
            // "neither" to unticked)
            if (mixed && CheckBox.IsChecked == false)
            {
                mixed = false;
                CheckBox.IsChecked = true;
                return;
            }

            SetValue(CheckBox.IsChecked ?? true);
        }

        public override void UpdateEditor()
        {
            // no value: several figures that differ. It was shown unticked, as if none
            // of them were (hidden, locked...).
            var value = GetValue() as bool?;
            mixed = false;
            CheckBox.IsChecked = value;
            mixed = value == null;
            CheckBox.IsEnabled = Value.CanSetValue;
        }
    }
}
