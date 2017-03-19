using System.Windows;
using System.Windows.Media;
using GuiLabs.Controls;
using SilverlightContrib.Controls;

namespace DynamicGeometry
{
    public class ColorEditorFactory 
        : BaseValueEditorFactory<ColorEditor, Color> { }

    public class ColorEditor : LabeledValueEditor, IValueEditor
    {
        public ColorPicker Picker { get; set; }

        protected override UIElement CreateEditor()
        {
            Picker = new ColorPicker();
            Picker.SelectedColorChanging += ColorChanged;
            Picker.VerticalAlignment = VerticalAlignment.Top;
            return Picker;
        }

        void ColorChanged(object sender, SelectedColorEventArgs e)
        {
            SetValue(e.SelectedColor);
        }

        public override void UpdateEditor()
        {
            Picker.SelectedColor = GetValue<Color>();
            Picker.IsHitTestVisible = Value.CanSetValue;
        }
    }
}
