using System.Windows;
using System.Windows.Media;
using GuiLabs.Controls;
using SilverlightContrib.Controls;

namespace DynamicGeometry
{
    public class BrushEditorFactory
        : BaseValueEditorFactory<BrushEditor, Brush> { }

    public class BrushEditor : LabeledValueEditor, IValueEditor
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
            SetValue(new SolidColorBrush(e.SelectedColor));
        }

        public override void UpdateEditor()
        {
            var brush = GetValue<Brush>();
            SolidColorBrush solidColorBrush = brush as SolidColorBrush;
            if (solidColorBrush != null)
            {
                Picker.SelectedColor = solidColorBrush.Color;
            }
            else
            {
                Picker.SelectedColor = Colors.Black;
            }

            Picker.IsHitTestVisible = Value.CanSetValue;
        }
    }
}