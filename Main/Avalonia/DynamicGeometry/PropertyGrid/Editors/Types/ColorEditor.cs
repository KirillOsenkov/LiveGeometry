using Avalonia.Controls;
using Avalonia.Media;

namespace DynamicGeometry
{
    public class ColorEditorFactory
        : BaseValueEditorFactory<ColorEditor, Color> { }

    public class ColorEditor : ExpandingPickerEditor, IValueEditor
    {
        public ColorPickerView Picker { get; private set; }

        protected override Control CreatePicker()
        {
            Picker = new ColorPickerView();
            Picker.ColorChanged += color => Commit(color);
            return Picker;
        }

        protected override void SetPickerSurface(IBrush surface)
        {
            Picker.Surface = surface;
        }

        protected override void UpdatePicker()
        {
            Picker.Color = GetValue<Color>();
        }

        protected override IBrush GetChipBrush()
        {
            return new SolidColorBrush(GetValue<Color>());
        }
    }
}
