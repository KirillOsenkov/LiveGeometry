using Avalonia.Controls;
using Avalonia.Media;

namespace DynamicGeometry
{
    public class BrushEditorFactory
        : BaseValueEditorFactory<BrushEditor, Brush> { }

    public class BrushEditor : ExpandingPickerEditor, IValueEditor
    {
        public BrushPickerView Picker { get; private set; }

        protected override Control CreatePicker()
        {
            Picker = new BrushPickerView();
            Picker.BrushChanged += brush => Commit(brush);
            return Picker;
        }

        protected override void SetPickerSurface(IBrush surface)
        {
            Picker.Surface = surface;
        }

        protected override void UpdatePicker()
        {
            Picker.Brush = GetValue<Brush>();
        }

        protected override IBrush GetChipBrush()
        {
            return GetValue<Brush>();
        }
    }
}
