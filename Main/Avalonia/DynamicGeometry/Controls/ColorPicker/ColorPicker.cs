using Avalonia.Media;

namespace SilverlightContrib.Controls
{
    /// <summary>
    /// Adapter with the SilverlightContrib ColorPicker API surface
    /// (SelectedColor + SelectedColorChanging/Changed) over the built-in
    /// Avalonia ColorPicker control.
    /// </summary>
    public class ColorPicker : Avalonia.Controls.ColorPicker
    {
        protected override System.Type StyleKeyOverride => typeof(Avalonia.Controls.ColorPicker);

        public event SelectedColorChangingHandler SelectedColorChanging;
        public event SelectedColorChangedHandler SelectedColorChanged;

        public ColorPicker()
        {
            ColorChanged += (s, e) =>
            {
                var args = new SelectedColorEventArgs(e.NewColor);
                SelectedColorChanging?.Invoke(this, args);
                SelectedColorChanged?.Invoke(this, args);
            };
        }

        public Color SelectedColor
        {
            get => Color;
            set => Color = value;
        }
    }
}
