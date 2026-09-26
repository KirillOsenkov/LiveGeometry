using Avalonia.Controls;
using Avalonia.Layout;

namespace DynamicGeometry
{
    public class DomainDoubleEditorFactory : BaseValueEditorFactory<SliderEditor, double>
    {
        public override bool SupportsValue(IValueProvider value)
        {
            return base.SupportsValue(value) && value.GetAttribute<DomainAttribute>() != null;
        }
    }

    public class SliderEditor : LabeledValueEditor, IValueEditor
    {
        public Slider Slider { get; set; }
        public TextBox TextBox { get; set; }
        public Grid Panel { get; set; }

        protected override UIElement CreateEditor()
        {
            Slider = new Slider();
            Slider.VerticalAlignment = VerticalAlignment.Center;
            Slider.MinWidth = 110;
            Slider.Margin = new Avalonia.Thickness(8, 0, 0, 0);
            Slider.ValueChanged += Slider_ValueChanged;

            TextBox = new TextBox();
            // fixed, so that the row doesn't shift as the number gets longer or shorter
            TextBox.Width = 52;
            TextBox.VerticalAlignment = VerticalAlignment.Center;
            // square on the right, where the up/down buttons are attached
            TextBox.CornerRadius = new Avalonia.CornerRadius(4, 0, 0, 4);
            TextBox.TextChanged += TextBox_TextChanged;

            // the slider is for sweeping through the range, these are for exact single steps
            UpDown = new UpDownControl();
            UpDown.VerticalAlignment = VerticalAlignment.Center;
            TextBox.SizeChanged += (s, e) => UpDown.Height = TextBox.Bounds.Height;
            UpDown.Up += () => StepValue(up: true);
            UpDown.Down += () => StepValue(up: false);
            UpDown.AttachTo(TextBox);

            Panel = new Grid();
            Panel.ColumnDefinitions.Add(new ColumnDefinition() { Width = GridLength.Auto });
            Panel.ColumnDefinitions.Add(new ColumnDefinition() { Width = GridLength.Auto });
            Panel.ColumnDefinitions.Add(new ColumnDefinition());
            Panel.Children.Add(TextBox);
            Grid.SetColumn(UpDown, 1);
            Panel.Children.Add(UpDown);
            Grid.SetColumn(Slider, 2);
            Panel.Children.Add(Slider);
            Panel.HorizontalAlignment = HorizontalAlignment.Stretch;
            return Panel;
        }

        public UpDownControl UpDown { get; set; }

        void StepValue(bool up)
        {
            if (Value == null || !Value.CanSetValue)
            {
                return;
            }

            // whole numbers, unless the whole range is only a few units wide
            double step = Slider.Maximum - Slider.Minimum > 10 ? 1 : 0.1;
            Slider.Value = UpDownControl.Step(Slider.Value, up, step, Slider.Minimum, Slider.Maximum);
        }

        bool guard = false;

        // what the editor itself put in the box last: TextChanged arrives late, through the
        // dispatcher, and parsing a rounded display back would change the value
        string shownText;

        void Show(double value)
        {
            shownText = value.Round(Settings.DisplayDecimals).ToStringInvariant();
            TextBox.Text = shownText;
        }

        void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (guard || TextBox.Text == shownText)
            {
                return;
            }

            string source = TextBox.Text;
            double result;
            if (!string.IsNullOrEmpty(source) 
                && double.TryParse(source, out result)
                && result >= Slider.Minimum
                && result <= Slider.Maximum)
            {
                guard = true;
                Slider.Value = result;
                SetValue(result);
                guard = false;
            };
        }

        void Slider_ValueChanged(object sender, Avalonia.Controls.Primitives.RangeBaseValueChangedEventArgs e)
        {
            if (guard)
            {
                return;
            }

            var value = Slider.Value;
            if (Slider.Maximum - Slider.Minimum > 50 && value > Slider.Minimum && value < Slider.Maximum)
            {
                value = value.Round(0);
            }

            guard = true;
            Show(value);
            SetValue((object)value);
            guard = false;
        }

        protected override void InitCore()
        {
            var attribute = Value.GetAttribute<DomainAttribute>();
            Slider.Minimum = attribute.MinValue;
            Slider.Maximum = attribute.MaxValue;
        }

        public override void UpdateEditor()
        {
            guard = true;
            var value = GetValue<double>();
            Slider.Value = value;
            Slider.IsEnabled = Value.CanSetValue;
            TextBox.IsEnabled = Slider.IsEnabled;
            Show(value);
            guard = false;
        }
    }
}
