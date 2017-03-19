using System;
using System.Windows;
using System.Windows.Controls;

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
            Slider.ValueChanged += Slider_ValueChanged;

            TextBox = new TextBox();
            TextBox.MinWidth = 40;
            TextBox.TextChanged += TextBox_TextChanged;

            Panel = new Grid();
            Panel.ColumnDefinitions.Add(new ColumnDefinition() { Width = GridLength.Auto });
            Panel.ColumnDefinitions.Add(new ColumnDefinition());
            Panel.Children.Add(TextBox);
            Grid.SetColumn(Slider, 1);
            Panel.Children.Add(Slider);
            Panel.HorizontalAlignment = HorizontalAlignment.Stretch;
            return Panel;
        }

        bool guard = false;
        void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (guard)
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

        void Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
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
            TextBox.Text = value.Round(1).ToStringInvariant();
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
            TextBox.Text = value.Round(1).ToStringInvariant();
            guard = false;
        }
    }
}
