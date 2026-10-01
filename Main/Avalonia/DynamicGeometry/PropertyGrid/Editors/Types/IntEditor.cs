using Avalonia.Controls;
using Avalonia.Layout;

namespace DynamicGeometry
{
    public class IntEditorFactory : BaseValueEditorFactory<IntEditor, int> { }

    public class IntEditor : StringEditor
    {
        public UpDownControl UpDown { get; set; }

        protected override UIElement CreateEditor()
        {
            CreateTextBox();
            TextBox.AcceptsReturn = false;
            TextBox.Width = 52;
            TextBox.VerticalAlignment = VerticalAlignment.Center;
            // square on the right, where the up/down buttons are attached
            TextBox.CornerRadius = new Avalonia.CornerRadius(4, 0, 0, 4);

            UpDown = new UpDownControl();
            UpDown.VerticalAlignment = VerticalAlignment.Center;
            UpDown.Up += () => StepValue(up: true);
            UpDown.Down += () => StepValue(up: false);
            UpDown.AttachTo(TextBox);
            TextBox.SizeChanged += (s, e) => UpDown.Height = TextBox.Bounds.Height;

            var panel = new StackPanel() { Orientation = Orientation.Horizontal };
            panel.Children.Add(TextBox);
            panel.Children.Add(UpDown);

            // with room under it for what is wrong with the number typed
            return WithErrorBox(panel);
        }

        void StepValue(bool up)
        {
            // (no value: several figures that differ - there is nothing to step from)
            if (Value == null || !Value.CanSetValue || GetValue() == null)
            {
                return;
            }

            // limits come from [Domain(min, max)] on the property, if it has one
            var domain = Value.GetAttribute<DomainAttribute>();
            double minimum = domain != null ? domain.MinValue : int.MinValue;
            double maximum = domain != null ? domain.MaxValue : int.MaxValue;

            int next = (int)UpDownControl.Step(GetValue<int>(), up, step: 1, minimum, maximum);

            // Not through TextBox.Text: its TextChanged comes later, after the refresh below
            // would already have put the old number back.
            SetValue(next.ToString());

            // the property may have refused the number; show what it actually holds
            UpdateEditor();
        }

        protected override ValidationResult Validate(object value)
        {
            ValidationResult result = new ValidationResult();
            string source = value.ToString();
            int intValue = 0;

            // within the limits of the property, if it has any: a number outside them was
            // dropped without a word, and the box went on showing it
            var domain = Value?.GetAttribute<DomainAttribute>();
            if (string.IsNullOrEmpty(source) || !int.TryParse(source, out intValue))
            {
                result.Error = "Type a whole number.";
            }
            else if (domain != null && (intValue < domain.MinValue || intValue > domain.MaxValue))
            {
                result.Error = "Type a whole number from " + domain.MinValue.ToStringInvariant() + " to " + domain.MaxValue.ToStringInvariant() + ".";
            }
            else
            {
                result.IsValid = true;
                result.Value = intValue;
            }

            return result;
        }
    }
}
