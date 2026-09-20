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
            base.CreateEditor();
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
            return panel;
        }

        void StepValue(bool up)
        {
            if (Value == null || !Value.CanSetValue)
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
            if (!string.IsNullOrEmpty(source) && int.TryParse(source, out intValue))
            {
                result.IsValid = true;
                result.Value = intValue;
            };
            return result;
        }
    }
}
