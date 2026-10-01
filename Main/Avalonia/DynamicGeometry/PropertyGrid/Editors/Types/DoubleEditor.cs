namespace DynamicGeometry
{
    public class DoubleEditorFactory : BaseValueEditorFactory<DoubleEditor, double>
    {
        public DoubleEditorFactory()
        {
            LoadOrder = 2;
        }
    }

    public class DoubleEditor : StringEditor
    {
        /// <summary>Shown rounded (Settings.DisplayDecimals); the value keeps its digits</summary>
        public override void UpdateEditor()
        {
            var value = GetValue();
            Show(value is double number ? Math.Round(number, Settings.DisplayDecimals).ToStringInvariant() : (value ?? "").ToString());
            TextBox.IsEnabled = Value.CanSetValue;
        }

        protected override ValidationResult Validate(object value)
        {
            var result = new ValidationResult();
            double doubleResult;
            string source = value.ToString();
            // a number there is: "Infinity" and 1e999 parse too, and a figure at infinity is nowhere
            if (!string.IsNullOrEmpty(source) && double.TryParse(source, out doubleResult) && doubleResult.IsValidValue())
            {
                result.IsValid = true;
                result.Value = doubleResult;
            }
            else
            {
                result.Error = "Type a number.";
            }
            return result;
        }
    }
}
