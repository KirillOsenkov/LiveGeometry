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
        protected override ValidationResult Validate(object value)
        {
            var result = new ValidationResult();
            double doubleResult;
            string source = value.ToString();
            if (!string.IsNullOrEmpty(source) && double.TryParse(source, out doubleResult))
            {
                result.IsValid = true;
                result.Value = doubleResult;
            };
            return result;
        }
    }
}
