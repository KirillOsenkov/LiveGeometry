namespace DynamicGeometry
{
    public class IntEditorFactory : BaseValueEditorFactory<IntEditor, int> { }

    public class IntEditor : StringEditor
    {
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
