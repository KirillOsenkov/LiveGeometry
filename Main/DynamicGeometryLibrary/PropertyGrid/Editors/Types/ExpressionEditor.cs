namespace DynamicGeometry
{
    public class DrawingExpressionEditorFactory
        : BaseValueEditorFactory<ExpressionEditor, DrawingExpression> { }

    public class ExpressionEditor : StringEditor
    {
        protected override ValidationResult Validate(object value)
        {
            ValidationResult result = new ValidationResult();
            string source = value.ToString();

            DrawingExpression expression = Value as DrawingExpression;

            if (!string.IsNullOrEmpty(source))
            {
                var compileResult = MEFHost.Instance.CompilerService.CompileExpression(
                    expression.ParentFigure.Drawing,
                    source,
                    f => !f.DependsOn(expression.ParentFigure));

                if (compileResult.IsSuccess)
                {
                    result.IsValid = true;
                    result.Value = source;
                    expression.ParentFigure.Drawing.ClearStatus();
                }
                else
                {
                    result.Error = compileResult.ToString();
                    expression.ParentFigure.Drawing.RaiseStatusNotification(result.Error);
                }
            };
            return result;
        }
    }
}
