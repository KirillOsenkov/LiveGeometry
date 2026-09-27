namespace DynamicGeometry
{
    public class DrawingExpressionEditorFactory
        : BaseValueEditorFactory<ExpressionEditor, DrawingExpression> { }

    /// <summary>
    /// An expression of a figure (a coordinate of a point by coordinates): each keystroke that
    /// makes a valid expression applies it; what is wrong shows under the box on commit
    /// </summary>
    public class ExpressionEditor : StringEditor
    {
        protected override ValidationResult Validate(object value)
        {
            string source = value.ToString();
            DrawingExpression expression = Value as DrawingExpression;
            var compileResult = Compiler.Instance.CompileExpression(
                expression.ParentFigure.Drawing,
                source,
                f => !f.DependsOn(expression.ParentFigure));

            return new ValidationResult()
            {
                IsValid = compileResult.IsSuccess,
                Value = source,
                Error = compileResult.GetErrorText()
            };
        }
    }
}
