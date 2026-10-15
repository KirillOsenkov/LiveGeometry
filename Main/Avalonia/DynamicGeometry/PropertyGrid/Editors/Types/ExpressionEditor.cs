using System.Linq;

namespace DynamicGeometry;

public class DrawingExpressionEditorFactory
    : BaseValueEditorFactory<ExpressionEditor, DrawingExpression> { }

/// <summary>
/// An expression of a figure (a coordinate of a point by coordinates): each keystroke that
/// makes a valid expression applies it; what is wrong shows under the box on commit
/// </summary>
public class ExpressionEditor : StringEditor
{
    /// <summary>
    /// Valid when it compiles for every figure the row edits: one, or each of a selection
    /// (the X of two points), against its own figure. The row of a selection is the
    /// expressions together (<see cref="CompositeValueProvider"/>), which was taken for an
    /// expression, and typing into it threw.
    /// </summary>
    protected override ValidationResult Validate(object value)
    {
        string source = value.ToString();
        var expressions = Value is CompositeValueProvider composite
            ? composite.InnerList.OfType<DrawingExpression>()
            : new[] { Value as DrawingExpression };
        foreach (var expression in expressions.Where(expression => expression?.ParentFigure.Drawing != null))
        {
            var compileResult = Compiler.Instance.CompileExpression(
                expression.ParentFigure.Drawing,
                source,
                f => !f.DependsOn(expression.ParentFigure));
            if (!compileResult.IsSuccess)
            {
                return new ValidationResult()
                {
                    IsValid = false,
                    Value = source,
                    Error = compileResult.GetErrorText()
                };
            }
        }

        return new ValidationResult()
        {
            IsValid = true,
            Value = source
        };
    }
}
