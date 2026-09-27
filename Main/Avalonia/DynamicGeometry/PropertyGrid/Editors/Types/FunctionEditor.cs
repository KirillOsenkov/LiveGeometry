namespace DynamicGeometry;

/// <summary>
/// Picked by <c>[PropertyGridPreferredEditor("Function")]</c> on a string: the expression in
/// x of a function graph. Like <see cref="ExpressionEditor"/>, a keystroke that makes a valid
/// function applies it, and what is wrong shows under the box on commit.
/// </summary>
public class FunctionEditorFactory : BaseValueEditorFactory<FunctionEditor, string>
{
    public FunctionEditorFactory()
    {
        // after the plain string editor: only the attribute chooses this one
        LoadOrder = 1;
    }
}

public class FunctionEditor : StringEditor
{
    protected override ValidationResult Validate(object value)
    {
        string source = value.ToString();
        var drawing = (Value.Parent as IFigure)?.Drawing;
        if (drawing == null)
        {
            return base.Validate(value);
        }

        var compileResult = Compiler.Instance.CompileFunction(drawing, source);
        return new ValidationResult()
        {
            IsValid = compileResult.IsSuccess,
            Value = source,
            Error = compileResult.GetErrorText(whenEmpty: FunctionGraphCreator.EmptyFunctionError)
        };
    }
}
