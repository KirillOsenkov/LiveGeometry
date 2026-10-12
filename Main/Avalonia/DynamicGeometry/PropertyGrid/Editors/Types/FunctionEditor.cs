using System.Linq;

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
    /// <summary>
    /// Valid when it compiles for every graph the row edits: one, or each of a selection,
    /// against its own graph. The row of a selection has no one figure for a parent, and
    /// was taken for a plain string: "sin(" went into both graphs and their files.
    /// </summary>
    protected override ValidationResult Validate(object value)
    {
        string source = value.ToString();
        var graphs = (Value is CompositeValueProvider composite ? composite.InnerList.Select(inner => inner.Parent) : new[] { Value.Parent })
            .OfType<IFigure>()
            .Where(graph => graph.Drawing != null)
            .ToList();
        if (graphs.Count == 0)
        {
            return base.Validate(value);
        }

        foreach (var graph in graphs)
        {
            var compileResult = Compiler.Instance.CompileFunction(graph.Drawing, source, figure => !figure.DependsOn(graph));
            if (!compileResult.IsSuccess)
            {
                return new ValidationResult()
                {
                    IsValid = false,
                    Value = source,
                    Error = compileResult.GetErrorText(whenEmpty: FunctionGraphCreator.EmptyFunctionError)
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
