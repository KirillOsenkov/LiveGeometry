using System;
using System.Collections.Generic;
using System.Text;

namespace LiveGeometry;

public class CompileResult
{
    public Func<double, double> Function;
    public Func<double> Expression { get; set; }

    /// <summary>The tree the delegate was made of, its names resolved (what <see cref="LinearForms"/> reads); null when it didn't compile</summary>
    public BoundExpression Bound { get; set; }
    public readonly List<IFigure> Dependencies = new List<IFigure>();
    public readonly List<CompileError> Errors = new List<CompileError>();

    /// <summary>
    /// What the pieces of the text were read as, in the order of the text: each name as the
    /// binder resolved it, each number as typed (<see cref="ExpressionClassifier"/>).
    /// Operators, parentheses and commas are left out. Whole only when the text compiled.
    /// </summary>
    public readonly List<ClassifiedSpan> Classification = new List<ClassifiedSpan>();

    public bool IsSuccess => Errors.IsEmpty() && (Expression != null || Function != null);

    public void Classify(Token token, ExpressionClass kind)
    {
        Classification.Add(new ClassifiedSpan(token.Start, token.Length, kind));
    }

    public void AddError(string error)
    {
        Errors.Add(new CompileError()
        {
            Text = error
        });
    }

    public void AddBindError(string figureName)
    {
        AddError(string.Format("Could not find figure with name '{0}'", figureName));
    }

    public void AddPropertyNotFoundError(IFigure figure, string propertyName)
    {
        AddError(string.Format("Could not find property '{0}' on figure '{1}'", propertyName, figure.Name));
    }

    public void AddMethodNotFoundError(string functionName)
    {
        AddError(string.Format("There is no function '{0}'", functionName));
    }

    /// <summary>
    /// What to show under the box the text came from: null when it compiled. An empty
    /// text compiles to nothing without an error, so it gets <paramref name="whenEmpty"/>.
    /// </summary>
    public string GetErrorText(string whenEmpty = "Type a number or an expression.")
    {
        if (IsSuccess)
        {
            return null;
        }

        return Errors.IsEmpty() ? whenEmpty : ToString();
    }

    public override string ToString()
    {
        StringBuilder sb = new StringBuilder();
        foreach (var error in Errors)
        {
            if (sb.Length > 0)
            {
                sb.AppendLine();
            }
            sb.Append(error.Text);
        }
        return sb.ToString();
    }

    public void AddDependencyCycleError(string figureName)
    {
        AddError(string.Format("Using figure '{0}' will create a cycle. Circular dependencies are not allowed.", figureName));
    }

    public void AddUnknownIdentifierError(string text)
    {
        AddError(string.Format("Unknown identifier: '{0}'", text));
    }

    public void AddFigureIsNotAPointError(string longestPrefix)
    {
        AddError(string.Format("Figure '{0}' is not a point.", longestPrefix));
    }

    public void AddIncorrectNumberOfArgumentsError(FunctionInfo function, int actualNumberOfArguments)
    {
        AddError(string.Format("'{0}' is written {1}, and here it has {2} {3}",
            function.Name,
            function.Signature,
            actualNumberOfArguments,
            actualNumberOfArguments == 1 ? "argument" : "arguments"));
    }
}
