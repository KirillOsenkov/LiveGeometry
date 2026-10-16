using System.Linq;
using Avalonia.Controls;

namespace LiveGeometry;

public class DrawingExpressionEditorFactory
    : BaseValueEditorFactory<ExpressionEditor, DrawingExpression> { }

/// <summary>
/// An expression of a figure (a coordinate of a point by coordinates, a number's formula):
/// each keystroke that makes a valid expression applies it; what is wrong shows under the box
/// on commit, and what the text is worth shows there while it compiles. The box is an
/// <see cref="ExpressionTextBox"/>: a click on a figure puts a reference into it.
/// </summary>
public class ExpressionEditor : StringEditor, IExpressionHost
{
    protected override TextBox CreateTextBoxCore()
    {
        return new ExpressionTextBox() { Host = this };
    }

    /// <summary>The expression the row edits; of a selection, the first (a value of the selection is not shown)</summary>
    DrawingExpression Expression
    {
        get
        {
            return Value is CompositeValueProvider composite
                ? composite.InnerList.OfType<DrawingExpression>().FirstOrDefault()
                : Value as DrawingExpression;
        }
    }

    public Drawing Drawing => Expression?.ParentFigure?.Drawing;

    public IFigure Owner => Expression?.ParentFigure;

    public InsertContext Context => InsertContexts.FromRowName(Value?.DisplayName);

    protected override void OnBoxTextChanged(string text)
    {
        ValueText = Value is CompositeValueProvider ? null : ((ExpressionTextBox)TextBox).LiveValueText(text);
    }

    /// <summary>Enter is "done": the keyboard goes back to the canvas, and clicks are the tool's again</summary>
    protected override void OnEnter()
    {
        ((ExpressionTextBox)TextBox).EndEditing();
    }

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

/// <summary>
/// Picked by <c>[PropertyGridPreferredEditor("Expression")]</c> on a string: a row of a tool's
/// panel that holds an expression (the X of the point to add, a radius). The panel checks
/// the text when its command runs; the box takes references by a click, and says what the
/// text is worth. The drawing is the one the panel is shown for, unless the panel says
/// itself whose expression it holds (<see cref="IExpressionHost"/>: the Number tool's).
/// </summary>
public class ExpressionStringEditorFactory : BaseValueEditorFactory<ExpressionStringEditor, string>
{
    public ExpressionStringEditorFactory()
    {
        // after the plain string editor: only the attribute chooses this one
        LoadOrder = 1;
    }
}

public class ExpressionStringEditor : StringEditor, IExpressionHost
{
    protected override TextBox CreateTextBoxCore()
    {
        return new ExpressionTextBox() { Host = this };
    }

    IExpressionHost Panel => Value?.Parent as IExpressionHost;

    public Drawing Drawing => Panel?.Drawing;

    public IFigure Owner => Panel?.Owner;

    public InsertContext Context => Panel?.Context ?? InsertContexts.FromRowName(Value?.DisplayName);

    protected override void OnBoxTextChanged(string text)
    {
        ValueText = ((ExpressionTextBox)TextBox).LiveValueText(text);
    }
}

/// <summary>
/// Picked by <c>[PropertyGridPreferredEditor("LabelText")]</c> on the text of a label: the
/// words and the [...] parts of a text label. A reference clicked in goes in as a [...] part,
/// unless the caret is in one already (<see cref="ExpressionTextBox.Insert"/>).
/// </summary>
public class LabelTextEditorFactory : BaseValueEditorFactory<LabelTextEditor, string>
{
    public LabelTextEditorFactory()
    {
        LoadOrder = 1;
    }
}

public class LabelTextEditor : StringEditor, IExpressionHost
{
    protected override TextBox CreateTextBoxCore()
    {
        return new ExpressionTextBox() { Host = this };
    }

    IFigure LabelFigure => Value?.Parent as IFigure;

    public Drawing Drawing => LabelFigure?.Drawing;

    public IFigure Owner => LabelFigure;

    public InsertContext Context => InsertContext.Text;
}
