using System;

namespace DynamicGeometry;

/// <summary>
/// A tool's panel whose command checks what was typed into its rows (Add point checks X and
/// Y). What is wrong goes under the row it is about (<see cref="StringEditor.ErrorText"/>)
/// when the command runs, never while the user is typing, and stays until that text is
/// edited or the command finds it right.
/// </summary>
public class ToolPanel
{
    /// <summary>A property's name and what is wrong with its text; null takes the error away</summary>
    public event Action<string, string> PropertyError;

    protected void ReportError(string propertyName, string error)
    {
        PropertyError?.Invoke(propertyName, error);
    }

    /// <summary>A property whose row should take the keyboard now</summary>
    public event Action<string> FocusRequested;

    /// <summary>
    /// Puts the keyboard into a row, so that a command that keeps the panel (Add point, for
    /// the next point) leaves the user ready to type again
    /// </summary>
    protected void FocusRow(string propertyName)
    {
        FocusRequested?.Invoke(propertyName);
    }

    /// <summary>Compiles the text of a row and says under the row what is wrong with it</summary>
    protected CompileResult Compile(Drawing drawing, string propertyName, string text)
    {
        var result = drawing.CompileExpression(text);
        ReportError(propertyName, result.GetErrorText());
        return result;
    }

    /// <summary>
    /// The value of a row's compiled text right now; NaN, with the error under the row, when
    /// it has none (the square root of a negative number) - a point there can't be shown.
    /// </summary>
    protected double Evaluate(string propertyName, CompileResult result)
    {
        double value = result.Expression();
        if (!value.IsValidValue())
        {
            ReportError(propertyName, "This has no value right now.");
            return double.NaN;
        }

        return value;
    }
}
