using System;
using DynamicGeometry;

namespace LiveGeometry;

/// <summary>
/// What the side panel shows for an exception: the message and the whole text, to be read
/// and copied. See <see cref="MainView.ReportException"/>.
/// </summary>
[PropertyGridName("Something went wrong")]
public class ExceptionReport
{
    /// <param name="stackTrace">The stack at the throw, shown when the exception's own trace is empty</param>
    public ExceptionReport(Exception exception, string stackTrace = null)
    {
        Exception = exception;
        StackTrace = stackTrace;
    }

    public Exception Exception { get; }

    public string StackTrace { get; }

    [PropertyGridVisible]
    [PropertyGridName("Error")]
    public string Message => Exception.Message;

    [PropertyGridVisible]
    [PropertyGridName("Details")]
    public string Details
    {
        get
        {
            // a JSException prints only its message: add whatever stack there is
            var text = Exception.ToString();
            foreach (var stack in new[] { Exception.StackTrace, StackTrace })
            {
                if (!string.IsNullOrEmpty(stack) && !text.Contains(stack))
                {
                    text += Environment.NewLine + stack;
                }
            }

            return text;
        }
    }
}
