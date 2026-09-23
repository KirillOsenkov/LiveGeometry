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
    public ExceptionReport(Exception exception)
    {
        Exception = exception;
    }

    public Exception Exception { get; }

    [PropertyGridVisible]
    [PropertyGridName("Error")]
    public string Message => Exception.Message;

    [PropertyGridVisible]
    [PropertyGridName("Details")]
    public string Details => Exception.ToString();
}
