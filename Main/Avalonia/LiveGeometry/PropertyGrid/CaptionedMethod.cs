using System;
using System.Collections.Generic;

namespace LiveGeometry;

/// <summary>
/// A method button whose caption the object decides (<see cref="IConditionalProperties.Caption"/>):
/// "Fix radius" on a circle, "Snap to Segment1" on a point.
/// </summary>
public class CaptionedMethod : IOperationDescription
{
    readonly IOperationDescription method;
    readonly IConditionalProperties captions;

    public CaptionedMethod(IOperationDescription method, IConditionalProperties captions)
    {
        this.method = method;
        this.captions = captions;
    }

    public string DisplayName => captions.Caption(method.Name, method.DisplayName);

    public string Name => method.Name;

    public object Parent => method.Parent;

    public IEnumerable<IValueProvider> Parameters => method.Parameters;

    public void Invoke(object target, IEnumerable<object> arguments)
    {
        method.Invoke(target, arguments);
    }

    public T GetAttribute<T>() where T : Attribute
    {
        return method.GetAttribute<T>();
    }

    public IEnumerable<T> GetAttributes<T>() where T : Attribute
    {
        return method.GetAttributes<T>();
    }
}
