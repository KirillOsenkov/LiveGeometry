using System;
using System.Collections.Generic;
using System.Linq;

namespace DynamicGeometry;

/// <summary>
/// A button of the property grid that is not a method of the shown object: a caption, an
/// icon and what to do, for a panel whose buttons are decided at run time
/// (<see cref="TiedValuesPanel"/>: one "Type the ..." per tied value).
/// </summary>
public class DelegateOperation : IOperationDescription
{
    readonly Action action;
    readonly PropertyGridIconAttribute icon;

    public DelegateOperation(
        string name,
        string caption,
        PropertyGridIcon icon,
        Action action)
    {
        Name = name;
        DisplayName = caption;
        this.icon = new PropertyGridIconAttribute(icon);
        this.action = action;
    }

    public string Name { get; }

    public string DisplayName { get; }

    public object Parent
    {
        get { return null; }
    }

    public IEnumerable<IValueProvider> Parameters
    {
        get { return Enumerable.Empty<IValueProvider>(); }
    }

    public void Invoke(object target, IEnumerable<object> arguments)
    {
        action();
    }

    public T GetAttribute<T>() where T : Attribute
    {
        return icon as T;
    }

    public IEnumerable<T> GetAttributes<T>() where T : Attribute
    {
        var one = GetAttribute<T>();
        return one != null ? new[] { one } : Enumerable.Empty<T>();
    }
}
