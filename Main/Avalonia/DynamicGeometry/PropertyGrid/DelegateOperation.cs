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
    readonly Attribute[] attributes;

    /// <param name="group">The <see cref="PropertyGridGroupAttribute"/> box the button sits in, under its rows; null for a plain button</param>
    public DelegateOperation(
        string name,
        string caption,
        PropertyGridIcon icon,
        Action action,
        string group = null)
    {
        Name = name;
        DisplayName = caption;
        this.action = action;
        attributes = group != null
            ? new Attribute[] { new PropertyGridIconAttribute(icon), new PropertyGridGroupAttribute(group) }
            : new Attribute[] { new PropertyGridIconAttribute(icon) };
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
        return attributes.OfType<T>().FirstOrDefault();
    }

    public IEnumerable<T> GetAttributes<T>() where T : Attribute
    {
        return attributes.OfType<T>();
    }
}
