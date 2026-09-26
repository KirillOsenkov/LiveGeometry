using System;
using System.Collections.Generic;

namespace DynamicGeometry;

/// <summary>
/// The side panel right after a figure with a length is made (a segment, a vector, a
/// regular polygon, a circle): just its length and Fix length or Free length, forwarded to
/// the figure, so that "a segment of length 2" is draw, type 2, click Fix - without a
/// length box on every tool. Captions come from the figure (Side, Radius, Fix radius);
/// the title is the figure's name. The figure's own grid has the same rows.
/// </summary>
public class LengthPanel : IConditionalProperties, ICustomMethodProvider
{
    readonly IFixableLength figure;

    public LengthPanel(IFixableLength figure)
    {
        this.figure = figure;
    }

    [PropertyGridVisible]
    [PropertyGridPreferredEditor("UpDown")]
    [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
    public double Length
    {
        get { return figure.Length; }
        set { figure.Length = value; }
    }

    [PropertyGridName("Fix length")]
    [PropertyGridIcon(PropertyGridIcon.Lock)]
    public void FixLength()
    {
        figure.FixLength();
        ShowAgain();
    }

    [PropertyGridName("Free length")]
    [PropertyGridIcon(PropertyGridIcon.Unlock)]
    public void FreeLength()
    {
        figure.FreeLength();
        ShowAgain();
    }

    // the figure shows its own grid after the verb; this panel comes back on top of that
    void ShowAgain()
    {
        figure.Drawing.RaiseDisplayProperties(this);
    }

    /// <summary>The verb that applies, captioned by the figure ("Fix radius")</summary>
    public IEnumerable<IOperationDescription> GetMethods()
    {
        foreach (var name in new[] { "FixLength", "FreeLength" })
        {
            if (figure.CanEdit(name))
            {
                yield return new CaptionedMethod(MethodDescription.Get<LengthPanel>(name), figure);
            }
        }
    }

    public bool CanEdit(string propertyName)
    {
        return figure.CanEdit(propertyName);
    }

    public string Caption(string propertyName, string defaultCaption)
    {
        return figure.Caption(propertyName, defaultCaption);
    }

    public override string ToString()
    {
        return figure.Name;
    }

    /// <summary>A method button whose caption the figure decides</summary>
    class CaptionedMethod : IOperationDescription
    {
        readonly MethodDescription method;
        readonly IConditionalProperties captions;

        public CaptionedMethod(MethodDescription method, IConditionalProperties captions)
        {
            this.method = method;
            this.captions = captions;
        }

        public string DisplayName
        {
            get { return captions.Caption(method.Name, method.DisplayName); }
        }

        public string Name
        {
            get { return method.Name; }
        }

        public object Parent
        {
            get { return method.Parent; }
        }

        public IEnumerable<IValueProvider> Parameters
        {
            get { return method.Parameters; }
        }

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
}
