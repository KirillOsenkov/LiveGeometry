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

    /// <summary>
    /// A distance measurement on the figure, added or removed. The grid records the
    /// property set as the undo step, and this setter runs inside it - so the drawing is
    /// changed directly (the undo library refuses an action recorded from within another);
    /// undo sets the property back, which removes or re-adds the measurement.
    /// </summary>
    [PropertyGridVisible]
    [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
    public bool Show
    {
        get
        {
            return LengthConstraint.FindMeasurement(figure) != null;
        }
        set
        {
            var existing = LengthConstraint.FindMeasurement(figure);
            if (value && existing == null)
            {
                figure.Drawing.Figures.Add(LengthConstraint.CreateMeasurement(figure));
            }
            else if (!value && existing != null)
            {
                figure.Drawing.Figures.Remove(existing);
            }
        }
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

    /// <summary>Closes the panel; the figure stays as it is</summary>
    [PropertyGridIcon(PropertyGridIcon.Check)]
    public void Done()
    {
        figure.Drawing.RaiseDisplayProperties(null);
        figure.Drawing.ClearStatus();
    }

    // the figure shows its own grid after the verb; this panel comes back on top of that
    void ShowAgain()
    {
        figure.Drawing.RaiseDisplayProperties(this);
    }

    /// <summary>The verb that applies, captioned by the figure ("Fix radius"), then Done</summary>
    public IEnumerable<IOperationDescription> GetMethods()
    {
        foreach (var name in new[] { "FixLength", "FreeLength" })
        {
            if (figure.CanEdit(name))
            {
                yield return new CaptionedMethod(MethodDescription.Get<LengthPanel>(name), figure);
            }
        }

        yield return MethodDescription.Get<LengthPanel>("Done");
    }

    public bool CanEdit(string propertyName)
    {
        // no figure to hang a measurement on (a regular polygon's sides are its own children)
        if (propertyName == "Show")
        {
            return figure.MeasuredFigures != null;
        }

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
