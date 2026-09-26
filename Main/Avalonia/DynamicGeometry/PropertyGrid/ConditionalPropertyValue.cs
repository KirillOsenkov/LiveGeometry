namespace DynamicGeometry;

/// <summary>
/// A figure whose property grid rows are editable or not depending on its state (a
/// segment's Length only when an end can move; a translated point's Distance only when it
/// is free or held by a Number), and whose captions may say why.
/// </summary>
public interface IConditionalProperties
{
    bool CanEdit(string propertyName);

    /// <summary>The caption to show for the row; the default one when there is nothing to add</summary>
    string Caption(string propertyName, string defaultCaption);
}

/// <summary>
/// Put on a property with <see cref="PropertyGridCustomValueProvider"/>: the grid asks the
/// figure (an <see cref="IConditionalProperties"/>) whether the row can be edited and what
/// it is called, each time the row is built or refreshed.
/// </summary>
public class ConditionalPropertyValue : PropertyValue
{
    IConditionalProperties Conditions
    {
        get { return (IConditionalProperties)Parent; }
    }

    public override bool CanSetValue
    {
        get { return base.CanSetValue && Conditions.CanEdit(Property.Name); }
    }

    public override string DisplayName
    {
        get { return Conditions.Caption(Property.Name, base.DisplayName); }
    }
}
