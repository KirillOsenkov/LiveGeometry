namespace DynamicGeometry;

/// <summary>
/// An object whose enum row offers only some of the enum's values, as the object is now: a
/// length measurement of a circle offers a radius, a diameter and a circumference, not a
/// perimeter. The enum editors ask it for each value.
/// </summary>
public interface IOffersValues
{
    bool Offers(string propertyName, object value);
}
