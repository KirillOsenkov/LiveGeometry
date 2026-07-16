namespace DynamicGeometry
{
    /// <summary>
    /// A type can implement this interface to declare that it knows that it can be displayed
    /// in a property grid. As a reward, any time it is selected into a property grid,
    /// the host property grid will set this property on the object to itself,
    /// so that the object can refer to it's host property grid anytime it wants.
    /// </summary>
    public interface IPropertyGridHost
    {
#if !PLAYER
        PropertyGrid PropertyGrid { get; set; }
#endif
    }
}