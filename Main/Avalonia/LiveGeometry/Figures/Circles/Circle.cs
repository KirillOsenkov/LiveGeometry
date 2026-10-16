using Avalonia;

namespace LiveGeometry;

public class Circle : CircleBase, IShapeWithInterior
{
    public override Point Center => Point(0);

    public override double Radius => Center.Distance(Point(1));

    /// <summary>"with center A through B"</summary>
    public override string Construction
    {
        get
        {
            if (Dependencies.Count < 2)
            {
                return null;
            }

            return "with center " + ConstructionText.Of(Dependencies[0]) + " through " + ConstructionText.Of(Dependencies[1]);
        }
    }

    protected override IPoint RadiusPivot => (IPoint)Dependencies[0];

    protected override IPoint RadiusEnd => (IPoint)Dependencies[1];

    public override double Inclination => Math.GetAngle(Center, Point(1));
}
