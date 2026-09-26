using Avalonia;

namespace DynamicGeometry
{
    /// <summary>
    /// The radius is either the distance between two points or the length of a figure
    /// (a segment); the center is the last dependency either way.
    /// </summary>
    public class CircleByRadius : CircleBase, IShapeWithInterior
    {
        public override Point Center
        {
            get { return Point(Dependencies.Count - 1); }
        }

        public override double Radius
        {
            get
            {
                var length = Dependencies[0] as ILengthProvider;
                if (length != null)
                {
                    return length.Length;
                }

                return Point(0).Distance(Point(1));
            }
        }

        // the radius points, when the radius is two points and not a figure's length
        protected override IPoint RadiusPivot
        {
            get { return Dependencies.Count > 2 ? Dependencies[0] as IPoint : null; }
        }

        protected override IPoint RadiusEnd
        {
            get { return Dependencies.Count > 2 ? Dependencies[1] as IPoint : null; }
        }
    }
}
