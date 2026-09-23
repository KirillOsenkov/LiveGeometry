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
    }
}
