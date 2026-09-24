using Avalonia;

namespace DynamicGeometry
{
    public class Ellipse : EllipseBase, IShapeWithInterior
    {
        public override Point Center
        {
            get { return Point(0); }
        }

        public override double SemiMajor
        {
            get { return Center.Distance(Point(1)); }
        }

        /// <summary>
        /// The third point sets the short axis by its distance from the long axis, so a
        /// point on the short axis (where the tool puts it) is on the ellipse.
        /// </summary>
        public override double SemiMinor
        {
            get { return Math.GetDistanceToLine(Point(2), new PointPair(Center, Point(1))); }
        }

        public override double Inclination
        {
            get { return Math.GetAngle(Center, Point(1)); }
        }

    }
}
