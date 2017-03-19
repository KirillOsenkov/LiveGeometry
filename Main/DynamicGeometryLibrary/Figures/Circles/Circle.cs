using System.Windows;

namespace DynamicGeometry
{
    public class Circle : CircleBase, IShapeWithInterior
    {
        public override Point Center
        {
            get { return Point(0); }
        }

        public override double Radius
        {
            get { return Center.Distance(Point(1)); }
        }

        public override double Inclination
        {
            get
            {
                return Math.GetAngle(Center, Point(1));
            }
        }
    }
}
