using System.Windows;
using System.Windows.Shapes;

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

        public override double SemiMinor
        {
            get { return Center.Distance(Point(2)); }
        }

        public override double Inclination
        {
            get { return Math.GetAngle(Center, Point(1)); }
        }

    }
}
