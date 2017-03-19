using System.Windows;
using System.Windows.Shapes;

namespace DynamicGeometry
{
    public class CircleByRadius : CircleBase, IShapeWithInterior
    {
        public override Point Center
        {
            get { return Point(2); }
        }

        public override double Radius
        {
            get { return Point(0).Distance(Point(1)); }
        }
    }
}
