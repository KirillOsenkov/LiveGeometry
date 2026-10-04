using Avalonia;

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

        protected override IPoint RadiusPivot
        {
            get { return (IPoint)Dependencies[0]; }
        }

        protected override IPoint RadiusEnd
        {
            get { return (IPoint)Dependencies[1]; }
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
