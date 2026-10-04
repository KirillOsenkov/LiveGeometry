using Avalonia;

namespace DynamicGeometry
{
    public class Ellipse : EllipseBase, IShapeWithInterior
    {
        protected override string Kind
        {
            get
            {
                return "Ellipse";
            }
        }

        /// <summary>"with center O and axes to A and B" (B is anywhere at the short axis's distance)</summary>
        public override string Construction
        {
            get
            {
                if (Dependencies.Count < 3)
                {
                    return null;
                }

                return "with center " + ConstructionText.Of(Dependencies[0])
                    + " and axes to " + ConstructionText.Of(Dependencies[1])
                    + " and " + ConstructionText.Of(Dependencies[2]);
            }
        }

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
