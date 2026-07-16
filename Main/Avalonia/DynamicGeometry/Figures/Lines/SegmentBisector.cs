using System.Linq;

namespace DynamicGeometry
{
    public class SegmentBisector : LineTwoPoints
    {
        PointPair coordinates;

        public override PointPair Coordinates
        {
            get
            {
                return coordinates;
            }
        }

        public override void Recalculate()
        {
            var p1 = Point(0);
            var p2 = Point(1);
            var line = (Flipped) ? new PointPair(p2, p1) : new PointPair(p1, p2);
            var midpoint = Math.Midpoint(p1, p2);
            var perpendicular = Math.GetPerpendicularLine(line, midpoint);
            coordinates = perpendicular;
        }
    }
}