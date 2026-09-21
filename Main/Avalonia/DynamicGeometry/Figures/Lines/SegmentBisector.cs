using Avalonia;

namespace DynamicGeometry
{
    public class SegmentBisector : PerpendicularLineBase
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

        protected override bool TryGetRightAngle(out Point vertex, out PointPair baseLine, out Point pointAcross)
        {
            // in the middle of the two points, whether or not a segment is drawn between them
            baseLine = new PointPair(Point(0), Point(1));
            vertex = Math.Midpoint(baseLine.P1, baseLine.P2);

            // nothing to prefer a side by: the mark starts in corner 0
            var along = baseLine.P2.Minus(baseLine.P1);
            pointAcross = vertex.Plus(new Point(-along.Y, along.X));
            return true;
        }
    }
}
