using Avalonia;

namespace DynamicGeometry
{
    public class PerpendicularLine : PerpendicularLineBase
    {
        public override PointPair Coordinates
        {
            get
            {
                var line = Dependencies.Line(0);
                if (Flipped) line = new PointPair(line.P2,line.P1);
                var point = Point(1);
                var coordinates = Math.GetPerpendicularLine(line, point);
                return coordinates;
            }
        }

        protected override bool TryGetRightAngle(out Point vertex, out PointPair baseLine, out Point pointAcross)
        {
            // where this line crosses the one it is perpendicular to - if it does:
            // the foot can be beyond the end of a segment
            var baseFigure = Dependencies[0];
            baseLine = Dependencies.Line(0);
            pointAcross = Point(1);
            vertex = Math.GetProjectionPoint(pointAcross, baseLine);
            return baseFigure.Visible && baseFigure.HitTest(vertex) != null;
        }
    }
}
