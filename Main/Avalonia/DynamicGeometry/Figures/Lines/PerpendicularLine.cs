using Avalonia;

namespace DynamicGeometry
{
    public class PerpendicularLine : PerpendicularLineBase
    {
        protected override string Kind
        {
            get
            {
                return "Perpendicular line";
            }
        }

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
            baseLine = Dependencies.Line(0);
            pointAcross = Point(1);
            vertex = Math.GetProjectionPoint(pointAcross, baseLine);
            return Dependencies[0].Visible && BaseFigure.HitTest(vertex) != null;
        }

        // a vector by the segment inside it: its arrowhead is a polygon that has no points
        // while the file is read, and the mark is about where the vector ends, not its head
        protected override IFigure BaseFigure
        {
            get
            {
                var vector = Dependencies[0] as Vector;
                if (vector != null)
                {
                    return vector.Line;
                }

                return Dependencies[0];
            }
        }
    }
}
