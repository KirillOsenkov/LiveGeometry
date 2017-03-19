using System.Linq;

namespace DynamicGeometry
{
    public class Ray : LineBase, ILine
    {
        public override PointPair OnScreenCoordinates
        {
            get
            {
                var c = Coordinates;
                c = Math.GetLineFromSegment(c, CanvasLogicalBorders);
                c.P1 = Coordinates.P1;
                return c;
            }
        }

        public override double GetNearestParameterFromPoint(System.Windows.Point point)
        {
            var parameter = base.GetNearestParameterFromPoint(point);
            if (parameter < 0)
            {
                parameter = 0;
            }
            return parameter;
        }

        public override IFigure HitTest(System.Windows.Point point)
        {
            var hit = base.HitTest(point) != null;
            var line = Coordinates;
            var basement = Math.GetProjectionPoint(point, Coordinates);
            var inside = 
                   ((line.P1.X < line.P2.X && basement.X >= line.P1.X)
                   || (line.P1.X >= line.P2.X && basement.X <= line.P1.X))
                && ((line.P1.Y < line.P2.Y && basement.Y >= line.P1.Y)
                   || (line.P1.Y >= line.P2.Y && basement.Y <= line.P1.Y));
            if (hit && inside)
            {
                return this;
            }
            return null;
        }

        public override Tuple<double, double> GetParameterDomain()
        {
            return new Tuple<double, double>(0, base.GetParameterDomain().Item2);
        }

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert to line")]
        public void ConvertToLine()
        {
            LineTwoPoints.Convert(this, Factory.CreateLineTwoPoints(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert to segment")]
        public void ConvertToSegment()
        {
            LineTwoPoints.Convert(this, Factory.CreateSegment(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Reverse")]
        public void Reverse()
        {
            LineTwoPoints.Convert(this, Factory.CreateRay(this.Drawing, this.Dependencies.Reverse().ToList()));
        }

#endif
    }
}