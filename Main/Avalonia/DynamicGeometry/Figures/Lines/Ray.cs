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

        public override double GetNearestParameterFromPoint(Avalonia.Point point)
        {
            var parameter = base.GetNearestParameterFromPoint(point);
            if (parameter < 0)
            {
                parameter = 0;
            }
            return parameter;
        }

        public override IFigure HitTest(Avalonia.Point point)
        {
            var hit = base.HitTest(point) != null;
            var line = Coordinates;

            // the start is on the ray however the rounding of a point there went
            var inside = Math.GetProjection(point, line).Ratio >= -Math.EndTolerance(line);
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

        protected override string NameFromDependencies()
        {
            return NameFromPoints();
        }

        protected override string Kind
        {
            get
            {
                return "Ray";
            }
        }

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert to line")]
        [PropertyGridIcon(PropertyGridIcon.Line)]
        public void ConvertToLine()
        {
            LineTwoPoints.Convert(this, Factory.CreateLineTwoPoints(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert to segment")]
        [PropertyGridIcon(PropertyGridIcon.Segment)]
        public void ConvertToSegment()
        {
            LineTwoPoints.Convert(this, Factory.CreateSegment(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Reverse")]
        [PropertyGridIcon(PropertyGridIcon.Reverse)]
        public void Reverse()
        {
            LineTwoPoints.Convert(this, Factory.CreateRay(this.Drawing, this.Dependencies.Reverse().ToList()));
        }

#endif
    }
}