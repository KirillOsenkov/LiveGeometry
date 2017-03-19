using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;

namespace DynamicGeometry
{
    public abstract partial class EllipseBase : ShapeBase<Shape>, IEllipse
    {

        public abstract double SemiMajor
        {
            get;
        }

        public abstract double SemiMinor
        {
            get;
        }

        /// <summary>
        /// Angle of inclination of ellipse in radians.
        /// </summary>
        public abstract double Inclination
        {
            get;
        }

        public double Area
        {
            get
            {
                return Math.PI * SemiMajor * SemiMinor;
            }
        }

        public override IFigure HitTest(Point point)
        {
            var width = LogicalWidth();
            var r = Math.Distance(Center, point);
            var angleToPoint = Math.GetAngle(Center, point);

            // Find the point relative to the ellipse in canonical form(unrotated).
            var canonicalPoint = Math.RotatePoint(Center, r, angleToPoint - Inclination).Minus(Center);
            var equationLeft = canonicalPoint.X.Sqr() / SemiMajor.Sqr() + canonicalPoint.Y.Sqr() / SemiMinor.Sqr();

            // HitTest for the edge
            if ((equationLeft - 1).Abs() < CursorTolerance + width / 2)
            {
                return this;
            }

            // HitTest for the fill.  Use this instead of base.HitTest() because it is not working properly with ellipses.
            ShapeStyle shapeStyle = Style as ShapeStyle;
            if (shapeStyle != null)
            {
                if (shapeStyle.IsFilled && equationLeft < 1)
                {
                    return this;
                }
            }
            return null;
        }

        public double LogicalWidth()
        {
            return ToLogical(shape.StrokeThickness);
        }

        public virtual double GetNearestParameterFromPoint(Point point)
        {
            double result;
            if (Settings.PointsOnEllipticalsUseAbsoluteAngle)
            {
                result = Math.GetAngle(Center, point);
            }
            else
            {
                result = Math.GetAngle(Center, point) - Inclination;
            }
            if (Flipped) result = -result;
            return result;
        }

        public virtual Point GetPointFromParameter(double parameter)
        {
            var center = Center;
            var inclination = Inclination;
            var angleToPoint = parameter;
            if (Flipped) angleToPoint = -angleToPoint;
            if (!Settings.PointsOnEllipticalsUseAbsoluteAngle) angleToPoint += inclination;
            var intersections = Math.GetIntersectionOfEllipseAndLine(this, new PointPair(center, Math.GetTranslationPoint(center, 1, angleToPoint)));
            var direction1 = Math.GetAngle(center, intersections.P1);
            var cDiff = System.Math.Cos(angleToPoint) - System.Math.Cos(direction1);
            var sDiff = System.Math.Sin(angleToPoint) - System.Math.Sin(direction1);
            if (cDiff.IsWithinEpsilon() && sDiff.IsWithinEpsilon())
            {
                return intersections.P1;
            }
            else
            {
                return intersections.P2;
            }
        }

        public virtual Tuple<double, double> GetParameterDomain()
        {
            return Tuple.Create(0.0, DynamicGeometry.Math.DOUBLEPI);
        }

        public override void UpdateVisual()
        {
            var center = ToPhysical(Center);
            var logicalWidth = LogicalWidth();
            var major = ToPhysical(SemiMajor * 2 + logicalWidth);
            var minor = ToPhysical(SemiMinor * 2 + logicalWidth);
            double angle = -Inclination.ToDegrees();
            RotateTransform rotation = new RotateTransform();
            rotation.CenterX = major / 2;
            rotation.CenterY = minor / 2;
            rotation.Angle = angle;
            Shape.RenderTransform = rotation;
            Shape.Width = major;
            Shape.Height = minor;
            Shape.CenterAt(center);
        }

        protected override Shape CreateShape()
        {
            return Factory.CreateCircleShape();
        }
    }
}
