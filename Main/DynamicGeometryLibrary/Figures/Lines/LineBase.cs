using System.Windows;
using System.Windows.Shapes;

namespace DynamicGeometry
{
    public abstract partial class LineBase : ShapeBase<Line>, ILinearFigure
    {
        protected override Line CreateShape()
        {
            return Factory.CreateLineShape();
        }

        public virtual PointPair OnScreenCoordinates
        {
            get
            {
                return Coordinates;
            }
        }

        public override void UpdateVisual()
        {
            if (Exists && Visible)
            {
                Shape.Set(ToPhysical(OnScreenCoordinates));
                Shape.Visibility = Visibility.Visible;
            }
            else
            {
                Shape.Visibility = Visibility.Collapsed;
            }
        }

        public virtual PointPair Coordinates
        {
            get { return new PointPair(Point(0), Point(1)); }
        }

        public override Point Center
        {
            get
            {
                return Coordinates.Midpoint;
            }
        }

        public override IFigure HitTest(System.Windows.Point point)
        {
            var epsilon = ToLogical(this.Shape.StrokeThickness) / 2 + CursorTolerance;
            if (Math.IsPointOnLine(Coordinates, point, epsilon))
            {
                return this;
            }
            return null;
        }

        public virtual double GetNearestParameterFromPoint(System.Windows.Point point)
        {
            var projection = Math.GetProjection(point, Coordinates);
            return projection.Ratio;
        }

        public Point GetPointFromParameter(double parameter)
        {
            return Math.ScalePointBetweenTwo(Coordinates, parameter);
        }

        public virtual Tuple<double, double> GetParameterDomain()
        {
            var coordinates = OnScreenCoordinates;
            var p1 = GetNearestParameterFromPoint(coordinates.P1);
            var p2 = GetNearestParameterFromPoint(coordinates.P2);
            return new Tuple<double, double>(p1 * 2, p2 * 2);
        }

#if TABULA
        [PropertyGridVisible]   // I expose this to the user.  Not sure if it should be exposed in other implementations.
#endif
        public virtual double Angle
        {
            get
            {
                return Math.GetAngle(Coordinates.P1, Coordinates.P2).ToDegrees();
            }
            set
            {
                if (Dependencies.Count == 2)
                {
                    var startpoint = Dependencies[0] as IPoint;
                    var endpoint = Dependencies[1] as FreePoint;
                    if (startpoint != null && endpoint != null)
                    {
                        var angleToRotate = value - Angle;
                        var newCoordinates = Math.GetRotationPoint(endpoint.Coordinates, startpoint.Coordinates, angleToRotate.ToRadians());
                        endpoint.MoveTo(newCoordinates);
                        endpoint.RecalculateAllDependents();
                    }
                }
            }
        }

    }
}