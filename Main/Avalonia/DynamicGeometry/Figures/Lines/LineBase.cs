using Avalonia;
using Avalonia.Controls.Shapes;

namespace DynamicGeometry
{
    public abstract partial class LineBase : ShapeBase<Line>, ILinearFigure
    {
        protected override Line CreateShape()
        {
            return Factory.CreateLineShape();
        }

        /// <summary>Line g, h... when no points name it</summary>
        protected override string FirstLetter
        {
            get
            {
                return "g";
            }
        }

        /// <summary>"to line g": a parallel or a bisector is a line to what is built on it</summary>
        public override string Noun
        {
            get
            {
                return "line";
            }
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
            if (IsShown)
            {
                // (a line whose points are nowhere - asked for a point it is not built on
                // while a file is read - is not laid out: Avalonia throws for a place at infinity)
                var coordinates = OnScreenCoordinates;
                if (coordinates.P1.Exists() && coordinates.P2.Exists())
                {
                    Shape.Set(ToPhysical(coordinates));
                    Shape.Visibility = Visibility.Visible;
                    return;
                }
            }

            Shape.Visibility = Visibility.Collapsed;
        }

        public virtual PointPair Coordinates
        {
            get { return new PointPair(Point(0), Point(1)); }
        }

        /// <summary>The name written next to the line (<see cref="FigureLabel"/>)</summary>
        [PropertyGridVisible]
        [PropertyGridName("Show name")]
        public virtual bool ShowName
        {
            get { return HasNameLabel; }
            set { HasNameLabel = value; }
        }

        public override Point Center
        {
            get
            {
                return Coordinates.Midpoint;
            }
        }

        public override IFigure HitTest(Avalonia.Point point)
        {
            var epsilon = ToLogical(this.Shape.StrokeThickness) / 2 + CursorTolerance;
            if (Math.IsPointOnLine(Coordinates, point, epsilon))
            {
                return this;
            }
            return null;
        }

        public virtual double GetNearestParameterFromPoint(Avalonia.Point point)
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