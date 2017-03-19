using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Collections.Generic;
using System.Linq;

namespace DynamicGeometry
{
    public interface IAngleProvider
    {
        double Angle { get; }
    }

    public interface IArc : IFigure, ILinearFigure, IEllipse, IAngleProvider
    {
        // Implemented by EllipseArc, EllipseSegment, and CircleSegment.
        double EndAngle { get; }
        double StartAngle { get; }
        Point EndLocation { get; }
        Point BeginLocation { get; }
        bool Clockwise { get; set; }
    }

    public abstract partial class EllipseArcBase : ShapeBase<Path>, IArc, ILengthProvider
    {

        public PathFigure Figure { get; set; }
        protected ArcSegment ArcShape { get; set; }

        public virtual double SemiMajor
        {
            get
            {
                return Math.Distance(Point(0), Point(1));
            }
        }

        public virtual double SemiMinor
        {
            get
            {
                return Math.Distance(Point(0), Point(2));
            }
        }

        protected override Path CreateShape()
        {
            ArcShape = new ArcSegment()
            {
                SweepDirection = SweepDirection.Counterclockwise,
                RotationAngle = 0
            };
            Figure = new PathFigure()
            {
                IsClosed = false,
                IsFilled = true,
                Segments = new PathSegmentCollection()
                {
                    ArcShape
                }
            };
            return new Path()
            {
                Data = new PathGeometry()
                {
                    Figures = new PathFigureCollection()
                    {
                        Figure
                    }
                },
                Stroke = new SolidColorBrush(Colors.Black),
                StrokeThickness = 1
            };
        }

        public double LogicalWidth()
        {
            return ToLogical(Shape.StrokeThickness);
        }

        bool mClockwise = false;
        [PropertyGridVisible]
        public bool Clockwise
        {
            get
            {
                return mClockwise;
            }
            set
            {
                if (mClockwise != value)
                {
                    mClockwise = value;
                    if (mClockwise)
                    {
                        ArcShape.SweepDirection = SweepDirection.Clockwise;
                    }
                    else
                    {
                        ArcShape.SweepDirection = SweepDirection.Counterclockwise;
                    }
                    if (Drawing != null)
                    {
                        UpdateVisual();
                    }
                }

            }
        }

        public virtual double Length
        {
            get
            {
                // Not yet calculated for ellipse arcs.
                return double.NaN;
            }
        }

        public override void UpdateVisual()
        {
            var center = Center;
            var startPoint = BeginLocation;
            var endPoint = EndLocation;

            ArcShape.Size = new Size(ToPhysical(SemiMajor), ToPhysical(SemiMinor));
            Figure.StartPoint = ToPhysical(startPoint);
            ArcShape.Point = ToPhysical(endPoint);
            ArcShape.RotationAngle = -Inclination.ToDegrees();
            ArcShape.IsLargeArc = Clockwise ? Math.OAngle(endPoint, center, startPoint) > Math.PI :
                                              Math.OAngle(startPoint, center, endPoint) > Math.PI;
        }

        public virtual double GetNearestParameterFromPoint(Point point)
        {
            var result = Math.GetAngle(Center, point);
            var a1 = (Clockwise) ? EndAngle : StartAngle;
            var a2 = (Clockwise) ? StartAngle : EndAngle;
            if (!Settings.PointsOnEllipticalsUseAbsoluteAngle)
            {
                var inclination = Inclination;
                result -= inclination;
                a1 -= inclination;
                a2 -= inclination;
            }

            if (a2 < a1)
            {
                if (result <= a2 || result >= a1)
                {
                    //return result;
                }
                else if (result > (a1 + a2) / 2)
                {
                    result = a1;
                }
                else
                {
                    result = a2;
                }
            }
            else if (result < a1)
            {
                result = a1;
            }
            else if (result > a2)
            {
                result = a2;
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

        public override IFigure HitTest(Point point)
        {
            // HitTest for the fill.
            ShapeStyle shapeStyle = Style as ShapeStyle;
            if (shapeStyle != null)
            {
                if (shapeStyle.IsFilled)
                {
                    var fillResult = HitTestShape(point);
                    if (fillResult != null) return fillResult;
                }
            }

            // HitTest for the edge
            var width = LogicalWidth();
            var r = Math.Distance(Center, point);
            var angleToPoint = Math.GetAngle(Center, point);
            bool between = Math.IsAngleBetweenAngles(angleToPoint, StartAngle, EndAngle, Clockwise);
            if (between)
            {
                // Find the point relative to the ellipse in canonical form(unrotated).
                var canonicalPoint = Math.RotatePoint(Center, r, angleToPoint - Inclination).Minus(Center);
                var equationLeft = canonicalPoint.X.Sqr() / SemiMajor.Sqr() + canonicalPoint.Y.Sqr() / SemiMinor.Sqr();

                // A cheap way to deal with small arcs.
                var tolerance = CursorTolerance + width / 2;
                if (SemiMajor < 1 || SemiMinor < 1)
                {
                    tolerance += .25;
                }

                if ((equationLeft - 1).Abs() < tolerance)
                {
                    return this;
                }
            }

            // HitTest for the chord (if necessary).
            if (this is EllipseSegment || this is CircleSegment)
            {
                var epsilon = ToLogical(this.Shape.StrokeThickness) / 2 + CursorTolerance;
                if (Math.IsPointOnSegment(new PointPair(BeginLocation, EndLocation), point, epsilon))
                {
                    return this;
                }
            }

            // HitTest for the radii (if necessary).
            if (this is EllipseSector || this is CircleSector)
            {
                var epsilon = ToLogical(this.Shape.StrokeThickness) / 2 + CursorTolerance;
                if (Math.IsPointOnSegment(new PointPair(Center, BeginLocation), point, epsilon) ||
                    Math.IsPointOnSegment(new PointPair(Center, EndLocation), point, epsilon))
                {
                    return this;
                }
            }

            return null;
        }

        public virtual Tuple<double, double> GetParameterDomain()
        {
            var a1 = (Clockwise) ? EndAngle : StartAngle;
            var a2 = (Clockwise) ? StartAngle : EndAngle;
            if (a2 < a1)
            {
                a2 += 2 * Math.PI;
            }
            return Tuple.Create(a1, a2);
        }

        public virtual double Inclination
        {
            get
            {
                return Math.GetAngle(Center, Point(1));
            }
        }

        public abstract int BeginPointIndex { get; }
        public abstract int EndPointIndex { get; }

        public virtual Point BeginLocation
        {
            get
            {
                var center = Center;
                var point3 = Point(BeginPointIndex);
                var intersections = Math.GetIntersectionOfEllipseAndLine(center, SemiMajor, SemiMinor, Inclination, new PointPair(center, point3));
                var i1 = intersections.P1.Distance(point3);
                var i2 = intersections.P2.Distance(point3);
                var endPoint = (i1 < i2) ? intersections.P1 : intersections.P2;
                return endPoint;
            }
        }

        public virtual Point EndLocation
        {
            get
            {
                var center = Center;
                var point4 = Point(EndPointIndex);
                var intersections = Math.GetIntersectionOfEllipseAndLine(center, SemiMajor, SemiMinor, Inclination, new PointPair(center, point4));
                var i1 = intersections.P1.Distance(point4);
                var i2 = intersections.P2.Distance(point4);
                var endPoint = (i1 < i2) ? intersections.P1 : intersections.P2;
                return endPoint;
            }
        }

        public virtual double StartAngle
        {
            get
            {
                return Math.GetAngle(Center, BeginLocation);
                //return Inclination;
            }
        }

        public virtual double EndAngle
        {
            get
            {
                return Math.GetAngle(Center, EndLocation);
            }
        }

        /// <summary>
        /// The central angle of the arc.
        /// </summary>
        public virtual double Angle
        {
            get
            {
                //return Math.GetAngle(StartAngle, EndAngle);
                return Clockwise ? Math.OAngle(EndLocation, Center, BeginLocation) :
                                   Math.OAngle(BeginLocation, Center, EndLocation);
            }
        }

        public override Point Center
        {
            get
            {
                return Point(0);
            }
        }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            base.ReadXml(element);
            Clockwise = element.ReadBool("Clockwise", false);
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            if (Clockwise)
            {
                writer.WriteAttributeBool("Clockwise", Clockwise);
            }
        }
    }

    public abstract partial class CircleArcBase : EllipseArcBase, ICircle
    {

        public override Point BeginLocation
        {
            get
            {
                return Point(BeginPointIndex);
            }
        }

        public override int BeginPointIndex
        {
            get { return 1; }
        }

        public override Point EndLocation
        {
            get
            {
                return Math.ScalePointBetweenTwo(
                Center,
                Point(2),
                Radius / Center.Distance(Point(2)));
            }
        }

        public override int EndPointIndex
        {
            get { return 2; }
        }

        public override double Length
        {
            get
            {
                return Radius * Angle;
            }
        }

        public virtual double Radius
        {
            get
            {
                return SemiMajor;
            }
        }

        public override double SemiMinor
        {
            get
            {
                return SemiMajor;
            }
        }
    }

}
