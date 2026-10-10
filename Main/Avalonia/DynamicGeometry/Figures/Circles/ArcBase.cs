using Avalonia;
using Avalonia.Media;
using Avalonia.Controls.Shapes;

namespace DynamicGeometry
{
    public interface IAngleProvider
    {
        double Angle { get; }
    }

    public interface IArc : IFigure, ILinearFigure, IEllipse, IAngleProvider, IHasSweep
    {
        // Implemented by EllipseArc, EllipseSegment, and CircleSegment.
        double EndAngle { get; }
        double StartAngle { get; }
        Point EndLocation { get; }
        Point BeginLocation { get; }

        /// <summary>Whether the arc goes clockwise from its begin to its end right now (see <see cref="Sweep"/>)</summary>
        bool IsClockwise { get; }
    }

    public abstract partial class EllipseArcBase : ShapeBase<Path>, IArc, ILengthProvider
    {
        protected EllipseArcBase()
        {
            sweep = DefaultSweep;
        }

        /// <summary>Named as circles are: c, d...</summary>
        protected override string FirstLetter
        {
            get
            {
                return "c";
            }
        }

        /// <summary>"with center O from A to B"</summary>
        public override string Construction
        {
            get
            {
                if (Dependencies.Count <= System.Math.Max(BeginPointIndex, EndPointIndex))
                {
                    return null;
                }

                return "with center " + ConstructionText.Of(Dependencies[0])
                    + " from " + ConstructionText.Of(Dependencies[BeginPointIndex])
                    + " to " + ConstructionText.Of(Dependencies[EndPointIndex]);
            }
        }

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
                // as for Ellipse: the third point's distance from the long axis
                return Math.GetDistanceToLine(Point(2), new PointPair(Point(0), Point(1)));
            }
        }

        protected override Path CreateShape()
        {
            ArcShape = new ArcSegment()
            {
                SweepDirection = SweepDirection.CounterClockwise,
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

        AngleSweep sweep;

        /// <summary>
        /// Which of the two arcs between the begin and the end this is (<see cref="AngleSweep"/>):
        /// the one counterclockwise from the begin, the one clockwise, or whichever is under
        /// or over 180°. The arc drawn, its length, the area of its sector or segment, the
        /// parameter of a point on it and where it is hit all follow it.
        /// </summary>
        [PropertyGridVisible]
        public virtual AngleSweep Sweep
        {
            get
            {
                return sweep;
            }
            set
            {
                if (sweep != value)
                {
                    sweep = value;
                    if (Drawing != null)
                    {
                        // a point on the arc, its length: what is built on it follows
                        this.RecalculateAllDependents();
                    }
                }
            }
        }

        /// <summary>What a new one is: the way round the tool's clicks went (an angle's mark says the smaller angle)</summary>
        protected virtual AngleSweep DefaultSweep
        {
            get
            {
                return AngleSweep.Counterclockwise;
            }
        }

        /// <summary>The counterclockwise angle from the begin to the end, 0 to 2π: what the sweep chooses from</summary>
        protected double CounterclockwiseAngle
        {
            get
            {
                return Math.OAngle(BeginLocation, Center, EndLocation);
            }
        }

        /// <summary>Whether the arc goes clockwise from its begin to its end right now</summary>
        public bool IsClockwise
        {
            get
            {
                return Sweep.IsClockwise(CounterclockwiseAngle);
            }
        }

        /// <summary>
        /// The length of the arc. An ellipse has no formula for it: Simpson's rule over the
        /// parameter. (It was "not a number", and so was whatever took its length from an
        /// elliptical arc: a measurement, the radius of a circle.)
        /// </summary>
        public virtual double Length
        {
            get
            {
                double start = ParametricAngle(IsClockwise ? EndLocation : BeginLocation);
                return Math.EllipseArcLength(SemiMajor, SemiMinor, start, ParametricSweep);
            }
        }

        /// <summary>The point of the arc's ellipse at the parametric angle t</summary>
        public Point PointAtParametricAngle(double t)
        {
            return Math.PointOnEllipse(Center, SemiMajor, SemiMinor, Inclination, t);
        }

        /// <summary>The point half way along the arc (where a perimeter's number sits)</summary>
        public Point ArcMiddle
        {
            get
            {
                double begin = ParametricAngle(BeginLocation);
                double sweep = ParametricSweep;
                return PointAtParametricAngle(begin + (IsClockwise ? -sweep : sweep) / 2);
            }
        }

        public override void UpdateVisual()
        {
            var startPoint = BeginLocation;
            var endPoint = EndLocation;

            ArcShape.Size = new Size(ToPhysical(SemiMajor), ToPhysical(SemiMinor));
            Figure.StartPoint = ToPhysical(startPoint);
            ArcShape.Point = ToPhysical(endPoint);
            ArcShape.RotationAngle = -Inclination.ToDegrees();
            ArcShape.SweepDirection = IsClockwise ? SweepDirection.Clockwise : SweepDirection.CounterClockwise;
            ArcShape.IsLargeArc = Angle > Math.PI;
        }

        public virtual double GetNearestParameterFromPoint(Point point)
        {
            var result = Math.GetAngle(Center, point);
            bool clockwise = IsClockwise;
            var a1 = clockwise ? EndAngle : StartAngle;
            var a2 = clockwise ? StartAngle : EndAngle;
            if (!Settings.PointsOnEllipticalsUseAbsoluteAngle)
            {
                var inclination = Inclination;
                result -= inclination;
                a1 -= inclination;
                a2 -= inclination;
            }

            // Off the arc: its nearer end, the short way round. (An arc that doesn't cross
            // the angle 0 took every direction below its start for the start and every one
            // above its end for the end: from the far side of the circle a dragged point
            // jumped to the wrong end half of the time.)
            bool onArc = a2 < a1
                ? result <= a2 || result >= a1
                : result >= a1 && result <= a2;
            if (!onArc)
            {
                double toStart = System.Math.Abs(System.Math.IEEERemainder(result - a1, 2 * Math.PI));
                double toEnd = System.Math.Abs(System.Math.IEEERemainder(result - a2, 2 * Math.PI));
                result = toStart <= toEnd ? a1 : a2;
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
            var angleToPoint = Math.GetAngle(Center, point);
            bool between = Math.IsAngleBetweenAngles(angleToPoint, StartAngle, EndAngle, IsClockwise);
            if (between)
            {
                var fromEdge = Math.RadialDistanceToEllipse(
                    Center,
                    SemiMajor,
                    SemiMinor,
                    Inclination,
                    point);
                if (fromEdge.Abs() < CursorTolerance + width / 2)
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
            bool clockwise = IsClockwise;
            var a1 = clockwise ? EndAngle : StartAngle;
            var a2 = clockwise ? StartAngle : EndAngle;
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
        /// The central angle of the arc, 0 to 2π: the measure of the region the sweep chooses
        /// (what a tool or an expression takes the arc for).
        /// </summary>
        public virtual double Angle
        {
            get
            {
                return Sweep.Measure(CounterclockwiseAngle);
            }
        }

        /// <summary>
        /// How far the arc goes around, 0 to 2π, in the angle t that parametrizes its
        /// ellipse (x = a cos t, y = b sin t; on a circle, the central angle). The areas go
        /// by it: stretching the unit circle into the ellipse keeps t and scales every area
        /// by a * b.
        /// </summary>
        double ParametricSweep
        {
            get
            {
                double begin = ParametricAngle(BeginLocation);
                double end = ParametricAngle(EndLocation);
                double result = IsClockwise ? begin - end : end - begin;
                if (result < 0)
                {
                    result += 2 * Math.PI;
                }

                return result;
            }
        }

        double ParametricAngle(Point point)
        {
            var center = Center;
            double inclination = Inclination;
            double cos = System.Math.Cos(inclination);
            double sin = System.Math.Sin(inclination);
            double dx = point.X - center.X;
            double dy = point.Y - center.Y;

            // in the ellipse's own axes
            double along = dx * cos + dy * sin;
            double across = -dx * sin + dy * cos;
            return System.Math.Atan2(across / SemiMinor, along / SemiMajor);
        }

        /// <summary>The area between the arc and the two radii to its ends: a * b * t / 2</summary>
        protected double SectorArea
        {
            get
            {
                double a = SemiMajor;
                double b = SemiMinor;
                return a > 0 && b > 0 ? a * b * ParametricSweep / 2 : 0;
            }
        }

        /// <summary>
        /// The area between the arc and its chord: a * b * (t - sin t) / 2, which past half
        /// way round is the sector plus the triangle (the sine is negative there)
        /// </summary>
        protected double SegmentArea
        {
            get
            {
                double a = SemiMajor;
                double b = SemiMinor;
                if (!(a > 0 && b > 0))
                {
                    return 0;
                }

                double sweep = ParametricSweep;
                return a * b * (sweep - System.Math.Sin(sweep)) / 2;
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
            sweep = element.ReadSweep(DefaultSweep);
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            if (sweep != DefaultSweep)
            {
                writer.WriteAttributeString("Sweep", sweep.ToString());
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

        /// <summary>
        /// No arc while its radius is 0 or its end is on the center (Shift snaps both to
        /// one place on the grid): there is no direction to end in, and its length, area
        /// and path were "NaN".
        /// </summary>
        public override void UpdateExistence()
        {
            base.UpdateExistence();
            if (Exists && Dependencies.Count > 2 && (!(Radius > 0) || !(Center.Distance(Point(2)) > 0)))
            {
                Exists = false;
            }
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
