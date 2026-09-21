using System.Collections.Generic;
using Avalonia;

namespace DynamicGeometry
{
    public class AngleArc : CircleArc
    {
        public AngleArc()
            : base()
        {
            Size = 16;
            ArcShape.Size = new Size(this.Size, this.Size);
        }

        public double Size { get; set; }

        /// <summary>
        /// In radians: 0.005 degrees, which is when the label (two decimals) reads "90°".
        /// The sign and the number never disagree; an angle that merely looks right doesn't get it.
        /// </summary>
        public const double RightAngleTolerance = 0.005 * System.Math.PI / 180;

        readonly Avalonia.Media.LineSegment rightAngleSide1 = new Avalonia.Media.LineSegment();
        readonly Avalonia.Media.LineSegment rightAngleSide2 = new Avalonia.Media.LineSegment();
        bool showsRightAngle;

        public override double Radius
        {
            get
            {
                return ToLogical(Size);
            }
        }

        public override double SemiMajor
        {
            get
            {
                return Radius;
            }
        }

        public override double SemiMinor
        {
            get
            {
                return Radius;
            }
        }

        public override Point BeginLocation
        {
            get
            {
                return Math.ScalePointBetweenTwo(
                Center,
                Point(1),
                Radius / Center.Distance(Point(1)));
            }
        }
        
        [PropertyGridVisible]
        public virtual double Measure
        {
            get 
            {
                return Angle;
                //return Math.OAngle(BeginLocation, Center, EndLocation); 
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert to opposite angle")]
        public void ConvertToOpposite()
        {
            IList<IFigure> dependencies = Dependencies as IList<IFigure>;
            if (dependencies != null)
            {
                var t = dependencies[1];
                dependencies[1] = dependencies[2];
                dependencies[2] = t;
            }
            this.RecalculateAndUpdateVisual();
        }

        public override void UpdateVisual()
        {
            var center = Point(0);
            var radius = Radius;
            var distance1 = center.Distance(Point(1));
            var distance2 = center.Distance(Point(2));
            if (distance1 == 0 || distance2 == 0)
            {
                Shape.Visibility = Visibility.Collapsed;
                return;
            }

            var angle = Math.OAngle(BeginLocation, center, EndLocation);
            var isRightAngle = System.Math.Abs(angle - Math.PI / 2) < RightAngleTolerance;
            if (isRightAngle != showsRightAngle)
            {
                showsRightAngle = isRightAngle;
                Figure.Segments = isRightAngle
                    ? new Avalonia.Media.PathSegments() { rightAngleSide1, rightAngleSide2 }
                    : new Avalonia.Media.PathSegments() { ArcShape };
            }

            if (isRightAngle)
            {
                // the school sign for 90 degrees: two sides of a little square, not an arc
                var corner = ToPhysical(center);
                var first = RightAngleMark.Direction(corner, ToPhysical(Point(1)));
                var second = RightAngleMark.Direction(corner, ToPhysical(Point(2)));
                var points = RightAngleMark.GetPoints(corner, first.Value, second.Value, RightAngleMark.Size);
                Figure.StartPoint = points[0];
                rightAngleSide1.Point = points[1];
                rightAngleSide2.Point = points[2];
                return;
            }

            Figure.StartPoint = ToPhysical(BeginLocation);
            ArcShape.Point = ToPhysical(EndLocation);

            ArcShape.IsLargeArc = angle > Math.PI;

            // I commented out these lines because this causes AngleArc not to honor Visible property. Is this a mistake?  
            // I believe there will be times when a user would like to hide the arc. D. H.
            //if (Shape.Visibility != Visibility.Visible)
            //{
            //    Shape.Visibility = Visibility.Visible;
            //}
        }
    }
}
