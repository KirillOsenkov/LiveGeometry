using System.Collections.Generic;
using System.Windows;

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

            Figure.StartPoint = ToPhysical(BeginLocation);
            ArcShape.Point = ToPhysical(EndLocation);

            ArcShape.IsLargeArc = Math.OAngle(BeginLocation, center, EndLocation) > Math.PI;

            // I commented out these lines because this causes AngleArc not to honor Visible property. Is this a mistake?  
            // I believe there will be times when a user would like to hide the arc. D. H.
            //if (Shape.Visibility != Visibility.Visible)
            //{
            //    Shape.Visibility = Visibility.Visible;
            //}
        }
    }
}
