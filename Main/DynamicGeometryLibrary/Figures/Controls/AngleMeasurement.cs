using System.Collections.Generic;
using System.Windows;

namespace DynamicGeometry
{
    public class AngleMeasurementBase : Measurement, IAngleProvider
    {
        private bool radians;

        [PropertyGridVisible]
        public bool Radians
        {
            get { return radians; }
            set 
            { 
                radians = value;
                UpdateVisual();
            }
        }

        public double Measure
        {
            get
            {
                var measure = Math.OAngle(Point(1), Point(0), Point(2));
                return (Radians) ? measure : measure.ToDegrees();
            }
        }

        public double Angle
        {
            get
            {
                return Math.OAngle(Point(1), Point(0), Point(2));
            }
        }

        public override void MoveToCore(Point newPosition)
        {
            Point newOffset = newPosition.Minus(Point(0));
            Offset = newOffset;
            base.MoveToCore(newPosition);
        }

        public override void UpdateVisual()
        {
            var p = Point(0).Plus(Offset);
            MoveToCore(p);
            base.UpdateVisual();
            var text = Math.Round(Measure, DecimalsToShow).ToString();
            Text = (Radians) ? text + " rad" : text + "°";
        }
    }

    public class AngleMeasurement : AngleMeasurementBase
    {
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
    }

    public class HorizontalAngleMeasurement : AngleMeasurementBase
    {
        public override void MoveToCore(Point newPosition)
        {
            base.MoveToCore(newPosition.Plus(0.2));
        }

        public override void UpdateVisual()
        {
            var p = Point(0).Plus(Offset);
            MoveToCore(p);
            Shape.CenterAt(ToPhysical(Coordinates));
            Text = Math
                .OHAngle(Point(0), Point(1))
                .ToDegrees()
                .ToDegreeString();
        }
    }
}
