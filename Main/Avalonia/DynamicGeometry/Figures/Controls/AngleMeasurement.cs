using System.Collections.Generic;
using System.Linq;
using Avalonia;

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

        public override Point Anchor
        {
            get
            {
                return Point(0);
            }
        }

        public override void UpdateVisual()
        {
            base.UpdateVisual();
            var text = Math.Round(Measure, DecimalsToShow).ToString();
            Text = (Radians) ? text + " rad" : text + "°";
        }
    }

    public class AngleMeasurement : AngleMeasurementBase
    {
        /// <summary>
        /// The arc that was created together with this label: same vertex, same two sides.
        /// Null if it has been deleted.
        /// </summary>
        AngleArc FindArc()
        {
            return AngleArc.FindCompanion(this) as AngleArc;
        }

        /// <summary>
        /// The arcs belong to the <see cref="AngleArc"/>, a figure of its own. They can be set
        /// from the label too: with no arcs there is nothing left of the arc to click on.
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridName("Arcs")]
        [Domain(0, 3)]
        public int ArcCount
        {
            get
            {
                var arc = FindArc();
                return arc != null ? arc.ArcCount : 0;
            }
            set
            {
                var arc = FindArc();
                if (arc != null)
                {
                    arc.ArcCount = value;
                }
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert to opposite angle")]
        public void ConvertToOpposite()
        {
            // the arc goes along
            AngleArc.ConvertToOpposite(this);
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
            if (Dependencies.IsEmpty())
            {
                return;
            }

            Coordinates = PlaceFromOffset();
            Shape.CenterAt(ToPhysical(Coordinates));
            Text = Math
                .OHAngle(Point(0), Point(1))
                .ToDegrees()
                .ToDegreeString();
        }
    }
}
