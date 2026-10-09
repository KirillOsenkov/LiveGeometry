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

        protected override string Kind
        {
            get
            {
                return "Angle";
            }
        }

        /// <summary>"ABC": "Angle ABC"</summary>
        public override string Construction
        {
            get
            {
                return Dependencies.Count < 3 ? null : ConstructionText.Points(Dependencies[1], Dependencies[0], Dependencies[2]);
            }
        }

        public virtual double Measure
        {
            get
            {
                var measure = Math.OAngle(Point(1), Point(0), Point(2));
                return (Radians) ? measure : measure.ToDegrees();
            }
        }

        public virtual double Angle
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

        public override void UpdateExistence()
        {
            base.UpdateExistence();
            if (Exists && !AngleArc.HasSides(this))
            {
                Exists = false;
            }
        }

        public override void UpdateVisual()
        {
            base.UpdateVisual();
            var text = Math.Round(Measure, DecimalsToShow).ToString();
            Text = (Radians) ? text + " rad" : text + "°";
        }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            base.ReadXml(element);
            radians = element.ReadBool("Radians", false);
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            if (Radians)
            {
                writer.WriteAttributeBool("Radians", true);
            }
        }
    }

    public class AngleMeasurement : AngleMeasurementBase, IConditionalProperties
    {
        /// <summary>Without its arc (deleted) the number has no arcs to count</summary>
        public bool CanEdit(string propertyName)
        {
            return propertyName != nameof(ArcCount) || FindArc() != null;
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            return defaultCaption;
        }

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
        [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
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
        [PropertyGridIcon(PropertyGridIcon.Angle)]
        public void ConvertToOpposite()
        {
            // the arc goes along
            AngleArc.ConvertToOpposite(this);
        }
    }

    public class HorizontalAngleMeasurement : AngleMeasurementBase
    {
        /// <summary>"of AB to the x axis"</summary>
        public override string Construction
        {
            get
            {
                return Dependencies.Count < 2 ? null : "of " + ConstructionText.Points(Dependencies[0], Dependencies[1]) + " to the x axis";
            }
        }

        /// <summary>The angle of the segment to the x axis, counterclockwise, in radians (the base reads a third point)</summary>
        public override double Angle
        {
            get
            {
                return Math.OHAngle(Point(0), Point(1));
            }
        }

        public override double Measure
        {
            get
            {
                return Radians ? Angle : Angle.ToDegrees();
            }
        }

        public override void UpdateVisual()
        {
            if (Dependencies.IsEmpty())
            {
                return;
            }

            // placed by its top left corner as every label is (ControlBase.UpdateVisual):
            // it centered the shape, which has no Width to halve, and the shape stayed in the
            // top left corner of the window
            Coordinates = PlaceFromOffset();
            Shape.MoveTo(ToPhysical(Coordinates));
            Text = Angle
                .ToDegrees()
                .ToDegreeString();
        }
    }
}
