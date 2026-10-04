using Avalonia;
using System.Xml;
using System.Xml.Linq;

namespace DynamicGeometry
{
    public interface ILengthProvider : IFigure
    {
        double Length { get; }
    }

    public class DistanceMeasurement : Measurement, ILengthProvider
    {
        public override Point Anchor
        {
            get
            {
                return Midpoint();
            }
        }

        protected override string Kind
        {
            get
            {
                return "Distance";
            }
        }

        /// <summary>"AB": "Distance AB"</summary>
        public override string Construction
        {
            get
            {
                if (Dependencies.Count > 1 && Dependencies[0] is IPoint && Dependencies[1] is IPoint)
                {
                    return ConstructionText.Points(Dependencies[0], Dependencies[1]);
                }

                return Dependencies.Count > 0 ? ConstructionText.Length(Dependencies[0]) : null;
            }
        }

        /// <summary>
        /// A figure with a length, or two points. Anything else has nothing to measure (a
        /// segment that was replaced by a ray, a file that says so): the measurement then
        /// doesn't exist, where it used to throw on every redraw.
        /// </summary>
        bool HasSomethingToMeasure
        {
            get
            {
                return Dependencies.Count > 0
                    && (Dependencies[0] is ILengthProvider
                        || Dependencies.Count > 1 && Dependencies[0] is IPoint && Dependencies[1] is IPoint);
            }
        }

        public override void UpdateExistence()
        {
            base.UpdateExistence();
            if (Exists && !HasSomethingToMeasure)
            {
                Exists = false;
            }
        }

        public Point Midpoint()
        {
            if (!HasSomethingToMeasure)
            {
                return Math.InfinitePoint;
            }

            if (Dependencies[0] is ILengthProvider)
            {
                return Dependencies[0].Center;
            }
            return Math.Midpoint(Point(0), Point(1));
        }

        public double Distance
        {
            get
            {
                if (!HasSomethingToMeasure)
                {
                    return double.NaN;
                }

                if (Dependencies[0] is ILengthProvider)
                {
                    return (Dependencies[0] as ILengthProvider).Length;
                }
                return Point(0).Distance(Point(1));
            }
        }

        public double Measure
        {
            get
            {
                return Distance * ConversionFactor;
            }
        }

        public double Length
        {
            get
            {
                return Distance;
            }
        }

        public override void UpdateVisual()
        {
            if (!HasSomethingToMeasure)
            {
                return;
            }

            base.UpdateVisual();

            //Text = Math.Round(Distance,DecimalsToShow).ToString();
            var distance = Math.Round(Measure, DecimalsToShow).ToString();
            if (Units == Math.lengthUnit.Inches) Text = distance + "\"";
            else if (Units == Math.lengthUnit.Centimeter) Text = distance + "cm";
            else Text = distance;

        }

        private Math.lengthUnit mUnits = Settings.Instance.DistanceUnit;
        [PropertyGridVisible]
        public Math.lengthUnit Units
        {
            get
            {
                return mUnits;
            }
            set
            {
                mUnits = value;
                UpdateVisual();
            }
        }

        double ConversionFactor
        {
            get
            {
                if (Units == Math.lengthUnit.Inches) return 1 / Math.inchesLogicalLength;
                if (Units == Math.lengthUnit.Centimeter) return 1 / Math.centimeterLogicalLength;
                return 1;
            }
        }

#if !PLAYER

        public override void WriteXml(XmlWriter writer)
        {
            base.WriteXml(writer);
            if (Units == Math.lengthUnit.Inches)
            {
                writer.WriteAttributeString("Units", "Inches");
            }
            else if (Units == Math.lengthUnit.Centimeter)
            {
                writer.WriteAttributeString("Units", "Centimeters");
            }
        }

#endif

        public override void ReadXml(XElement element)
        {
            base.ReadXml(element);
            var unitsAsString = element.ReadString("Units");
            if (unitsAsString == "Inches")
            {
                Units = Math.lengthUnit.Inches;
            }
            else if (unitsAsString == "Centimeters")
            {
                Units = Math.lengthUnit.Centimeter;
            }
        }

    }
}
