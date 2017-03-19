using System.Windows;
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

        public override void MoveToCore(Point newPosition)
        {
            Offset = newPosition.Minus(Midpoint());
            base.MoveToCore(newPosition);
        }

        public Point Midpoint()
        {
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
            var p = Midpoint().Plus(Offset);
            MoveToCore(p);
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
