using Avalonia;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace DynamicGeometry
{
    public class AreaMeasurement : Measurement
    {
        protected override string Kind
        {
            get
            {
                return "Area";
            }
        }

        /// <summary>"of triangle ABC", "of circle c", "of ABCD" (points)</summary>
        public override string Construction
        {
            get
            {
                if (Dependencies.Count == 0)
                {
                    return null;
                }

                return "of " + (Dependencies[0] is IShapeWithInterior
                    ? ConstructionText.Of(Dependencies[0])
                    : ConstructionText.Points(Dependencies.ToArray()));
            }
        }

        private Math.lengthUnit mUnits = Settings.Instance.DistanceUnit;
        [PropertyGridVisible]
        public Math.lengthUnit Units {
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
                if (Units == Math.lengthUnit.Inches) return 1 / Math.inchesLogicalLength.Sqr();
                if (Units == Math.lengthUnit.Centimeter) return 1 / Math.centimeterLogicalLength.Sqr();
                return 1;
            }
        }

        public double Measure
        {
            get
            {
                if (Dependencies[0] is IShapeWithInterior)
                {
                    return (Dependencies[0] as IShapeWithInterior).Area * ConversionFactor;
                }
                return Dependencies.ToPoints().Area() * ConversionFactor;
            }
        }

        public override Point Anchor
        {
            get
            {
                return Origin;
            }
        }

        public override void UpdateVisual()
        {
            base.UpdateVisual();
            var areaText = Math.Round(Measure,DecimalsToShow).ToString();
            if (Units == Math.lengthUnit.Inches) Text = areaText + "in²";
            else if (Units == Math.lengthUnit.Centimeter) Text = areaText + "cm²";
            else Text = areaText;
        }

        private Point Origin
        {
            get
            {
                if (Dependencies[0] is PointBase)
                {
                    return Dependencies.ToPoints().Midpoint();
                }
                else
                {
                    return Dependencies[0].Center;
                }               
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
