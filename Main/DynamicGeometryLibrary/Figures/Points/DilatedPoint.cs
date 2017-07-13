using System.Linq;
using System.Windows.Shapes;
using System;

namespace DynamicGeometry
{
    public class DilatedPoint : PointBase, IPoint
    {

        //[PropertyGridVisible]
        [PropertyGridName("Dilation Of ")]
        public IPoint Source
        {
            get
            {
                return (Dependencies.Count >= 1) ? Dependencies.ElementAt(0) as IPoint : null;
            }
        }

        //[PropertyGridVisible]
        [PropertyGridName("About Point ")]
        public new IPoint Center
        {
            get
            {
                return (Dependencies.Count >= 2) ? Dependencies.ElementAt(1) as IPoint : null;
            }
        }

        //[PropertyGridVisible]
        public double Factor_Numerator
        {
            get
            {
                return (Dependencies.Count >= 3) ? (Dependencies.ElementAt(2) as ILengthProvider).Length : 1;
            }
        }

        //[PropertyGridVisible]
        public double Factor_Denominator
        {
            get
            {
                return (Dependencies.Count >= 4) ? (Dependencies.ElementAt(3) as ILengthProvider).Length : 1;
            }
        }

        double factor;
        [PropertyGridVisible]
        public double Factor
        {
            get 
            {
                if (Dependencies.Count == 3 && Dependencies[2] is ILengthProvider)
                {
                    return (Dependencies[2] as ILengthProvider).Length;
                }
                else if (Dependencies.Count > 3)
                {
                    var denominator = Factor_Denominator;
                    if (denominator != 0)
                    {
                        return Factor_Numerator / denominator;
                    }
                    else
                    {
                        return 99999;    // double.MaxValue causes error.
                    }
                }
                else
                {
                    return factor;
                }
            }
            set 
            { 
                factor = value;
                Recalculate();
                this.RecalculateAllDependents();
                UpdateVisual();
            }
        }

        protected override Shape CreateShape()
        {
            return Factory.CreateDependentPointShape();
        }

        public override void Recalculate()
        {
            if (Source != null && Center != null)
            {
                Coordinates = Math.GetDilationPoint(Source.Coordinates, Center.Coordinates,Factor);
            }
            Exists = Coordinates.Exists();
        }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            base.ReadXml(element);
            Factor = element.ReadDouble("Factor");
            Recalculate();
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            writer.WriteAttributeDouble("Factor", Factor);
        }

    }
}