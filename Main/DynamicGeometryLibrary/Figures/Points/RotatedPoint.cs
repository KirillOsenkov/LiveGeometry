using System.Linq;
using System.Windows.Shapes;
using System.Windows;

namespace DynamicGeometry
{
    public class RotatedPoint : PointBase, IPoint
    {
        //[PropertyGridVisible]
        [PropertyGridName("Rotation of ")]
        public IPoint Source
        {
            get
            {
                return (Dependencies.Count >= 1) ? Dependencies.ElementAt(0) as IPoint : null;
            }
        }

        //[PropertyGridVisible]
        [PropertyGridName("About Center ")]
        public new IPoint Center
        {
            get
            {
                return (Dependencies.Count >= 2) ? Dependencies.ElementAt(1) as IPoint : null;
            }
        }

        double angle;
        [PropertyGridVisible]
        [PropertyGridName("Angle ")]
        public double Angle
        {
            get 
            {
                if (Dependencies.Count >= 3)
                {
                    var angleProvider = Dependencies.ElementAt(2) as IAngleProvider;
                    if (angleProvider != null)
                    {
                        return angleProvider.Angle.ToDegrees();
                    }
                    return 0;
                }
                else
                {
                    return angle;
                }
            }
            set 
            {
                if (angle != value)
                {
                    angle = value;
                    Recalculate();
                    this.RecalculateAllDependents();
                    UpdateVisual();
                }
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
                Coordinates = Math.GetRotationPoint(Source.Coordinates, Center.Coordinates, Math.ToRadians(Angle));
            }
            Exists = Coordinates.Exists();
        }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            base.ReadXml(element);
            angle = element.ReadDouble("Angle");
            Recalculate();
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            writer.WriteAttributeDouble("Angle", angle);
        }

    }
}