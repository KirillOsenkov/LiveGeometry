using System.Linq;
using System.Windows.Shapes;

namespace DynamicGeometry
{
    public class TranslatedPoint : PointBase, IPoint
    {
        //[PropertyGridVisible]
        [PropertyGridName("Translation of ")]
        public IPoint Source {
            get
            {
                return (Dependencies.Count >= 1) ? Dependencies.ElementAt(0) as IPoint : null;
            }
        }

        //[PropertyGridVisible]
        [PropertyGridName("Using Vector ")]
        public Vector Vector
        {
            get
            {
                return (Dependencies.Count >= 2) ? Dependencies.ElementAt(1) as Vector : null;
            }
        }
        double magnitude;

        [PropertyGridVisible]
        public double Magnitude
        {
            get 
            {
                if (Dependencies.Count > 1 && Dependencies[1] is Vector)
                {
                    return (Dependencies[1] as Vector).Magnitude;
                }
                else if (Dependencies.Count > 1 && Dependencies[1] is ILengthProvider)
                {
                    return (Dependencies[1] as ILengthProvider).Length;
                }
                else
                {
                    return magnitude;
                }
            }
            set 
            {
                magnitude = value;
                Recalculate();
                this.RecalculateAllDependents();
                UpdateVisual();
            }
        }
        double direction;

        [PropertyGridVisible]
        public double Direction
        {
            get
            {
                if (Dependencies.Count > 1 && Dependencies[1] is Vector)
                {
                    return (Dependencies[1] as Vector).Direction;
                }
                else if (Dependencies.Count > 2 && Dependencies[2] is IAngleProvider)
                {
                    return (Dependencies[2] as IAngleProvider).Angle;
                }
                else
                {
                    return direction;
                }
            }
            set
            {
                direction = Math.ToRadians(value);
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
            if (Source != null)
            {
                Coordinates = Math.GetTranslationPoint(Source.Coordinates, Magnitude, Direction);
            }
            Exists = Coordinates.Exists();
        }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            base.ReadXml(element);
            Magnitude = element.ReadDouble("Magnitude");
            Direction = element.ReadDouble("Direction");
            Recalculate();
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            writer.WriteAttributeDouble("Magnitude", Magnitude);
            writer.WriteAttributeDouble("Direction", Direction);
        }

    }
}