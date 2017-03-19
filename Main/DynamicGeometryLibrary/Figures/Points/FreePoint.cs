using System.Windows;
using System.Xml.Linq;

namespace DynamicGeometry
{
    public partial class FreePoint : PointBase, IMovable
    {
        public override void ReadXml(XElement element)
        {
            base.ReadXml(element);
            var x = element.ReadDouble("X");
            var y = element.ReadDouble("Y");
            Coordinates = new Point(x, y);
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            var coordinates = Coordinates;
            writer.WriteAttributeDouble("X", coordinates.X);
            writer.WriteAttributeDouble("Y", coordinates.Y);
        }

        /// <summary>
        /// Perf optimization
        /// We know it exists, no need to call base
        /// </summary>
        public override void UpdateExistence()
        {
        }

        [PropertyGridVisible]
        public override double X
        {
            get
            {
                return base.X;
            }
            set
            {
                this.MoveTo(new Point(value, Y));
                this.RecalculateAllDependents();
            }
        }

        [PropertyGridVisible]
        public override double Y
        {
            get
            {
                return base.Y;
            }
            set
            {
                this.MoveTo(new Point(X, value));
                this.RecalculateAllDependents();
            }
        }
    }
}
