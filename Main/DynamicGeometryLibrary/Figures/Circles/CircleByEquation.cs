using System.Windows;

namespace DynamicGeometry
{
    public class CircleByEquation : CircleBase, ICircle, IShapeWithInterior
    {
        [PropertyGridVisible(false)]
        public override Point Center
        {
            get 
            {
                if (X.Value == null || Y.Value == null)
                {
                    return new Point();
                }

                return new Point(X.Value(), Y.Value());
            }
        }

        [PropertyGridVisible(false)]
        public override double Radius
        {
            get 
            {
                if (R.Value == null)
                {
                    return 0;
                }
                var radius = R.Value();
                return radius > 0 ? radius : Math.Epsilon;
            }
        }

        [PropertyGridVisible]
        public DrawingExpression X { get; set; }

        [PropertyGridVisible]
        public DrawingExpression Y { get; set; }

        [PropertyGridVisible]
        public DrawingExpression R { get; set; }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            X = new DrawingExpression(this, "Center X =", element.ReadString("X"));
            Y = new DrawingExpression(this, "Center Y =", element.ReadString("Y"));
            R = new DrawingExpression(this, "Radius =", element.ReadString("R"));
            base.ReadXml(element);
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            writer.WriteAttributeString("X", X.Text);
            writer.WriteAttributeString("Y", Y.Text);
            writer.WriteAttributeString("R", R.Text);
        }
    }
}