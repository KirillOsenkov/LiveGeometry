namespace DynamicGeometry
{
    public class PointByCoordinates : PointBase, IPoint
    {
        public PointByCoordinates()
        {
            XExpression = new DrawingExpression(this) { Name = "X = " };
            YExpression = new DrawingExpression(this) { Name = "Y = " };
        }

        [PropertyGridVisible]
        [PropertyGridName("X = ")]
        public DrawingExpression XExpression { get; private set; }
        
        [PropertyGridVisible]
        [PropertyGridName("Y = ")]
        public DrawingExpression YExpression { get; private set; }

        public override void Recalculate()
        {
            if (XExpression == null 
                || XExpression.Value == null 
                || YExpression == null
                || YExpression.Value == null)
            {
                return;
            }
            Coordinates = new System.Windows.Point(XExpression.Value(), YExpression.Value());
            Exists = Coordinates.Exists();
        }

        public override void OnAddingToDrawing(Drawing drawing)
        {
            // Recalculate in order to compile expressions and have accurate coordinates.  
            // Needed when autoLabelPoints is on. -D.H.
            Recalculate();
            base.OnAddingToDrawing(drawing);
        }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            base.ReadXml(element);
            XExpression.Text = element.ReadString("X");
            YExpression.Text = element.ReadString("Y");
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            writer.WriteAttributeString("X", XExpression.Text);
            writer.WriteAttributeString("Y", YExpression.Text);
        }
    }
}
