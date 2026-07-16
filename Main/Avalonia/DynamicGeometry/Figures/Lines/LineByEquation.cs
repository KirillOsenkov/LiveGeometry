namespace DynamicGeometry
{
    public class LineByEquation : LineBase, ILine
    {
        public override PointPair OnScreenCoordinates
        {
            get
            {
                return Math.GetLineFromSegment(Coordinates, CanvasLogicalBorders);
            }
        }

        public override PointPair Coordinates
        {
            get
            {
                return Equation.LineCoordinates;
            }
        }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            Equation = LineEquation.Read(this, element);
            base.ReadXml(element);
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            Equation.Write(writer);
        }

        [PropertyGridComplexTypeState(ComplexTypeState.Expanded)]
        [PropertyGridVisible]
        [PropertyGridName("Equation")]
        public ILineEquation Equation { get; set; }
    }
}
