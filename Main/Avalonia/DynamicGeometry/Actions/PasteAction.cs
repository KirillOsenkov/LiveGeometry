using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq; 

namespace DynamicGeometry
{
    public class PasteAction : GeometryAction
    {
        public PasteAction(Drawing drawing, string copiedFigures)
            : base(drawing)
        {
            List<IFigure> list = new List<IFigure>();
            new DrawingDeserializer().ReadFigureList(list, XElement.Parse(copiedFigures), Drawing);
            Figures = list.ToArray();
        }

        public IEnumerable<IFigure> Figures { get; set; }

        protected override void ExecuteCore()
        {
            Drawing.Figures.Add(Figures.ToArray<IFigure>());
        }

        protected override void UnExecuteCore()
        {
            Drawing.Figures.Remove(Figures);
        }
    }
}
