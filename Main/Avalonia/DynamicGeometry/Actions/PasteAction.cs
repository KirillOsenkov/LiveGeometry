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
            // the labels of the pasted points are among the figures: the points must not
            // make or bring back labels of their own
            bool suppressed = PointBase.SuppressAutoLabelPoints;
            PointBase.SuppressAutoLabelPoints = true;
            try
            {
                Drawing.Figures.Add(Figures.ToArray<IFigure>());
            }
            finally
            {
                PointBase.SuppressAutoLabelPoints = suppressed;
            }

            // A copy of segment AB could not be AB again and was numbered by its type
            // (Segment1) before it was in the drawing; now that it is, it takes the name of
            // its own points, as a figure read from a file does.
            foreach (var figure in Figures)
            {
                (figure as FigureBase)?.UpdateDefaultName();
            }
        }

        protected override void UnExecuteCore()
        {
            Drawing.Figures.Remove(Figures);
        }
    }
}
