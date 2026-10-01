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
            new DrawingDeserializer().ReadFigureList(list, XElement.Parse(copiedFigures), Drawing, byFileName);
            Figures = list.ToArray();
        }

        public IEnumerable<IFigure> Figures { get; set; }

        // the copies by the names the clipboard gives them, which are the originals' names
        readonly Dictionary<string, IFigure> byFileName = new Dictionary<string, IFigure>();

        bool expressionsRebound;

        /// <summary>
        /// The expressions of the copies (a label's [A.X], a graph, a point by coordinates)
        /// name the copies, under the names they have now. Read from the clipboard, they
        /// named the originals - the copy of a label measured the original segment, not the
        /// copy beside it - since they are compiled as they are read, when the drawing has
        /// only the originals by those names.
        /// </summary>
        void RebindExpressions()
        {
            var oldNames = new Dictionary<IFigure, string>();
            foreach (var pair in byFileName)
            {
                oldNames[pair.Value] = pair.Key;
            }

            var renamer = new ExpressionRenamer(Drawing, oldNames, preferred: Figures.ToArray());
            var holders = Figures.OfType<IRenamableExpressions>().ToArray();
            foreach (var holder in holders)
            {
                holder.RenameInExpressions(renamer);
            }

            foreach (var holder in holders)
            {
                holder.RebindExpressions();
            }
        }

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

            // once: on redo the same copies come back with their texts as they were left
            if (!expressionsRebound)
            {
                expressionsRebound = true;
                RebindExpressions();
            }
        }

        protected override void UnExecuteCore()
        {
            Drawing.Figures.Remove(Figures);
        }
    }
}
