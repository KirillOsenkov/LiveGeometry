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
            var copied = XElement.Parse(copiedFigures);
            var deserializer = new DrawingDeserializer();
            BringStyles(copied, deserializer);
            List<IFigure> list = new List<IFigure>();
            deserializer.ReadFigureList(list, copied, Drawing, byFileName);
            Figures = list.ToArray();

            // nothing to paste, and no undo step: the styles go again
            if (list.Count == 0)
            {
                foreach (var style in addedStyles)
                {
                    Drawing.StyleManager.Withdraw(style);
                }
            }
        }

        public IEnumerable<IFigure> Figures { get; set; }

        // the styles of the copies that the drawing didn't have, in it while the copies are
        readonly List<IFigureStyle> addedStyles = new List<IFigureStyle>();

        /// <summary>
        /// The styles the copied figures name (<see cref="DrawingSerializer.WriteFiguresWithStyles"/>):
        /// one the drawing has, looking the same, is taken as it is; one it lacks comes along;
        /// one whose name the drawing gives to another look comes along under a free name,
        /// which the copies are made to name. In the drawing before the figures are read,
        /// since a figure looks its style up as it is read.
        /// </summary>
        void BringStyles(XElement copied, DrawingDeserializer deserializer)
        {
            var styles = copied.Element("Styles");
            var figures = copied.Name == "Drawing" ? copied.Element("Figures") : copied;
            if (styles == null || figures == null)
            {
                return;
            }

            var manager = Drawing.StyleManager;
            foreach (var styleNode in styles.Elements())
            {
                var style = deserializer.ReadKnownStyle(styleNode);
                if (style == null || style.Name.IsEmpty())
                {
                    continue;
                }

                var existing = manager[style.Name];
                if (existing != null && existing.GetType() == style.GetType() && existing.GetSignature() == style.GetSignature())
                {
                    continue;
                }

                if (existing != null)
                {
                    string oldName = style.Name;
                    style.Name = manager.FreeName(oldName);
                    foreach (var attribute in figures.Descendants().Attributes("Style").Where(a => a.Value == oldName))
                    {
                        attribute.Value = style.Name;
                    }
                }

                manager.Add(style);
                addedStyles.Add(style);
            }
        }

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
            // (on redo: they left with the copies)
            foreach (var style in addedStyles)
            {
                if (!Drawing.StyleManager.GetAllStyles().Contains(style))
                {
                    Drawing.StyleManager.Add(style);
                }
            }

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
            foreach (var style in addedStyles)
            {
                Drawing.StyleManager.Withdraw(style);
            }
        }
    }
}
