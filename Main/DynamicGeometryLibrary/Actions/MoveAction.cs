using System.Collections.Generic;
using System.Linq;
using System.Windows;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    public class MoveAction : GeometryAction
    {
        public MoveAction(
            Drawing drawing,
            IEnumerable<IMovable> points,
            Point offset,
            IEnumerable<IFigure> toRecalculate)
            : base(drawing)
        {
            Points = points;
            Offset = offset;
            ToRecalculate = toRecalculate;
        }

        public IEnumerable<IFigure> ToRecalculate { get; set; }
        public IEnumerable<IMovable> Points { get; set; }
        public Point Offset { get; set; }

        protected override void ExecuteCore()
        {
            Points.Move(Offset);
            Recalculate(Drawing, ToRecalculate);
        }

        public static void Recalculate(Drawing drawing, IEnumerable<IFigure> toRecalculate)
        {
            if (toRecalculate != null)
            {
                var list = toRecalculate.ToList();

                foreach (var figure in toRecalculate)
                {
                    // need to check because Recalculate() of a previous figure (polygon) might have deleted this one from the drawing
                    if (figure.Drawing != null)
                    {
                        figure.RecalculateAndUpdateVisual();
                    }
                    else
                    {
                        list.Remove(figure);
                    }
                }

                if (drawing != null)
                {
                    drawing.RaiseFigureCoordinatesChanged(
                        new Drawing.FigureCoordinatesChangedEventArgs(
                            list));
                }
            }
        }

        protected override void UnExecuteCore()
        {
            Points.Move(Offset.Minus());
            Recalculate(Drawing, ToRecalculate);
        }

        public override bool TryToMerge(IAction followingAction)
        {
            MoveAction next = followingAction as MoveAction;
            if (next != null
                && next.Points == this.Points)
            {
                Points.Move(next.Offset);
                Offset = Offset.Plus(next.Offset);
                Recalculate(Drawing, ToRecalculate);
                return true;
            }

            return false;
        }
    }
}
