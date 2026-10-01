using System.Collections.Generic;
using System.Linq;
using Avalonia;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    /// <summary>
    /// A move by an offset. Undo and redo don't move back and forth by the offset: a movable
    /// that can say where it is (<see cref="IRestorablePlace"/>) is put exactly where it
    /// was before and after. A point kept on a circle or behind a stop doesn't land where the
    /// offset says, and moving it back by the same offset left it somewhere else; a label
    /// sits a number of pixels from its anchor, which an offset in the plane only matches
    /// until the view is zoomed.
    /// </summary>
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

        IMovable[] movables;

        // where each movable was before the move and where it got to; null for one that
        // can't say (the coordinate system, which only knows offsets)
        object[] before;
        object[] after;

        protected override void ExecuteCore()
        {
            if (after == null)
            {
                movables = Points.ToArray();
                before = CapturePlaces();
                movables.Move(Offset);
                after = CapturePlaces();
            }
            else
            {
                // redo
                RestorePlaces(after, Offset);
            }

            Recalculate(Drawing, ToRecalculate);
        }

        object[] CapturePlaces()
        {
            var places = new object[movables.Length];
            for (int i = 0; i < movables.Length; i++)
            {
                places[i] = (movables[i] as IRestorablePlace)?.CapturePlace();
            }

            return places;
        }

        void RestorePlaces(object[] places, Point offsetForTheRest)
        {
            for (int i = 0; i < movables.Length; i++)
            {
                if (places[i] != null)
                {
                    ((IRestorablePlace)movables[i]).RestorePlace(places[i]);
                }
                else
                {
                    movables[i].MoveTo(movables[i].Coordinates.Plus(offsetForTheRest));
                }
            }
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
            RestorePlaces(before, Offset.Minus());
            Recalculate(Drawing, ToRecalculate);
        }

        public override bool TryToMerge(IAction followingAction)
        {
            MoveAction next = followingAction as MoveAction;
            if (next == null || next.Points != this.Points || movables == null)
            {
                return false;
            }

            // a merged action leaves what there is to redo in the history: a new drag must
            // end that, so its first move is an action of its own
            if (Drawing != null && Drawing.ActionManager != null && Drawing.ActionManager.CanRedo)
            {
                return false;
            }

            movables.Move(next.Offset);
            Offset = Offset.Plus(next.Offset);
            after = CapturePlaces();
            Recalculate(Drawing, ToRecalculate);
            return true;
        }
    }
}
