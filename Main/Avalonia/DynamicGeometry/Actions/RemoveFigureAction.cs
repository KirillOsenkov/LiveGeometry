using System.Collections.Generic;
using System.Linq;
using GuiLabs.Undo;
using System;

namespace DynamicGeometry
{
    public class RemoveFigureAction : GeometryAction
    {
        public RemoveFigureAction(Drawing drawing, IFigure figure)
            : base(drawing)
        {
            Figure = figure;
        }

        private IFigure Figure;
        private IFigure[] Deleted;
        private IList<IAction> CustomDependencyRemovers;

        // where each deleted figure was in the drawing's list, so that undo puts it back
        // there (the Figure List, the saved file) and not at the end; -1 if it wasn't there
        private int[] Indices;

        protected override void ExecuteCore()
        {
            // What goes is decided against the whole deletion: a figure that can do without
            // some of what it is built on (ISupportRemoveDependency) is spared only when it
            // can do without everything the deletion takes from it, and what is built on it
            // is spared with it. Asked for one dependency at a time, a show/hide box holding
            // A and segment AB said yes to each when A was deleted, was left holding nothing,
            // and threw on removing AB at the index it had had before A went. Everything goes
            // at first, and the figures that can keep going are spared round by round, until
            // no more can (each round only takes less away).
            var spared = new List<IFigure>();
            List<IFigure> deleted;
            while (true)
            {
                deleted = Figure.AsEnumerable<IFigure>()
                    .TopologicalSort(figure => figure.Dependents.Where(dependent => !spared.Contains(dependent)));
                var gone = new HashSet<IFigure>(deleted);
                var keeping = deleted
                    .Where(figure => figure != Figure
                        && figure is ISupportRemoveDependency remover
                        && remover.CanRemoveDependencies(gone))
                    .ToList();
                if (keeping.Count == 0)
                {
                    break;
                }

                spared.AddRange(keeping);
            }

            // all made before any runs: each finds its dependency when it runs
            var removed = new HashSet<IFigure>(deleted);
            CustomDependencyRemovers = new List<IAction>();
            foreach (var figure in spared)
            {
                foreach (var dependency in figure.Dependencies.Where(removed.Contains).Distinct().ToList())
                {
                    CustomDependencyRemovers.Add(((ISupportRemoveDependency)figure).GetRemoveDependencyAction(dependency));
                }
            }

            // The parts of a composite (the vertices a regular polygon works out, which are
            // built on the polygon) are among the dependents so that what is built on them
            // goes too. They are not figures of the drawing: they leave and come back with
            // their composite, which sees to their dependencies. (Put back as figures on
            // undo, they ended up in the drawing's own list, and on the canvas twice.)
            deleted.RemoveAll(figure => figure != Figure && !Drawing.Figures.Contains(figure));
            // after their dependents, so that undo brings them back first
            deleted.AddRange(FindOrphanedAuxiliaries(deleted));
            Deleted = deleted.ToArray();

            foreach (var item in CustomDependencyRemovers)
            {
                item.Execute();
            }

            Indices = new int[Deleted.Length];
            for (int i = 0; i < Deleted.Length; i++)
            {
                var item = Deleted[i];
                Indices[i] = Drawing.Figures.IndexOf(item);
                if (Indices[i] >= 0)
                {
                    Drawing.Figures.RemoveAt(Indices[i]);
                }
                else
                {
                    item.UnregisterFromDependencies();
                }
            }

            // Not for what a tool takes away of its own (the point following the cursor, its
            // preview), which is never recorded: the side panel then showed whatever had been
            // selected before the construction began, every time one was given up.
            if (!Drawing.ActionManager.ExecuteImmediatelyWithoutRecording || Deleted.Any(FigureParts.HasSelection))
            {
                Drawing.RaiseSelectionChanged(new Drawing.SelectionChangedEventArgs());
            }
        }

        /// <summary>
        /// A figure created on demand for another one (a Number holding a typed length) is
        /// auxiliary: it goes when its last user goes. Transitively, in case an auxiliary
        /// figure has auxiliary dependencies of its own. The two axis lines go together, once
        /// nothing is built on either, as they came (<see cref="AxisLine"/>) - also when one of
        /// them is what is deleted.
        /// </summary>
        public static List<IFigure> FindOrphanedAuxiliaries(IEnumerable<IFigure> dying)
        {
            var gone = new HashSet<IFigure>(dying);
            var orphans = new List<IFigure>();
            var toVisit = new Queue<IFigure>(dying);

            void Visit(IFigure orphan)
            {
                gone.Add(orphan);
                orphans.Add(orphan);
                toVisit.Enqueue(orphan);
            }

            void VisitAxes(Drawing drawing)
            {
                var axes = drawing.Figures.OfType<AxisLine>().Where(axis => !gone.Contains(axis)).ToList();
                if (axes.All(axis => axis.Dependents.All(gone.Contains)))
                {
                    axes.ForEach(Visit);
                }
            }

            var axisDrawing = gone.OfType<AxisLine>().FirstOrDefault()?.Drawing;
            if (axisDrawing != null)
            {
                VisitAxes(axisDrawing);
            }

            while (toVisit.Count > 0)
            {
                var figure = toVisit.Dequeue();
                foreach (var dependency in figure.Dependencies)
                {
                    if (dependency is AxisLine axis)
                    {
                        if (!gone.Contains(axis) && axis.Drawing != null)
                        {
                            VisitAxes(axis.Drawing);
                        }
                    }
                    else if (dependency.Auxiliary
                        && !gone.Contains(dependency)
                        && dependency.Dependents.All(gone.Contains))
                    {
                        Visit(dependency);
                    }
                }
            }

            return orphans;
        }

        protected override void UnExecuteCore()
        {
            // Suppress auto labeling to prevent duplicate labels. try-catch added to ensure suppression is always stopped.
            PointBase.SuppressAutoLabelPoints = true;
            try
            {
                for (int i = Deleted.Length - 1; i >= 0; i--)
                {
                    var index = Indices[i];
                    if (index >= 0 && index <= Drawing.Figures.Count)
                    {
                        Drawing.Figures.Insert(index, Deleted[i]);
                    }
                    else
                    {
                        Drawing.Figures.Add(Deleted[i]);
                    }
                }

                foreach (var item in CustomDependencyRemovers.Reverse())
                {
                    item.UnExecute();
                }
            }
            catch (Exception ex)
            {
                Drawing.RaiseError(this, ex);
            }
            PointBase.SuppressAutoLabelPoints = false;
        }
    }
}
