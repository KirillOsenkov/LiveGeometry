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
            CustomDependencyRemovers = new List<IAction>();

            var deleted = Figure.AsEnumerable<IFigure>()
                .TopologicalSort(GetRemovableDependencies);

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

        private IEnumerable<IFigure> GetRemovableDependencies(IFigure figure)
        {
            var list = figure.Dependents.ToList();

            foreach (var item in figure.Dependents)
            {
                if (item is ISupportRemoveDependency customDependencyRemover
                    && customDependencyRemover.CanRemoveDependency(figure))
                {
                    list.Remove(item);
                    CustomDependencyRemovers.Add(customDependencyRemover.GetRemoveDependencyAction(figure));
                }
            }

            return list;
        }

        /// <summary>
        /// A figure created on demand for another one (a Number holding a typed length) is
        /// auxiliary: it goes when its last user goes. Transitively, in case an auxiliary
        /// figure has auxiliary dependencies of its own.
        /// </summary>
        public static List<IFigure> FindOrphanedAuxiliaries(IEnumerable<IFigure> dying)
        {
            var gone = new HashSet<IFigure>(dying);
            var orphans = new List<IFigure>();
            var toVisit = new Queue<IFigure>(dying);
            while (toVisit.Count > 0)
            {
                var figure = toVisit.Dequeue();
                foreach (var dependency in figure.Dependencies)
                {
                    if (dependency.Auxiliary
                        && !gone.Contains(dependency)
                        && dependency.Dependents.All(gone.Contains))
                    {
                        gone.Add(dependency);
                        orphans.Add(dependency);
                        toVisit.Enqueue(dependency);
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
