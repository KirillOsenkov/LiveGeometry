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

        protected override void ExecuteCore()
        {
            CustomDependencyRemovers = new List<IAction>();

            var deleted = Figure.AsEnumerable<IFigure>()
                .TopologicalSort(GetRemovableDependencies);
            // after their dependents, so that undo brings them back first
            deleted.AddRange(FindOrphanedAuxiliaries(deleted));
            Deleted = deleted.ToArray();

            foreach (var item in CustomDependencyRemovers)
            {
                item.Execute();
            }

            foreach (var item in Deleted)
            {
                if (!Drawing.Figures.Remove(item))
                {
                    item.UnregisterFromDependencies();
                }
            }

            Drawing.RaiseSelectionChanged(new Drawing.SelectionChangedEventArgs());
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
                foreach (var item in Deleted.Reverse())
                {
                    Drawing.Figures.Add(item);
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
