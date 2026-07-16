using System;
using System.Collections.Generic;
using System.Linq;

namespace DynamicGeometry
{
    public static partial class DependencyAlgorithms
    {
        public static List<T> TopologicalSort<T>(
            this IEnumerable<T> originalSet,
            Func<T, IEnumerable<T>> childrenSelector)
        {
            Dictionary<T, bool> visitedSet = new Dictionary<T, bool>();
            List<T> finalOrder = new List<T>();

            AddAllDependents(
                originalSet,
                childrenSelector,
                visitedSet,
                finalOrder.Add);

            return finalOrder;
        }

        private static void AddAllDependents<T>(
            IEnumerable<T> nodes,
            Func<T, IEnumerable<T>> childrenSelector,
            Dictionary<T, bool> visitedSet,
            Action<T> resultCollector)
        {
            foreach (var node in nodes)
            {
                if (!visitedSet.ContainsKey(node))
                {
                    visitedSet.Add(node, true);
                    var children = childrenSelector(node);
                    if (!children.IsEmpty())
                    {
                        AddAllDependents(
                            children,
                            childrenSelector,
                            visitedSet,
                            resultCollector);
                    }

                    if (resultCollector != null)
                    {
                        resultCollector(node);
                    }
                }
            }
        }

        public static List<T> FindDescendants<T>(Func<T, IEnumerable<T>> childrenSelector, IEnumerable<T> list)
        {
            return TopologicalSort(list, childrenSelector);
        }

        public static IEnumerable<T> FindDescendants<T>(Func<T, IEnumerable<T>> childrenSelector, T node)
        {
            return FindDescendants(childrenSelector, node.AsEnumerable());
        }

        public static IEnumerable<T> FindRoots<T>(Func<T, IEnumerable<T>> childrenSelector, params T[] list)
        {
            return FindRoots(childrenSelector, (IEnumerable<T>)list);
        }

        public static IEnumerable<T> FindRoots<T>(Func<T, IEnumerable<T>> childrenSelector, IEnumerable<T> list)
        {
            List<T> result = new List<T>();
            foreach (var figure in list)
            {
                FindRoots(childrenSelector, figure, result.Add);
            }

            return result.Distinct();
        }

        /// <summary>
        /// Doesn't use recursion because we can hit StackOverflow for deep figures
        /// </summary>
        public static void FindRoots<T>(Func<T, IEnumerable<T>> childrenSelector, T figure, Action<T> collector)
        {
            Stack<T> stack = new Stack<T>();
            stack.Push(figure);
            while (!stack.IsEmpty())
            {
                if (stack.Count > 10000)
                {
                    throw new InvalidOperationException("Weird, we hit a cycle in a DAG, need to investigate this bug");
                }
                figure = stack.Pop();
                var children = childrenSelector(figure);
                if (children.IsEmpty())
                {
                    collector(figure);
                }
                else
                {
                    foreach (var dependency in children.Reverse())
                    {
                        stack.Push(dependency);
                    }
                }
            }
        }

        public static bool FigureCompletelyDependsOnFigures(IFigure figure, IEnumerable<IFigure> set)
        {
            if (set.Contains(figure))
            {
                return true;
            }
            if (figure.Dependencies.IsEmpty())
            {
                return false;
            }
            foreach (var dependency in figure.Dependencies)
            {
                var result = FigureCompletelyDependsOnFigures(dependency, set);
                if (!result)
                {
                    return false;
                }
            }
            return true;
        }

        public static List<IFigure> FindImpactedDependencyChain(IFigure source, IFigure sink)
        {
            var visitedSet = new Dictionary<IFigure, bool>();

            AddAllDependents(
                source.AsEnumerable(),
                f => f.Dependents,
                visitedSet,
                null);

            var result = new List<IFigure>();
            AddImpactedDependency(sink, visitedSet, result);
            return result;
        }

        private static void AddImpactedDependency(IFigure sink, Dictionary<IFigure, bool> candidates, List<IFigure> result)
        {
            if (!candidates.ContainsKey(sink))
            {
                return;
            }

            foreach (var dependency in sink.Dependencies)
            {
                AddImpactedDependency(dependency, candidates, result);
            }

            result.Add(sink);
            candidates.Remove(sink);
        }
    }
}
