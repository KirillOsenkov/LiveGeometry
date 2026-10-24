using System;
using System.Collections.Generic;
using System.Linq;

namespace LiveGeometry;

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
        var visited = new HashSet<T>();
        foreach (var figure in list)
        {
            FindRoots(childrenSelector, figure, result.Add, visited);
        }

        return result.Distinct();
    }

    /// <summary>
    /// Doesn't use recursion because we can hit StackOverflow for deep figures
    /// </summary>
    public static void FindRoots<T>(Func<T, IEnumerable<T>> childrenSelector, T figure, Action<T> collector)
    {
        FindRoots(childrenSelector, figure, collector, new HashSet<T>());
    }

    /// <summary>
    /// Each figure once: what two figures are both built on was walked again for each of
    /// them, and a drawing of numbers worked out from numbers in long chains of such
    /// diamonds (a ray traced through a prism) took seconds at every press of the mouse.
    /// </summary>
    static void FindRoots<T>(Func<T, IEnumerable<T>> childrenSelector, T figure, Action<T> collector, HashSet<T> visited)
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
            if (!visited.Add(figure))
            {
                continue;
            }
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
        return FigureCompletelyDependsOnFigures(figure, set, new Dictionary<IFigure, bool>());
    }

    // (the answer for each figure once, as FindRoots walks each figure once)
    static bool FigureCompletelyDependsOnFigures(IFigure figure, IEnumerable<IFigure> set, Dictionary<IFigure, bool> known)
    {
        if (known.TryGetValue(figure, out var answer))
        {
            return answer;
        }

        answer = true;
        if (!set.Contains(figure))
        {
            if (figure.Dependencies.IsEmpty())
            {
                answer = false;
            }
            else
            {
                foreach (var dependency in figure.Dependencies)
                {
                    if (!FigureCompletelyDependsOnFigures(dependency, set, known))
                    {
                        answer = false;
                        break;
                    }
                }
            }
        }

        known[figure] = answer;
        return answer;
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
