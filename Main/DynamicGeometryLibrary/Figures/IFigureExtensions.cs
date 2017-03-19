using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace DynamicGeometry
{
    public static class IFigureExtensions
    {
        /// <summary>
        /// Determines if figure directly or indirectly depends 
        /// on <paramref name="possibleDependency"/>
        /// </summary>
        /// <param name="figure">figure to check</param>
        /// <param name="possibleDependency"></param>
        /// <returns></returns>
        public static bool DependsOn(this IFigure figure, IFigure possibleDependency)
        {
            // we consider that a figure depends on itself
            if (figure == possibleDependency)
            {
                return true;
            }

            // quick rejection - if it doesn't depend on anything,
            // it certainly doesn't depend on possibleDependency
            if (figure.Dependencies.IsEmpty())
            {
                return false;
            }
            
            // first do the cheap pre-test without going deep
            if (figure.DirectlyDependsOn(possibleDependency))
            {
                return true;
            }

            // if that failed, go deeper using recursion
            foreach (var directDependency in figure.Dependencies)
            {
                if (directDependency.DependsOn(possibleDependency))
                {
                    return true;
                }
            }

            // depth-first search didn't find anything
            return false;
        }

        public static bool DirectlyDependsOn(this IFigure figure, IFigure possibleDependency)
        {
            // we consider that a figure depends on itself
            if (figure == possibleDependency)
            {
                return true;
            }

            // quick rejection - if it doesn't depend on anything,
            // it certainly doesn't depend on possibleDependency
            if (figure.Dependencies.IsEmpty())
            {
                return false;
            }
            
            return figure.Dependencies.Contains(possibleDependency);
        }

        public static void RecalculateAllDependents(this IFigure figure)
        {
            var dependentsToRecalculate = DependencyAlgorithms
                .FindDescendants(f => f.Dependents, new IFigure[] { figure });
            dependentsToRecalculate.Reverse();

            foreach (var dependent in dependentsToRecalculate)
            {
                dependent.RecalculateAndUpdateVisual();
            }
        }

        public static void RecalculateAndUpdateVisual(this IFigure figure)
        {
            figure.UpdateExistence();
            figure.Recalculate();
            figure.UpdateVisual();
        }

        public static IEnumerable<Point> EnumeratePointsOnLinearFigure(this ILinearFigure figure)
        {
            var domain = figure.GetParameterDomain();
            for (double lambda = domain.Item1; lambda < domain.Item2; lambda += 0.01)
            {
                yield return figure.GetPointFromParameter(lambda);
            }
        }

        public static Point Point(this IFigure figure, int index)
        {
            return (figure.Dependencies.ElementAt(index) as IPoint).Coordinates;
        }

        public static void Move(this IEnumerable<IMovable> figures, Point offset)
        {
            foreach (var figure in figures)
            {
                figure.MoveTo(figure.Coordinates.Plus(offset));
            }
        }

        public static PointPair Line(this IFigure figure, int index)
        {
            return (figure.Dependencies.ElementAt(index) as ILine).Coordinates;
        }

        public static void RegisterWithDependencies(this IFigure figure)
        {
            figure.AddDependencies(figure.Dependencies);
        }

        public static void UnregisterFromDependencies(this IFigure figure)
        {
            figure.RemoveDependencies(figure.Dependencies);
        }

        public static void AddDependencies(this IFigure figure, IEnumerable<IFigure> dependencies)
        {
            if (figure == null || dependencies.IsEmpty())
            {
                return;
            }

            foreach (var dependency in dependencies)
            {
                dependency.Dependents.Add(figure);
            }
        }

        public static void RemoveDependencies(this IFigure figure, IEnumerable<IFigure> dependencies)
        {
            if (figure == null || dependencies.IsEmpty())
            {
                return;
            }
            foreach (var dependency in dependencies)
            {
                if (dependency != null)
                {
                    dependency.Dependents.Remove(figure);
                }
            }
        }

        public static void ReplaceDependency(this IFigure figure, int index, IFigure newDependency)
        {
            List<IFigure> temp = new List<IFigure>(figure.Dependencies);
            if (index < 0 || index >= temp.Count)
            {
                throw new ArgumentOutOfRangeException("index");
            }
            IFigure oldDependency = temp[index];
            oldDependency.Dependents.Remove(figure);
            temp[index] = newDependency;
            newDependency.Dependents.Add(figure);
            figure.Dependencies = temp;

            // Dive down into the children of composite figures updating dependencies.
            var compositeFigure = figure as CompositeFigure;
            if (compositeFigure != null)
            {
                foreach (IFigure child in compositeFigure.Children)
                {
                    temp.Clear();
                    index = child.Dependencies.IndexOf(oldDependency);
                    if (index >= 0)
                    {
                        temp.AddRange(child.Dependencies);
                        temp[index] = newDependency;
                        child.Dependencies = temp;
                    }
                }
            }
        }

        public static void ReplaceDependency(this IFigure figure, IFigure oldDependency, IFigure newDependency)
        {
            int index = figure.Dependencies.IndexOf(oldDependency);
            if (index == -1)
            {
                throw new Exception("Calling ReplaceDependency on a figure where oldDependency is not a dependency");
            }
            ReplaceDependency(figure, index, newDependency);
            figure.UpdateVisual();  // Necessary for Undo of CallMethodAction in Actions.ReplaceDependency(...)
        }

        public static void SubstituteWith(this IFigure figure, IFigure replacement)
        {
            List<IFigure> dependents = new List<IFigure>(figure.Dependents.Where(f => !(f is PointLabel)));
            if (dependents.IsEmpty())
            {
                return;
            }
            foreach (var dependent in dependents)
            {
                dependent.ReplaceDependency(figure, replacement);
            }
            replacement.Dependents.AddRange(figure.Dependents.ToArray());
            figure.Dependents.Clear();
        }

        public static string GenerateNewName(this IFigure figure)
        {
            // Old Scheme - all figures use same index.
            //return figure.GetType().Name + FigureBase.ID++;

            // New Scheme - each class essentially uses its own index.
            string className = figure.GetType().Name;
#if TABULA
            // Some of the classes in Tabula have prefixes that need to be trimmed. - D.H.
            if (className[0] == 'T' && className[1] == 'A' && className[2] == 'B')
            {
                className = className.TrimStart('T', 'A', 'B');
            }
#endif
            for (int i = 1; i < int.MaxValue; i++)
            {
                string number = i.ToString();
                var candidate = className + number;
                if (figure.NameAvailable(candidate))
                {
                    return candidate;
                }
            }

            // Report a naming error.
            if (figure.Drawing != null)
            {
                var message = "Error in generating name for figure of class ";
                figure.Drawing.RaiseError(Application.Current, new Exception(message + figure.GetType().Name));
            }
            return "error_generating_name";
        }

        public static void GenerateNewNameIfNecessary(this IFigure figure, Drawing drawing, List<string> blacklist)
        {
            while (figure.Name == null || drawing
                .Figures
                //.GetAllFiguresRecursive() // Do not look recursively.
                .Where(f => f.Name == figure.Name)
                .Where(f => f != figure)
                .Any())
            {
                figure.Name = figure.GenerateFigureName(blacklist);
            }
        }

        public static bool NameAvailable(this IFigure figure, string name)
        {
            if (figure.Drawing == null)
            {
                return true;
            }
            return !figure.Drawing.Figures.Contains(name);
        }

        public static void CheckConsistency(this IEnumerable<IFigure> list)
        {
            foreach (var figure in list)
            {
                if (figure.Dependencies != null)
                {
                    foreach (var dependency in figure.Dependencies)
                    {
                        if (!list.Contains(dependency))
                        {
                            throw new Exception(
                                "Consistency check failed: dependency {0} of figure {1} expected in the FigureList"
                                .Format(dependency, figure));
                        }
                        if (!dependency.Dependents.Contains(figure))
                        {
                            throw new Exception(
                                "Consistency check failed: figure {0} is not registered in the Dependents list of its dependency {1}"
                                .Format(figure, dependency));
                        }
                    }
                }
                if (figure.Dependents != null)
                {
                    foreach (var dependent in figure.Dependents)
                    {
                        if (!list.Contains(dependent))
                        {
                            throw new Exception(
                                "Consistency check failed: dependent {0} of figure {1} expected in the FigureList"
                                .Format(dependent, figure));
                        }

                        if (!dependent.Dependencies.Contains(figure))
                        {
                            throw new Exception(
                                "Consistency check failed: figure {0} is not registered in the Dependencies list of its dependent {1}"
                                .Format(figure, dependent));
                        }
                    }
                }
            }
        }

        public static void Scale(this IFigure figure, double scaleFactor)
        {
            foreach (IFigure f in figure.Dependencies)
            {
                if (f is PointBase)
                {
                    ((PointBase)f).X = ((PointBase)f).X * scaleFactor;
                    ((PointBase)f).Y = ((PointBase)f).Y * scaleFactor;
                }
                else
                {
                    f.Scale(scaleFactor);
                }
            }
        }
    }
}
