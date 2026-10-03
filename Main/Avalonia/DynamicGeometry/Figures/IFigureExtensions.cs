using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;

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

        /// <summary>
        /// Whether something built on the figure takes it for its length: a distance
        /// measurement of it, a circle with it for a radius, a translation or a dilation by
        /// it. A figure that becomes a kind without a length (a segment converted to a line
        /// or a ray) would leave them with nothing to measure - the measurement threw on
        /// every redraw - so the conversion is not offered then.
        /// </summary>
        public static bool IsUsedForLength(this IFigure figure)
        {
            return figure.Dependents.Any(dependent =>
                dependent is DistanceMeasurement
                || dependent is CircleByRadius
                || dependent is DilatedPoint
                || dependent is TranslatedPoint translated && translated.DistanceSource == figure);
        }

        /// <summary>The same for its area: a sector or a circular segment converted to a bare arc has none</summary>
        public static bool IsUsedForArea(this IFigure figure)
        {
            return figure.Dependents.Any(dependent => dependent is AreaMeasurement);
        }

        /// <summary>
        /// Whether a tool or a tied value may take a length from the figure: anything that
        /// has one, except a label that says no number (<see cref="Label.GivesNumber"/>)
        /// and the mark of an angle - an arc in the code only, a sign of a size in pixels,
        /// whose "length" changes with every zoom (a click just beside a vertex took it
        /// for a radius or a distance).
        /// </summary>
        public static bool GivesLength(this IFigure figure)
        {
            return figure is ILengthProvider && !(figure is AngleArc) && Label.GivesNumber(figure);
        }

        /// <summary>The same for an angle: anything that has one, except a label that says no number</summary>
        public static bool GivesAngle(this IFigure figure)
        {
            return figure is IAngleProvider && Label.GivesNumber(figure);
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

            // a file being read is checked once all of it is in (DrawingDeserializer.ReadDrawing)
            if (!figure.Drawing.IsReading)
            {
                figure.Drawing.Figures.CheckConsistencyInDebug();
            }
        }

        public static void RecalculateAndUpdateVisual(this IFigure figure)
        {
            if (figure.Drawing == null)
            {
                return;
            }

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

                        // a part that is registered with what it is built on (the sides of
                        // a regular polygon) is registered with the new one from now on
                        if (oldDependency.Dependents.Remove(child))
                        {
                            newDependency.Dependents.Add(child);
                        }
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

            // a stored position (a rotated point) follows the new dependency, and what is built
            // on the figure follows that - on undo of Actions.ReplaceDependency too
            figure.RecalculateAndUpdateVisual();
            figure.RecalculateAllDependents();
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
                // a part of a composite (a side of a regular polygon built on the figure)
                // has gone over with its composite already
                if (dependent.Dependencies.Contains(figure))
                {
                    dependent.ReplaceDependency(figure, replacement);
                }
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
            // its own name is available to a figure that looks for a new one (segment AB,
            // renamed when a point is, may well stay AB)
            return !figure.Drawing.Figures.Any(f => f != figure && f.Name == name);
        }

        public static bool ContainsRecursively(this IEnumerable<IFigure> list, IFigure figure)
        {
            foreach (var item in list)
            {
                if (item == figure)
                {
                    return true;
                }

                if (item is CompositeFigure composite && composite.Children.ContainsRecursively(figure))
                {
                    return true;
                }
            }

            return false;
        }

        public static void CheckConsistency(this IEnumerable<IFigure> list)
        {
            // the figures of the list and of their parts, collected once: a search of the list
            // for every dependency and dependent of every figure took time quadratic in the
            // size of the drawing, and the check runs whenever what is built on a figure is
            // recalculated (RecalculateAllDependents)
            var figures = new HashSet<IFigure>(ReferenceEqualityComparer.Instance);
            AddRecursively(figures, list);
            foreach (var figure in list)
            {
                if (figure.Dependencies != null)
                {
                    foreach (var dependency in figure.Dependencies)
                    {
                        if (!figures.Contains(dependency))
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
                        if (!figures.Contains(dependent))
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

        /// <summary>
        /// The app checks itself (after a click of a tool, a recalculation of what is built on a
        /// figure, a file read) in a Debug build only. In Release (the browser) a failed check
        /// could only throw after the fact, or refuse a file that opens otherwise; the tests
        /// and the harness call <see cref="CheckConsistency"/> themselves.
        /// </summary>
        [System.Diagnostics.Conditional("DEBUG")]
        public static void CheckConsistencyInDebug(this IEnumerable<IFigure> list)
        {
            list.CheckConsistency();
        }

        /// <summary>What <see cref="ContainsRecursively"/> finds: the figures and the parts of composites, down to the last</summary>
        static void AddRecursively(HashSet<IFigure> figures, IEnumerable<IFigure> list)
        {
            foreach (var item in list)
            {
                figures.Add(item);
                if (item is CompositeFigure composite)
                {
                    AddRecursively(figures, composite.Children);
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
