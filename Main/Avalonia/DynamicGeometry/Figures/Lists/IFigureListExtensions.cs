using System.Collections.Generic;
using System.Linq;
using Avalonia;

namespace DynamicGeometry
{
    public static class IFigureListExtensions
    {
        public static IEnumerable<IFigure> GetAllFiguresRecursive(
            this IFigure rootFigure)
        {
            CompositeFigure list = rootFigure as CompositeFigure;
            if (list != null)
            {
                foreach (var item in list.Children)
                {
                    foreach (var recursive in item.GetAllFiguresRecursive())
                    {
                        yield return recursive;
                    }
                }
            }
            else
            {
                yield return rootFigure;
            }
        }

        public static IEnumerable<IFigure> AllDependents(this IFigure figure)
        {
            var hashset = new HashSet<IFigure>();
            CollectAllDependents(figure, hashset);
            return hashset;
        }

        public static void CollectAllDependents(this IFigure figure, HashSet<IFigure> collector)
        {
            if (!collector.Add(figure))
            {
                return;
            }

            foreach (var dependent in figure.Dependents)
            {
                CollectAllDependents(dependent, collector);
            }
        }

        public static IEnumerable<IFigure> GetAllFiguresRecursive(
            this IEnumerable<IFigure> figureList)
        {
            List<IFigure> result = new List<IFigure>();
            foreach (var figure in figureList)
            {
                result.AddRange(figure.GetAllFiguresRecursive());
            }
            return result;
        }

        /// <summary>
        /// A visible segment, ray, line or vector on the two points: what already draws the
        /// side of a polygon between them. Not any line that depends on both - the
        /// perpendicular bisector of the two, the bisector of an angle at one of them, a
        /// hidden helper: taken for the side, it left the side undrawn.
        /// </summary>
        public static ILine FindLine(this IEnumerable<IFigure> figures, IPoint p1, IPoint p2)
        {
            foreach (var figure in figures)
            {
                var type = figure.GetType();
                bool runsThroughItsPoints = type == typeof(Segment)
                    || type == typeof(Ray)
                    || type == typeof(LineTwoPoints)
                    || type == typeof(Vector);
                if (runsThroughItsPoints
                    && figure.Visible
                    && figure.Dependencies.Count == 2
                    && figure.Dependencies.Contains(p1)
                    && figure.Dependencies.Contains(p2))
                {
                    return figure as ILine;
                }
            }

            return null;
        }

        public static IPoint FindPoint(this IEnumerable<IFigure> figures, Point coordinates, double epsilon)
        {
            foreach (var point in figures.OfType<IPoint>())
            {
                if (point.Coordinates.Distance(coordinates) < epsilon)
                {
                    return point;
                }
            }
            return null;
        }

        public static IFigure FindFigureWithTheseDependencies<TFigure>(
            this IEnumerable<IFigure> figures,
            params IFigure[] dependencies) where TFigure : IFigure
        {
            foreach (var figure in figures)
            {
                if (figure is TFigure && figure.Dependencies.Match(dependencies))
                {
                    return figure;
                }
            }
            return null;
        }

        public static IFigure FindFigureWithTheseDependenciesInSameOrder<TFigure>(
            this IEnumerable<IFigure> figures,
            params IFigure[] dependencies) where TFigure : IFigure
        {
            foreach (var figure in figures)
            {
                if (figure is TFigure && figure.Dependencies.MatchInSameOrder(dependencies))
                {
                    return figure;
                }
            }
            return null;
        }

        public static bool Match(
            this IEnumerable<IFigure> figures,
            params IFigure[] givenFigures)
        {
            if (figures.Count() != givenFigures.Length)
            {
                return false;
            }
            foreach (var given in givenFigures)
            {
                if (!figures.Contains(given))
                {
                    return false;
                }
            }
            return true;
        }

        public static bool MatchInSameOrder(
            this IEnumerable<IFigure> figures,
            params IFigure[] givenFigures)
        {
            if (figures.Count() != givenFigures.Length)
            {
                return false;
            }
            int i = 0;
            foreach (var given in givenFigures)
            {
                if (given != givenFigures[i++])
                {
                    return false;
                }
            }
            return true;
        }

        public static bool Exists(this IList<IFigure> figures)
        {
            for (int i = 0; i < figures.Count; i++)
            {
                if (!figures[i].Exists)
                {
                    return false;
                }
            }

            return true;
        }

        public static PointPair GetLogicalBounds(this IEnumerable<IFigure> figures)
        {
            IEnumerable<Point> points = figures.ToPoints();
            double minX = points.Min(p => p.X);
            double minY = points.Min(p => p.Y);
            double maxX = points.Max(p => p.X);
            double maxY = points.Max(p => p.Y);
            return new PointPair(minX, minY, maxX, maxY);
        }
    }
}
