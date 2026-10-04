using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Avalonia;

namespace DynamicGeometry
{
    /// <summary>
    /// Any set of IFigures
    /// </summary>
    public abstract partial class FigureList : CollectionWithEvents<IFigure>
    {
        public FigureList(Drawing drawing)
        {
            Drawing = drawing;
        }

        public IFigure this[string index]
        {
            get
            {
                return FindByName(this, index);
            }
        }

        /// <summary>
        /// The first figure of that name, in the order of the list, looking inside composite
        /// figures and not at them. Without making a list of every figure: an expression
        /// looks up each name it says, and a drawing of hundreds of points by coordinates
        /// spent a good part of its loading time in those lists.
        /// </summary>
        static IFigure FindByName(IList<IFigure> figures, string name)
        {
            for (int i = 0; i < figures.Count; i++)
            {
                var figure = figures[i];
                var found = figure is CompositeFigure composite
                    ? FindByName(composite.Children, name)
                    : figure.Name == name ? figure : null;
                if (found != null)
                {
                    return found;
                }
            }

            return null;
        }

        public bool Contains(string name)
        {
            // This should not be recursive. - D.H.
            //return this.GetAllFiguresRecursive().Where(f => f.Name == name).Any();
            return this.Where(f => f.Name == name).Any();
        }

        private Drawing mDrawing = null;
        public Drawing Drawing
        {
            get
            {
                return mDrawing;
            }
            set
            {
                if (mDrawing == value)
                {
                    return;
                }
                mDrawing = value;
                if (mDrawing != null)
                {
                    foreach (var item in this)
                    {
                        item.Drawing = mDrawing;
                    }
                }
            }
        }

        public void Add(params IFigure[] figures)
        {
            foreach (var figure in figures)
            {
                if (figure.Drawing == null)
                {
                    figure.Drawing = this.Drawing;
                }
                base.Add(figure);
            }
        }

        public void Remove(IEnumerable<IFigure> figures)
        {
            foreach (var figure in figures.ToArray())
            {
                base.Remove(figure);
            }
        }

        public void ClearSelection()
        {
            foreach (var figure in this)
            {
                // (a side of a polygon selected by itself too: off with the polygon)
                if (FigureParts.HasSelection(figure))
                {
                    figure.Selected = false;
                }
            }
            Drawing.RaiseSelectionChanged();
        }

        public void EnableAll()
        {
            foreach (var figure in this)
            {
                figure.Enabled = true;
            }
        }

        public virtual void Recalculate()
        {
            foreach (var figure in this)
            {
                figure.Recalculate();
            }
        }

        public void UpdateVisual()
        {
            foreach (var figure in this)
            {
                if (figure.Exists)
                {
                    figure.UpdateVisual();
                }
            }
        }

        #region HitTest

        /// <summary>
        /// Finds any figure at the point
        /// </summary>
        /// <param name="point">Hittest coordinates</param>
        /// <returns>A figure with topmost ZIndex or null if nothing found</returns>
        public virtual IFigure HitTest(Point point)
        {
            return HitTest(point, figure => figure.Visible && figure.IsHitTestVisible);
        }

        /// <summary>
        /// Finds a figure of a given type at the point
        /// </summary>
        /// <param name="point">Coordinates</param>
        /// <param name="figureType">A type (usually typeof(IPoint), typeof(ILine) or typeof(ICircle)
        /// but could be anything)</param>
        /// <returns>A figure or null if nothing was found</returns>
        public IFigure HitTest(Point point, Type figureType)
        {
            return HitTest(point, figure =>
                figure != null 
                && figure.Visible 
                && figure.IsHitTestVisible 
                && figureType.IsAssignableFrom(figure.GetType()));
        }

        /// <summary>
        /// Finds a figure of a given type at the point
        /// </summary>
        /// <typeparam name="T">A type (usually IPoint, ILine or ICircle
        /// but could be anything)</typeparam>
        /// <param name="point">Coordinates</param>
        /// <returns>A figure or null if nothing was found</returns>
        public T HitTest<T>(Point point)
        {
            return (T)HitTest(point, typeof(T));
        }

        /// <summary>
        /// Finds a figure at a point
        /// </summary>
        /// <param name="point">Coordinates of a point where we want to find objects</param>
        /// <param name="filter">Determines whether a figure should be included in hit-testing</param>
        /// <returns>A figure with topmost ZIndex that is under the point
        /// and for which the filter is true. Returns null if nothing is found.</returns>
        public virtual IFigure HitTest(Point point, Predicate<IFigure> filter)
        {
            IFigure bestFoundSoFar = null;

            foreach (var item in this)
            {
                // A figure that doesn't exist right now (an intersection whose lines no
                // longer cross, a circle built on it) is nowhere to be clicked. Its own hit
                // test still answers, from where it was last: a tool took the invisible
                // figure, and what was built on it never appeared.
                if (!item.Exists)
                {
                    continue;
                }

                IFigure found = item.HitTest(point);
                if (found != null && found.Exists && filter(found))
                {
                    if (bestFoundSoFar == null || bestFoundSoFar.ZIndex <= found.ZIndex)
                    {
                        // of two nearby points, pick the one which is closer to the hit point
                        var oldPoint = bestFoundSoFar as IPoint;
                        var newPoint = found as IPoint;
                        if (oldPoint != null 
                            && newPoint != null 
                            && oldPoint.Coordinates.Distance(point) < newPoint.Coordinates.Distance(point))
                        {
                            continue;
                        }

                        bestFoundSoFar = found;
                    }
                }
            }

            return bestFoundSoFar;
        }

        public ReadOnlyCollection<IFigure> HitTestMany(Point point)
        {
            List<IFigure> result = new List<IFigure>();

            // Changed to use reverse so that more recently added figures appear earlier in the result. - D.H. 6/29/2011
            var reverse = this.Reverse();
            foreach (var item in reverse)
            {
                if (!item.Exists)
                {
                    continue;
                }

                IFigure found = item.HitTest(point);
                if (found != null && found.Exists && found.Visible && found.IsHitTestVisible)
                {
                    result.Add(found);
                }
            }

            if (result.Count > 0)
            {
                // Use a stable sorting method.  Higher zIndexes should occur earlier in results.  - D.H. 6/29/2011
                FigureList.StableSort(result, (f1, f2) => f2.ZIndex.CompareTo(f1.ZIndex));

                // .NET collections sort is unstable (does not preserve order if comparison value is the same).
                //result.Sort((f1, f2) => f1.ZIndex.CompareTo(f2.ZIndex));
            }

            return result.AsReadOnly();
        }
        
        /// <summary>
        /// Sort the given list using the comparison while preserving order if comparison value is same.
        /// </summary>
        private static void StableSort<T>(IList<T> list, Comparison<T> comparison)
        {
            if (list == null)
                throw new ArgumentNullException("list");
            if (comparison == null)
                throw new ArgumentNullException("comparison");

            int count = list.Count;
            for (int j = 1; j < count; j++)
            {
                T key = list[j];

                int i = j - 1;
                for (; i >= 0 && comparison(list[i], key) > 0; i--)
                {
                    list[i + 1] = list[i];
                }
                list[i + 1] = key;
            }
        }

        #endregion
    }
}