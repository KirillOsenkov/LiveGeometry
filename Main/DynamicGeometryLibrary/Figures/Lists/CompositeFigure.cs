using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace DynamicGeometry
{
    public partial class CompositeFigure : FigureBase
    {
        public CollectionWithEvents<IFigure> Children { get; private set; }

        public CompositeFigure()
        {
            Children = new CollectionWithEvents<IFigure>();
        }

        public override void OnAddingToDrawing(Drawing drawing)
        {
            base.OnAddingToDrawing(drawing);
            foreach (var item in Children)
            {
                item.OnAddingToDrawing(drawing);
            }
        }

        public override void OnRemovingFromDrawing(Drawing drawing)
        {
            base.OnRemovingFromDrawing(drawing);
            foreach (var item in Children)
            {
                item.OnRemovingFromDrawing(drawing);
            }
        }

        public override void OnAddingToCanvas(Canvas newContainer)
        {
            foreach (var figure in Children)
            {
                figure.OnAddingToCanvas(newContainer);
            }
            base.OnAddingToCanvas(newContainer);
        }

        public override void OnRemovingFromCanvas(Canvas leavingContainer)
        {
            foreach (var figure in Children)
            {
                figure.OnRemovingFromCanvas(leavingContainer);
            }
        }

        public override void Recalculate()
        {
            foreach (var figure in Children)
            {
                if (figure.Exists)
                {
                    figure.Recalculate();
                }
            }
        }

        public override void UpdateVisual()
        {
            if (!this.Visible)
            {
                return;
            }

            foreach (var figure in Children)
            {
                if (figure.Exists)
                {
                    figure.UpdateVisual();
                }
            }
        }

        public override void UpdateExistence()
        {
            base.UpdateExistence();

            foreach (var figure in Children)
            {
                figure.UpdateExistence();
            }
        }

        public override bool Selected
        {
            get
            {
                return base.Selected;
            }
            set
            {
                base.Selected = value;
                foreach (var item in Children)
                {
                    item.Selected = value;
                }
            }
        }

        public override bool Visible
        {
            get
            {
                return mVisible;
            }
            set
            {
                mVisible = value;
                foreach (var item in Children)
                {
                    item.Visible = value;
                }
            }
        }

        public override Drawing Drawing
        {
            get
            {
                return base.Drawing;
            }
            set
            {
                base.Drawing = value;
                foreach (var item in Children)
                {
                    item.Drawing = value;
                }
            }
        }

        public override void ApplyStyle()
        {
            Children.ForEach(f => f.ApplyStyle());
        }

        #region HitTest

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
                figure != null && figure.Visible && figureType.IsAssignableFrom(figure.GetType()));
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

            foreach (var item in Children.Where(f => filter(f)))
            {
                IFigure found = item.HitTest(point);
                if (found != null)
                {
                    if (bestFoundSoFar == null || bestFoundSoFar.ZIndex <= found.ZIndex)
                    {
                        bestFoundSoFar = found;
                    }
                }
            }
            return bestFoundSoFar;
        }

        public ReadOnlyCollection<IFigure> HitTestMany(Point point)
        {
            List<IFigure> result = new List<IFigure>();

            foreach (var item in Children)
            {
                IFigure found = item.HitTest(point);
                if (found != null && found.Visible)
                {
                    result.Add(found);
                }
            }

            if (result.Count > 0)
            {
                result.Sort((f1, f2) => f1.ZIndex.CompareTo(f2.ZIndex));
            }

            return result.AsReadOnly();
        }

        /// <summary>
        /// Finds a figure of a given type at the point. Collections are not searched.
        /// </summary>
        public IFigure HitTestNoCollections(System.Windows.Point point, Type figureType)
        {
            IFigure bestFoundSoFar = null;

            foreach (var item in Children)
            {
                IFigure found = item.HitTest(point);
                if (found != null && figureType.IsAssignableFrom(found.GetType()))
                {
                    if (bestFoundSoFar == null || bestFoundSoFar.ZIndex <= found.ZIndex)
                    {
                        bestFoundSoFar = found;
                    }
                }
            }
            return bestFoundSoFar;
        }

        /// <summary>
        /// Finds any figure at the point
        /// </summary>
        /// <param name="point">Hittest coordinates</param>
        /// <returns>A figure with topmost ZIndex or null if nothing found</returns>
        public override IFigure HitTest(Point point)
        {
            return HitTest(point, f => f.Visible);
        }

        #endregion

        public override string ToString()
        {
            StringBuilder s = new StringBuilder();
            foreach (var item in Children)
            {
                DumpFigure(item, s);
            }
            return s.ToString();
        }

        private void DumpFigure(IFigure item, StringBuilder s)
        {
            s.AppendLine(item.ToString());
            string tab = "   ";
            if (!item.Dependencies.IsEmpty())
            {
                int i = 0;
                foreach (var dependency in item.Dependencies)
                {
                    s.AppendLine(tab + i++.ToString() + ". " + dependency.ToString());
                }
            }
            if (!item.Dependents.IsEmpty())
            {
                foreach (var dependent in item.Dependents)
                {
                    s.AppendLine(tab + dependent.ToString());
                }
            }
        }
    }
}