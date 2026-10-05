namespace DynamicGeometry
{
    public partial class RootFigureList : FigureList
    {
        public RootFigureList(Drawing drawing)
            : base(drawing)
        {
        }

        // where a figure that was taken out by Retire was in the list
        readonly System.Runtime.CompilerServices.ConditionalWeakTable<IFigure, object> retiredPlaces
            = new System.Runtime.CompilerServices.ConditionalWeakTable<IFigure, object>();

        /// <summary>
        /// Takes out a figure that a property setter adds and removes itself, outside the
        /// undo history (a point's label, a line's name, a length measurement), and remembers
        /// its place: <see cref="Return"/> puts it back there, so that undo of the set
        /// restores the order of the list too (the order a file is written in) and not only
        /// the figure.
        /// </summary>
        public void Retire(IFigure figure)
        {
            int index = IndexOf(figure);
            if (index < 0)
            {
                return;
            }

            retiredPlaces.AddOrUpdate(figure, index);
            RemoveAt(index);
        }

        /// <summary>
        /// Adds a figure of a setter: where it was when it was retired, if it was, and that
        /// is still after <paramref name="owner"/> (the figure it belongs to and depends on);
        /// else at the end.
        /// </summary>
        public void Return(IFigure figure, IFigure owner)
        {
            if (retiredPlaces.TryGetValue(figure, out var place))
            {
                retiredPlaces.Remove(figure);
                int index = (int)place;
                if (index > IndexOf(owner) && index <= Count)
                {
                    Insert(index, figure);
                    return;
                }
            }

            Add(figure);
        }

        /// <summary>
        /// Like <see cref="Return"/>, for a figure that <paramref name="dependent"/> is built
        /// on (the Number of a typed value): where it was if that is still before the
        /// dependent, else just before it.
        /// </summary>
        public void ReturnBefore(IFigure figure, IFigure dependent)
        {
            int limit = IndexOf(dependent);
            if (limit < 0)
            {
                limit = Count;
            }

            int index = limit;
            if (retiredPlaces.TryGetValue(figure, out var place))
            {
                retiredPlaces.Remove(figure);
                if ((int)place <= limit)
                {
                    index = (int)place;
                }
            }

            Insert(index, figure);
        }

        /// <summary>
        /// The figure of the list that the figure is, or is a part of (a side of a regular
        /// polygon, the knob of a slider); null if it is neither
        /// </summary>
        public IFigure FindTopLevel(IFigure figure)
        {
            foreach (var item in this)
            {
                if (item == figure || item is CompositeFigure composite && composite.Children.ContainsRecursively(figure))
                {
                    return item;
                }
            }

            return null;
        }

        protected override System.Collections.Generic.IEnumerable<IFigure> HitTestCandidates
        {
            get
            {
                return System.Linq.Enumerable.Concat(this, Drawing.UnlistedAxisLines());
            }
        }

        protected override void OnItemAdded(IFigure item)
        {
            item.RegisterWithDependencies();
            item.OnAddingToDrawing(Drawing);
            if (Drawing.Canvas != null)
            {
                item.OnAddingToCanvas(Drawing.Canvas);
                item.RecalculateAndUpdateVisual();
            }

            // segment AB that comes back before line AB (undo of its deletion) is AB again
            FigureBase.SettleDefaultNames(Drawing, item);
        }

        protected override void RemoveItem(int index)
        {
            var item = this[index];
            base.RemoveItem(index);

            // line AB2 is AB once segment AB is gone
            FigureBase.SettleDefaultNamesAfter(Drawing, item.Name);
        }

        protected override void MoveItem(int oldIndex, int newIndex)
        {
            base.MoveItem(oldIndex, newIndex);
            FigureBase.SettleDefaultNames(Drawing, this[newIndex]);
        }

        protected override void OnItemRemoved(IFigure item)
        {
            item.OnRemovingFromDrawing(Drawing);
            if (Drawing.Canvas != null)
            {
                item.OnRemovingFromCanvas(Drawing.Canvas);
            }
            item.UnregisterFromDependencies();
        }
    }
}
