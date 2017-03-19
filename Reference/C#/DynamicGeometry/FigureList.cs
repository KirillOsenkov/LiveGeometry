using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Windows.Controls;
using System.Collections.ObjectModel;

namespace DynamicGeometry
{
    public interface IFigureList :
        IFigure,
        IList<IFigure>,
        INotifyCollectionChanged,
        INotifyPropertyChanged
    {
    }

    public class FigureGroup : FigureList
    {
        public FigureGroup(Drawing drawing)
            : base()
        {
            Drawing = drawing;
        }

        public Drawing Drawing { get; set; }

        protected override void OnItemAdded(IFigure item)
        {
            base.OnItemAdded(item);
            item.OnPlacingOnContainer(Drawing.Parent);
        }

        protected override void OnItemRemoved(IFigure item)
        {
            base.OnItemRemoved(item);
            item.OnRemovingFromContainer(Drawing.Parent);
        }
    }

    public class FigureList : CollectionWithEvents<IFigure>, IFigureList
    {
        public IFigureList Dependencies
        {
            get { throw new System.NotImplementedException(); }
        }

        public IFigureList Dependents
        {
            get { throw new System.NotImplementedException(); }
        }

        public void Recalculate()
        {
            foreach (var item in this)
            {
                foreach (var dependent in item.Dependents)
                {
                    dependent.Recalculate();
                }
            }
        }

        public IFigure HitTest(System.Windows.Point point)
        {
            foreach (var item in this)
            {
                IFigure found = item.HitTest(point);
                if (found != null)
                {
                    return found;
                }
            }
            return null;
        }

        public void OnPlacingOnContainer(Canvas newContainer)
        {
            throw new System.NotImplementedException();
        }

        public void OnRemovingFromContainer(Canvas leavingContainer)
        {
            throw new System.NotImplementedException();
        }
    }

    public class CollectionWithEvents<T> : ObservableCollection<T>
    {
        protected override void InsertItem(int index, T item)
        {
            base.InsertItem(index, item);
            OnItemAdded(item);
        }

        protected virtual void OnItemAdded(T item)
        {
            
        }

        protected override void RemoveItem(int index)
        {
            OnItemRemoved(this[index]);
            base.RemoveItem(index);
        }

        protected virtual void OnItemRemoved(T item)
        {
            
        }

        protected override void SetItem(int index, T item)
        {
            OnItemRemoved(this[index]);
            base.SetItem(index, item);
            OnItemAdded(item);
        }
    }
}