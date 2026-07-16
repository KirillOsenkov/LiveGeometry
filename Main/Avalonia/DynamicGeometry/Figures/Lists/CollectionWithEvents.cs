using System;
using System.Collections.ObjectModel;

namespace DynamicGeometry
{
    public class CollectionWithEvents<T> : ObservableCollection<T>
    {
        protected virtual void OnItemAdded(T item)
        {
        }

        protected virtual void OnItemRemoved(T item)
        {
        }

        protected override void InsertItem(int index, T item)
        {
            base.InsertItem(index, item);
            OnItemAdded(item);
        }

        protected override void RemoveItem(int index)
        {
            CheckIndex(index);
            OnItemRemoved(this[index]);
            base.RemoveItem(index);
        }

        private void CheckIndex(int index)
        {
            if (index < 0 || index >= base.Count)
            {
                throw new ArgumentOutOfRangeException("index");
            }
        }

        protected override void SetItem(int index, T item)
        {
            CheckIndex(index);
            OnItemRemoved(this[index]);
            base.SetItem(index, item);
            OnItemAdded(item);
        }
    }
}