using System.Collections.Generic;
using System.Linq;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    public class RemoveItemAction<T> : AbstractAction
    {
        public RemoveItemAction(IList<T> list, T item)
        {
            this.item = item;
            this.list = list;
        }

        private T item;
        private IList<T> list;
        private IList<int> indices;

        protected override void ExecuteCore()
        {
            indices = new List<int>();
            for (int i = 0; i < list.Count; i++)
            {
                if (list[i].Equals(item))
                {
                    indices.Add(i);
                }
            }
            foreach (var index in indices.Reverse())
            {
                list.RemoveAt(index);
            }
        }

        protected override void UnExecuteCore()
        {
            foreach (var index in indices)
            {
                list.Insert(index, item);
            }
        }
    }
}