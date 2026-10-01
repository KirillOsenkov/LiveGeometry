using System.Collections;
using Avalonia.Controls;
using Avalonia.Layout;

namespace DynamicGeometry
{
    public partial class SelectorValueEditor : LabeledValueEditor, IValueEditor
    {
        public Selector Selector { get; set; }

        /// <summary>Each choice is an undo step of its own</summary>
        protected override bool CoalescesEdits
        {
            get { return false; }
        }

        protected override UIElement CreateEditor()
        {
            Selector = CreateSelector();
            Selector.VerticalAlignment = VerticalAlignment.Center;
            Selector.SelectionChanged += Selector_SelectionChanged;
            return Selector;
        }

        protected virtual Selector CreateSelector()
        {
            var result = new ComboBox();
            result.MaxDropDownHeight = 300;
            return result;
        }

        void Selector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Selector.SelectedItem != null && !guard)
            {
                SetValue(Selector.SelectedItem);
            }
        }

        public IEnumerable Items { get; set; }

        protected override void InitCore()
        {
            FillList();
        }

        protected bool guard = false;
        public virtual void FillList()
        {
            guard = true;
            Selector.Items.Clear();
            if (Items == null)
            {
                return;
            }

            foreach (var item in Items)
            {
                Selector.Items.Add(item);
            }
            if (Selector.Items.Count > 0)
            {
                Selector.SelectedIndex = 0;
            }
            guard = false;
        }

        public override void UpdateEditor()
        {
            var value = GetValue();
            ShowSelected(item => item.Equals(value));
        }

        /// <summary>
        /// Selects the item that stands for the value, or none when no item does: several
        /// figures whose values differ have no value (null), and a value may be none of
        /// the choices (a font that is not in the list). Left on the first item, as the
        /// list starts, the row showed a value nothing has, and that item could not be
        /// picked: it was selected already.
        /// </summary>
        protected void ShowSelected(System.Func<object, bool> standsForValue)
        {
            object selected = null;
            if (Items != null)
            {
                foreach (var item in Items)
                {
                    if (standsForValue(item))
                    {
                        selected = item;
                        break;
                    }
                }
            }

            guard = true;
            if (selected != null)
            {
                Selector.SelectedItem = selected;
            }
            else
            {
                Selector.SelectedIndex = -1;
            }

            guard = false;
        }
    }
}
