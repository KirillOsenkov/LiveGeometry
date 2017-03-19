using System.Collections;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

namespace DynamicGeometry
{
    public partial class SelectorValueEditor : LabeledValueEditor, IValueEditor
    {
        public Selector Selector { get; set; }

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
            foreach (var item in Items)
            {
                if (item.Equals(value))
                {
                    guard = true;
                    Selector.SelectedItem = item;
                    guard = false;
                }
            }
        }
    }
}
