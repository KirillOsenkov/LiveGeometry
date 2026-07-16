using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Interactivity;
using Avalonia.Controls;

namespace DynamicGeometry
{
    public class FigureExplorer : ListBox
    {
        // Avalonia matches theme templates by concrete type; keep using the ListBox template.
        protected override System.Type StyleKeyOverride => typeof(ListBox);

        public FigureExplorer()
        {
            this.SelectionMode = SelectionMode.Multiple;
        }

        public bool Visible
        {
            get
            {
                return this.Visibility == Visibility.Visible;
            }
            set
            {
                this.Visibility = value ? Visibility.Visible : Visibility.Collapsed;
                Settings.Instance.ShowFigureExplorer = value;
            }
        }
    }
}
