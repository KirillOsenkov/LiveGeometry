using System.Windows;
using System.Windows.Controls;

namespace DynamicGeometry
{
    public class FigureExplorer : ListBox
    {
        public FigureExplorer()
        {
            this.SelectionMode = SelectionMode.Extended;
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
