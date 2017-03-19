using System.Windows;
using System.Windows.Controls;
using System.Xml.Linq;
using System.Windows.Media;

namespace DynamicGeometry
{
    public class ShowHideControl : ControlBase
    {
        public CheckBox Checkbox { get; set; }

        public override void ReadXml(XElement element)
        {
            base.ReadXml(element);
            Checkbox.IsChecked = element.ReadBool("Show", true);
            Checkbox.Content = element.ReadString("Text");
            var x = element.ReadDouble("X");
            var y = element.ReadDouble("Y");
            MoveTo(new Point(x, y));
            UpdateFigureVisibility();
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            var coordinates = Coordinates;
            writer.WriteAttributeBool("Show", Checkbox.IsChecked == true);
            writer.WriteAttributeString("Text", Checkbox.Content.ToString());
            writer.WriteAttributeString("X", coordinates.X.ToStringInvariant());
            writer.WriteAttributeString("Y", coordinates.Y.ToStringInvariant());
        }

        protected override FrameworkElement CreateShape()
        {
            Checkbox = new CheckBox();
            Checkbox.Background = new SolidColorBrush(Color.FromArgb(255, 230, 230, 230));
            Checkbox.Foreground = Brushes.Black;
            Checkbox.Checked += result_Checked;
            Checkbox.Unchecked += result_Unchecked;
            return Checkbox;
        }

        void result_Unchecked(object sender, RoutedEventArgs e)
        {
            Show(false);
        }

        public void UpdateFigureVisibility()
        {
            Show(Checkbox.IsChecked == true);
        }

        private void Show(bool show)
        {
            foreach (var figure in Dependencies)
            {
                figure.Visible = show;
                figure.UpdateVisual();
            }
        }

        void result_Checked(object sender, RoutedEventArgs e)
        {
            Show(true);
        }
    }
}
