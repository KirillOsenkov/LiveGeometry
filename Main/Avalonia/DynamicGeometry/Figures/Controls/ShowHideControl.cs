using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;
using System.Xml.Linq;

namespace DynamicGeometry
{
    public class ShowHideControl : ControlBase
    {
        public CheckBox Checkbox { get; set; }

        // Fluent's check box template colors the caption by state through these resources,
        // over the control's own Foreground
        static readonly string[] captionResourceKeys =
        {
            "CheckBoxForegroundUnchecked",
            "CheckBoxForegroundUncheckedPointerOver",
            "CheckBoxForegroundUncheckedPressed",
            "CheckBoxForegroundChecked",
            "CheckBoxForegroundCheckedPointerOver",
            "CheckBoxForegroundCheckedPressed",
            "CheckBoxForegroundIndeterminate",
            "CheckBoxForegroundIndeterminatePointerOver",
            "CheckBoxForegroundIndeterminatePressed"
        };

        /// <summary>The caption in the text style, in every state (hovered, pressed, checked)</summary>
        public override void ApplyStyle()
        {
            if (Style == null)
            {
                return;
            }

            this.Apply(Checkbox, Style);
            foreach (var key in captionResourceKeys)
            {
                Checkbox.Resources[key] = Checkbox.Foreground;
            }

            base.ApplyStyle();
        }

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

            // no plate of its own on the canvas; the theme paints the hover
            Checkbox.Background = Brushes.Transparent;
            Checkbox.IsCheckedChanged += (s, e) =>
            {
                if (Checkbox.IsChecked == true)
                {
                    result_Checked(s, e);
                }
                else
                {
                    result_Unchecked(s, e);
                }
            };
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
