using System.Linq;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    public class ShowHideControl : ControlBase
    {
        public CheckBox Checkbox { get; set; }

        // Fluent's check box template colors the caption by state through these resources,
        // over the control's own Foreground; and the outline of the empty box, which is the
        // theme's and was dark on a drawing's own dark paper (a night sky) under the light theme
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
            "CheckBoxForegroundIndeterminatePressed",
            "CheckBoxCheckBackgroundStrokeUnchecked",
            "CheckBoxCheckBackgroundStrokeUncheckedPointerOver",
            "CheckBoxCheckBackgroundStrokeUncheckedPressed"
        };

        /// <summary>The caption and the empty box in the text style's color, in every state (hovered, pressed, checked)</summary>
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

            // Only the box: each figure says in the file whether it is hidden. Hiding them
            // all again here undid what the user had done since the box was last clicked -
            // one of its figures shown by hand came back hidden when the file was opened.
            SetBox(element.ReadBool("Show", true));
            Checkbox.Content = element.ReadString("Text");
            var x = element.ReadDouble("X");
            var y = element.ReadDouble("Y");
            MoveTo(new Point(x, y));
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
                if (!settingBox)
                {
                    Toggle(Checkbox.IsChecked == true);
                }
            };
            return Checkbox;
        }

        // the box is being ticked by the program (undo, redo): not a click to record
        bool settingBox;

        /// <summary>
        /// A click on the box. The box and what it shows are saved with the drawing, so this
        /// is an undo step: undo puts the box back and gives each figure the visibility it
        /// had, which need not have been what the box said.
        /// </summary>
        void Toggle(bool show)
        {
            // from a file, or in the middle of a construction, whose undo step is its figure's alone
            if (Drawing == null || !Drawing.Figures.Contains(this) || Drawing.IsRecordingTransaction)
            {
                Show(show);
                return;
            }

            var figures = Dependencies.ToArray();
            var before = figures.Select(figure => figure.Visible).ToArray();
            Drawing.ActionManager.RecordAction(new CallMethodAction(
                () =>
                {
                    SetBox(show);
                    Show(show);
                },
                () =>
                {
                    SetBox(!show);
                    for (int i = 0; i < figures.Length; i++)
                    {
                        figures[i].Visible = before[i];
                        figures[i].UpdateVisual();
                    }
                }));
        }

        void SetBox(bool isChecked)
        {
            settingBox = true;
            try
            {
                Checkbox.IsChecked = isChecked;
            }
            finally
            {
                settingBox = false;
            }
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

        /// <summary>
        /// The box stays whether or not its figures exist right now: they are what it shows and
        /// hides, not what it is built on. Taken for dependencies, a hint that was nowhere at the
        /// moment (a firework that hadn't burst yet) took the box along.
        /// </summary>
        public override void UpdateExistence()
        {
            Exists = true;
        }
    }
}
