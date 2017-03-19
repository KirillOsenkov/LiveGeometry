using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace DynamicGeometry
{
#if !SILVERLIGHT

    public class HyperlinkButton : TextBlock
    {
        public HyperlinkButton()
        {
            text = new System.Windows.Documents.Run();
            hyperlink = new System.Windows.Documents.Hyperlink(text);
            this.Inlines.Add(hyperlink);
            hyperlink.Click += new RoutedEventHandler(hyperlink_Click);
        }

        void hyperlink_Click(object sender, RoutedEventArgs e)
        {
            if (Click != null)
            {
                Click(sender, e);
            }
        }

        System.Windows.Documents.Run text;
        System.Windows.Documents.Hyperlink hyperlink;

        public event RoutedEventHandler Click;

        public string Content
        {
            get
            {
                return text.Text;
            }
            set
            {
                text.Text = value;
            }
        }
    }

#endif

    public class Hyperlink : CoordinatesShapeBase<HyperlinkButton>, IMovable
    {
        public Hyperlink()
        {
            Shape = CreateShape();
            ZIndex = (int)ZOrder.Labels;
            Shape.Click += Shape_Click;
            internet.DownloadStringCompleted += internet_DownloadStringCompleted;
            Enabled = true;
        }

        WebClient internet = new WebClient();

        private string mUrl = null;
        [PropertyGridVisible]
        public string Url
        {
            get
            {
                return mUrl;
            }
            set
            {
                mUrl = value;
                Shape.IsEnabled = false;
                internet.DownloadStringAsync(new System.Uri(value));
            }
        }

        void Shape_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(fileText))
            {
                Drawing.RaiseDocumentOpenRequested(new Drawing.DocumentOpenRequestedEventArgs()
                {
                    DocumentXml = fileText,
                    InWhichWindow = Drawing.DocumentOpenRequestedEventArgs.InWhichWindowChoice.DontCare
                });
            }
        }

        string fileText = null;

        void internet_DownloadStringCompleted(object sender, DownloadStringCompletedEventArgs e)
        {
            if (e.Cancelled || e.Error != null)
            {
                if (e.Cancelled)
                {
                    Shape.Content = "Cancelled";
                }
                if (e.Error != null)
                {
                    Shape.Content = "Error: " + e.Error.ToString();
                }
                return;
            }
            fileText = Utilities.StripByteOrderMark(e.Result);
            if (Enabled)
            {
                Shape.IsEnabled = true;
            }
        }

        public override void Recalculate()
        {
            if (Settings.ScaleTextWithDrawing)
            {
                var s = Drawing.CoordinateSystem.Scale;
                ScaleTransform scale = new ScaleTransform();
                scale.ScaleX = s;
                scale.ScaleY = s;
                Shape.RenderTransform = scale;
            }
            base.Recalculate();
        }

        public override void UpdateVisual()
        {
            if (!Visible || !Exists)
            {
                return;
            }

            shape.MoveTo(ToPhysical(Coordinates));
        }

        protected override HyperlinkButton CreateShape()
        {
            return new HyperlinkButton()
            {
                FontSize = 20,
                Foreground = new SolidColorBrush(Colors.Blue)
            };
        }

        [PropertyGridVisible]
        public string Text
        {
            get
            {
                return Shape.Content.ToString();
            }
            set
            {
                Shape.Content = value;
            }
        }

        [PropertyGridVisible]
        public override bool Enabled
        {
            get
            {
                return base.Enabled;
            }
            set
            {
                if (value)
                {
                    Shape.CaptureMouse();
                }
                else
                {
                    Shape.ReleaseMouseCapture();
                }
                base.Enabled = value;
                Shape.IsEnabled = value;
            }
        }

        public override IFigure HitTest(Point point)
        {
            double left = Canvas.GetLeft(Shape);
            double top = Canvas.GetTop(Shape);
            point = ToPhysical(point);

            if (left <= point.X
                && left + Shape.ActualWidth >= point.X
                && top <= point.Y
                && top + Shape.ActualHeight >= point.Y)
            {
                return this;
            }
            return null;
        }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            base.ReadXml(element);
            Url = element.ReadString("Url");
            Text = element.ReadString("Text");
            var x = element.ReadDouble("X");
            var y = element.ReadDouble("Y");
            Enabled = element.ReadBool("Enabled", true);
            this.MoveTo(x, y);
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            writer.WriteAttributeString("Url", Url);
            writer.WriteAttributeString("Text", Text);
            writer.WriteAttributeDouble("X", Coordinates.X);
            writer.WriteAttributeDouble("Y", Coordinates.Y);
            writer.WriteAttributeBool("Enabled", Enabled);
        }
    }
}