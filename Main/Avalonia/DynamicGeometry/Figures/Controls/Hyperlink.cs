using System;
using System.Net.Http;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;

namespace DynamicGeometry
{
#if !SILVERLIGHT

    public class HyperlinkButton : TextBlock
    {
        public HyperlinkButton()
        {
            TextDecorations = Avalonia.Media.TextDecorations.Underline;
            Cursor = new Avalonia.Input.Cursor(Avalonia.Input.StandardCursorType.Hand);
            PointerReleased += (sender, e) => Click?.Invoke(this, new RoutedEventArgs());
        }

        public event RoutedEventHandler Click;

        public string Content
        {
            get
            {
                return Text;
            }
            set
            {
                Text = value;
            }
        }
    }

#endif

    public class Hyperlink : CoordinatesShapeBase<HyperlinkButton>, IMovable
    {
        public Hyperlink()
        {
            Shape = CreateShape();
            Layer = ZOrder.Labels;
            Shape.Click += Shape_Click;
            Enabled = true;
        }

        static readonly HttpClient internet = new HttpClient();

        protected override string Kind
        {
            get
            {
                return "Link";
            }
        }

        protected override bool NamedByConstruction
        {
            get
            {
                return true;
            }
        }

        /// <summary>The text, in quotes</summary>
        public override string Construction
        {
            get
            {
                return ConstructionText.Quote(Shape?.Content?.ToString());
            }
        }

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
            }
        }

        /// <summary>
        /// The drawing at the link is fetched when the link is clicked. (It was fetched as
        /// soon as the link was read from a file, and a link that led nowhere put the whole
        /// error, stack trace and all, into its own text - which was then saved with it.)
        /// </summary>
        async void Shape_Click(object sender, RoutedEventArgs e)
        {
            if (!Enabled || string.IsNullOrEmpty(mUrl) || Drawing == null)
            {
                return;
            }

            var drawing = Drawing;
            string text;
            try
            {
                // the status is asked rather than thrown: an exception, even caught, is an
                // error report on screen
                using var response = await internet.GetAsync(mUrl);
                if (!response.IsSuccessStatusCode)
                {
                    drawing.RaiseStatusNotification("The drawing at " + mUrl + " could not be opened (" + (int)response.StatusCode + ").");
                    return;
                }

                text = Utilities.StripByteOrderMark(await response.Content.ReadAsStringAsync());
            }
            catch (Exception ex)
            {
                drawing.RaiseStatusNotification("The drawing at " + mUrl + " could not be opened: " + ex.Message);
                return;
            }

            drawing.RaiseDocumentOpenRequested(new Drawing.DocumentOpenRequestedEventArgs()
            {
                DocumentXml = text,
                InWhichWindow = Drawing.DocumentOpenRequestedEventArgs.InWhichWindowChoice.DontCare
            });
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
            if (!IsShown)
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