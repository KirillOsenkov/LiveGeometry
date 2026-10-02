using System;
using Avalonia;
using Avalonia.Media;
using System.Xml.Linq;
using System.Globalization;

namespace DynamicGeometry
{
    public class Label : LabelBase, IAngleProvider, ILengthProvider
    {
        public Label()
        {
            ShouldProcessText = true;
        }

        public double Value
        {
            get
            {
                var exact = ExactValue;
                if (exact != null)
                {
                    return exact.Value;
                }

                double result = 0;
                if (!double.TryParse(
                        ProcessedText,
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture,
                        out result)
                    && ProcessedText == UndefinedText)
                {
                    return double.NaN;
                }

                return result;
            }
        }

        /// <summary>
        /// Whether the label says a number - "[AB * 2]", the value of an expression - and
        /// so can stand for a length or an angle where a tool asks for one. Every label
        /// has the interfaces, but a caption has no number to give: taken for 0, a click on
        /// it tied a circle's radius or a rotation's angle to a piece of text.
        /// </summary>
        public bool IsNumber
        {
            get
            {
                return double.TryParse(
                    ProcessedText,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out _);
            }
        }

        /// <summary>
        /// For a tool that takes a figure with a length or an angle: any such figure but a
        /// label that says no number (<see cref="IsNumber"/>)
        /// </summary>
        public static bool GivesNumber(IFigure figure)
        {
            return !(figure is Label label) || label.IsNumber;
        }

        public double Angle
        {
            get {return Value;}
        }

        public double Length
        {
            get { return Value; }
        }

        #region Pinning

        LabelPin pin;

        /// <summary>
        /// A pinned label stays put on the screen while the plane zooms and pans under it: the
        /// named corner of the label sits <see cref="PinOffset"/> pixels inward from the same
        /// corner of the canvas. Changing the pin keeps the label where it is on the screen;
        /// <see cref="Coordinates"/> always tell where it is in the plane right now.
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridName("Pinned to")]
        [PropertyGridCustomValueProvider(typeof(PinValue))]
        public LabelPin Pin
        {
            get
            {
                return pin;
            }
            set
            {
                if (value == pin)
                {
                    return;
                }

                if (value != LabelPin.None && HasCanvas)
                {
                    PinOffset = OffsetFrom(value, ToPhysical(Coordinates), MeasureSize());
                }

                pin = value;

                // on the screen means on top of everything in the plane, like a control
                ZIndex = pin == LabelPin.None ? DefaultZOrder() : (int)ZOrder.Controls;
                if (HasCanvas)
                {
                    UpdateVisual();
                }
            }
        }

        /// <summary>
        /// The pin as the property grid sets it: undo puts the label back where it was - in
        /// the plane if it was not pinned, at its offset from the corner if it was. (Only
        /// the pin went back. A label pinned, the view zoomed, the pin undone: it stayed
        /// where the screen had carried it, somewhere else in the plane than before.)
        /// </summary>
        public class PinValue : PropertyValue, IRestorableValue
        {
            class SavedPin
            {
                public LabelPin Pin;
                public Point Offset;
                public Point Coordinates;
            }

            public object CaptureState()
            {
                var label = (Label)Parent;
                return new SavedPin()
                {
                    Pin = label.Pin,
                    Offset = label.PinOffset,
                    Coordinates = label.Coordinates
                };
            }

            public void RestoreState(object state)
            {
                var label = (Label)Parent;
                var saved = (SavedPin)state;
                label.Pin = saved.Pin;
                if (saved.Pin == LabelPin.None)
                {
                    label.MoveTo(saved.Coordinates);
                }
                else
                {
                    label.PinOffset = saved.Offset;
                    if (label.HasCanvas)
                    {
                        label.UpdateVisual();
                    }
                }
            }
        }

        /// <summary>Pixels from the pinned corner of the canvas to the same corner of the label</summary>
        public Point PinOffset { get; set; }

        // a pin is measured from the canvas: without one there is nothing to measure from
        bool HasCanvas
        {
            get
            {
                return Drawing != null && Drawing.Canvas != null;
            }
        }

        double wrapWidth;

        /// <summary>
        /// The width of the label in pixels, its text wrapping to fit; 0 lets every line run as
        /// long as it is.
        /// </summary>
        public double WrapWidth
        {
            get
            {
                return wrapWidth;
            }
            set
            {
                wrapWidth = value;
                TextBlock.TextWrapping = wrapWidth > 0 ? TextWrapping.Wrap : TextWrapping.NoWrap;
                TextBlock.Width = wrapWidth > 0 ? wrapWidth : double.NaN;
                if (Drawing != null)
                {
                    UpdateVisual();
                }
            }
        }

        bool backdrop;

        /// <summary>
        /// A plate of the paper's color behind the text, a little larger than the text, so
        /// that a caption stays readable over a grid or a figure that zoomed under it. Only a
        /// plain paper has a color to give; on a gradient there is no plate.
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridName("Backdrop")]
        public bool Backdrop
        {
            get
            {
                return backdrop;
            }
            set
            {
                backdrop = value;
                TextBlock.Padding = backdrop ? new Thickness(BackdropPadding) : new Thickness();
                ApplyBackdrop();
                if (HasCanvas)
                {
                    UpdateVisual();
                }
            }
        }

        public const double BackdropPadding = 8;

        void ApplyBackdrop()
        {
            TextBlock.Background = backdrop && Drawing != null ? Drawing.Background as SolidColorBrush : null;
        }

        /// <summary>The top-left corner of a pinned label, in pixels</summary>
        Point PinnedTopLeft(Size size)
        {
            var canvas = Drawing.CoordinateSystem.PhysicalSize;
            double x = pin == LabelPin.TopLeft || pin == LabelPin.BottomLeft
                ? PinOffset.X
                : canvas.X - PinOffset.X - size.Width;
            double y = pin == LabelPin.TopLeft || pin == LabelPin.TopRight
                ? PinOffset.Y
                : canvas.Y - PinOffset.Y - size.Height;
            return new Point(x, y);
        }

        /// <summary>The offset that puts a label of this size at this top-left corner</summary>
        Point OffsetFrom(LabelPin corner, Point topLeft, Size size)
        {
            var canvas = Drawing.CoordinateSystem.PhysicalSize;
            double x = corner == LabelPin.TopLeft || corner == LabelPin.BottomLeft
                ? topLeft.X
                : canvas.X - topLeft.X - size.Width;
            double y = corner == LabelPin.TopLeft || corner == LabelPin.TopRight
                ? topLeft.Y
                : canvas.Y - topLeft.Y - size.Height;
            return new Point(x, y);
        }

        public override void UpdateVisual()
        {
            // the paper may have changed since (a theme switch, the paper of a file, which
            // is read after its figures): a label's plate was the paper it was made on
            if (backdrop)
            {
                ApplyBackdrop();
            }

            if (pin == LabelPin.None)
            {
                base.UpdateVisual();
                return;
            }

            if (!HasCanvas)
            {
                return;
            }

            var topLeft = PinnedTopLeft(MeasureSize());
            Coordinates = ToLogical(topLeft);
            Shape.MoveTo(topLeft);
        }

        /// <summary>
        /// A label with live expressions depends on the figures they name, but dragging it
        /// moves the text, not those figures (as with a measurement).
        /// </summary>
        public override bool AllowMove()
        {
            return !Locked;
        }

        /// <summary>Moves a pinned label on the screen, by pixels: it scrolls along with the view</summary>
        public void ScrollPinned(Avalonia.Vector pixels)
        {
            if (pin == LabelPin.None || !HasCanvas)
            {
                return;
            }

            var size = MeasureSize();
            PinOffset = OffsetFrom(pin, PinnedTopLeft(size) + pixels, size);
            UpdateVisual();
        }

        /// <summary>A pinned label's place is its offset from the corner, in pixels; an unpinned one's is in the plane</summary>
        public override object CapturePlace()
        {
            return pin != LabelPin.None ? new PinnedPlace() { Offset = PinOffset } : base.CapturePlace();
        }

        public override void RestorePlace(object place)
        {
            if (place is PinnedPlace pinned)
            {
                PinOffset = pinned.Offset;
                if (HasCanvas)
                {
                    UpdateVisual();
                }
            }
            else
            {
                base.RestorePlace(place);
            }
        }

        class PinnedPlace
        {
            public Point Offset;
        }

        /// <summary>Dragging a pinned label changes its offset from the corner, not its place in the plane</summary>
        public override void MoveToCore(Point newLocation)
        {
            if (pin != LabelPin.None && HasCanvas)
            {
                PinOffset = OffsetFrom(pin, ToPhysical(newLocation), MeasureSize());
            }

            base.MoveToCore(newLocation);
        }

        #endregion

        public override void ReadXml(XElement element)
        {
            base.ReadXml(element);
            text = Unescape(element.ReadString("Text") ?? "");
            WrapWidth = element.ReadDouble("WrapWidth");
            Backdrop = element.ReadBool("Backdrop", false);
            var pinName = element.ReadString("Pin");
            if (pinName != null && Enum.TryParse(pinName, out LabelPin readPin) && readPin != LabelPin.None)
            {
                PinOffset = new Point(element.ReadDouble("OffsetX"), element.ReadDouble("OffsetY"));
                pin = readPin;
                ZIndex = (int)ZOrder.Controls;
                UpdateVisual();
            }
            else
            {
                var x = element.ReadDouble("X");
                var y = element.ReadDouble("Y");
                MoveTo(new Point(x, y));
            }
        }

        /// <summary>
        /// The text as the file says it: any line break (\r\n, or a bare \n, which a text
        /// box gives back) as the two characters \n, and a backslash as two, so that one
        /// typed before an n ("C:\notes") is not read back as a line break. Characters an
        /// XML file can't hold (a control character pasted from elsewhere) are left out:
        /// with one, Save threw and wrote nothing.
        /// </summary>
        static string Escape(string text)
        {
            var sb = new System.Text.StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '\r')
                {
                    sb.Append(@"\n");
                    if (i + 1 < text.Length && text[i + 1] == '\n')
                    {
                        i++;
                    }
                }
                else if (c == '\n')
                {
                    sb.Append(@"\n");
                }
                else if (c == '\\')
                {
                    sb.Append(@"\\");
                }
                else if (System.Xml.XmlConvert.IsXmlChar(c))
                {
                    sb.Append(c);
                }
                else if (i + 1 < text.Length && System.Xml.XmlConvert.IsXmlSurrogatePair(text[i + 1], c))
                {
                    sb.Append(c).Append(text[i + 1]);
                    i++;
                }
            }

            return sb.ToString();
        }

        static string Unescape(string text)
        {
            var sb = new System.Text.StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '\\' && i + 1 < text.Length && (text[i + 1] == 'n' || text[i + 1] == '\\'))
                {
                    sb.Append(text[i + 1] == 'n' ? Environment.NewLine : @"\");
                    i++;
                }
                else
                {
                    sb.Append(text[i]);
                }
            }

            return sb.ToString();
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            writer.WriteAttributeString("Text", Escape(Text));
            if (pin == LabelPin.None)
            {
                var coordinates = Coordinates;
                writer.WriteAttributeString("X", coordinates.X.ToStringInvariant());
                writer.WriteAttributeString("Y", coordinates.Y.ToStringInvariant());
            }
            else
            {
                writer.WriteAttributeString("Pin", pin.ToString());
                writer.WriteAttributeDouble("OffsetX", PinOffset.X);
                writer.WriteAttributeDouble("OffsetY", PinOffset.Y);
            }

            if (wrapWidth > 0)
            {
                writer.WriteAttributeDouble("WrapWidth", wrapWidth);
            }

            if (backdrop)
            {
                writer.WriteAttributeBool("Backdrop", true);
            }
        }
    }
}
