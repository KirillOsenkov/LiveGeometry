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
                double result = 0;
                double.TryParse(
                        ProcessedText,
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture,
                        out result);
                return result;
            }
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

        /// <summary>
        /// The size the text takes right now - laid out here and now, because Bounds are stale
        /// right after the text changed or the label was created. The shape is a Border around
        /// the TextBlock, and its own measure stays valid when the text or the width inside it
        /// changed, so it is invalidated first: without that it answered with the old size.
        /// </summary>
        public Size MeasureSize()
        {
            Shape.InvalidateMeasure();
            Shape.Measure(Size.Infinity);
            return Shape.DesiredSize;
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
            if (pin == LabelPin.None)
            {
                base.UpdateVisual();
                return;
            }

            if (!HasCanvas)
            {
                return;
            }

            if (backdrop)
            {
                // the paper may have changed since
                ApplyBackdrop();
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
            text = element.ReadString("Text").Replace(@"\n", Environment.NewLine);
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

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            writer.WriteAttributeString("Text", Text.Replace("\n", "").Replace("\r", @"\n"));
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
