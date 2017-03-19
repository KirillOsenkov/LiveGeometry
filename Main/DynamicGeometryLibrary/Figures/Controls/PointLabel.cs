using System.Linq;
using System.Windows;

namespace DynamicGeometry
{
    public class PointLabel : Measurement
    {
        public PointLabel()
        {
        }

        PointBase GetPoint()
        {
            return (PointBase)Dependencies.ElementAt(0);
        }

        public override void OnAddingToDrawing(Drawing drawing)
        {
            base.OnAddingToDrawing(drawing);
            // The undo system can attach this label to a point but not set the point's label property.
            // This ensures that the label property is set to this pointLabel.
            var point = GetPoint();
            if (point.Label == null)
            {
                point.Label = this;
            }
        }

        public override void MoveToCore(Point newPosition)
        {
            Point newOffset = newPosition.Minus(Point(0));
            newOffset = newOffset.TrimToMaxLength(ToLogical(100));
            Offset = newOffset;
            base.MoveToCore(newPosition);
        }

        protected override int DefaultZOrder()
        {
            return (int)ZOrder.PointLabels;
        }

        public override void UpdateVisual()
        {
            if (Dependencies.IsEmpty())
            {
                return;
            }

            var textWasEmpty = false;

            if (Text.IsEmpty())
            {
                textWasEmpty = true;
            }

            var coords = Point(0);
            MoveToCore(coords.Plus(Offset));
            base.UpdateVisual();
            UpdateText();

            if (textWasEmpty && !Text.IsEmpty())
            {
                selection.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                Offset = new Point(
                    ToLogical(-selection.DesiredSize.Width / 2),
                    ToLogical(selection.DesiredSize.Height
                        + GetPoint().Shape.ActualHeight / 2
                        + Math.CursorTolerance));
                MoveToCore(coords.Plus(Offset));
                base.UpdateVisual();
            }
        }

        private void UpdateText()
        {
            var text = "";
            if (ShowName)
            {
                var nameText = Dependencies.ElementAt(0).Name;
                text = nameText;
            }
            if (ShowCoordinates)
            {
                var coordinates = Point(0);
                var x = Math.Round(coordinates.X, DecimalsToShow);
                var y = Math.Round(coordinates.Y, DecimalsToShow);
                var coordinatesText = string.Format("({0};{1})", x , y );
                //var coordinatesText = string.Format("({0:0.0#};{1:0.0#})",
                //    coordinates.X,
                //    coordinates.Y);
                if (!text.IsEmpty())
                {
                    text += " ";
                }
                text += coordinatesText;
            }
            Text = text;
        }

        [PropertyGridVisible(false)]    // Handled in the point's property grid.
        public override string Text
        {
            get
            {
                return base.Text;
            }
        }

        bool showName;
        [PropertyGridVisible(false)]    // Handled in the point's property grid.
        public bool ShowName
        {
            get
            {
                return showName;
            }
            set
            {
                showName = value;
                UpdateVisual();
            }
        }

        bool showCoordinates;
        [PropertyGridVisible(false)]    // Handled in the point's property grid.
        public bool ShowCoordinates
        {
            get
            {
                return showCoordinates;
            }
            set
            {
                showCoordinates = value;
                UpdateVisual();
            }
        }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            base.ReadXml(element);
            this.showName = element.ReadBool("ShowName", true);
            this.showCoordinates = element.ReadBool("ShowCoordinates", false);
            UpdateText();
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            writer.WriteAttributeBool("ShowName", ShowName);
            writer.WriteAttributeBool("ShowCoordinates", ShowCoordinates);
        }

        [PropertyGridVisible]
        [PropertyGridName("Show name")]
        public bool ShowNameDisplay
        {
            get
            {
                return ShowName;
            }
            set
            {
                var point = Dependencies.ElementAt(0) as PointBase;
                if (point != null)
                {
                    point.ShowName = value;
                }
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Show coordinates")]
        public bool ShowCoordinatesDisplay
        {
            get
            {
                return ShowCoordinates;
            }
            set
            {
                var point = Dependencies.ElementAt(0) as PointBase;
                if (point != null)
                {
                    point.ShowCoordinates = value;
                }
            }
        }
    }
}
