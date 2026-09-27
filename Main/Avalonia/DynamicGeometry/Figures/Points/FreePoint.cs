using Avalonia;
using System.Xml.Linq;

namespace DynamicGeometry
{
    public partial class FreePoint : PointBase, IMovable, IConditionalProperties
    {
        /// <summary>
        /// Puts the point onto the one figure passing through it (<see cref="PointSnapping"/>).
        /// Shown only when there is exactly one, and captioned with its name; with more, the
        /// context menu lists them.
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridName("Snap to figure")]
        [PropertyGridIcon(PropertyGridIcon.Lock)]
        public void SnapToFigure()
        {
            var figures = PointSnapping.FiguresToSnapTo(this);
            if (figures.Count == 1)
            {
                PointSnapping.SnapTo(this, figures[0]);
            }
        }

        public bool CanEdit(string propertyName)
        {
            return propertyName != nameof(SnapToFigure) || PointSnapping.FiguresToSnapTo(this).Count == 1;
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            if (propertyName == nameof(SnapToFigure))
            {
                var figures = PointSnapping.FiguresToSnapTo(this);
                if (figures.Count == 1)
                {
                    return "Snap to " + PointSnapping.Describe(figures[0]);
                }
            }

            return defaultCaption;
        }

        public override void ReadXml(XElement element)
        {
            base.ReadXml(element);
            var x = element.ReadDouble("X");
            var y = element.ReadDouble("Y");
            Coordinates = new Point(x, y);
        }

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            var coordinates = Coordinates;
            writer.WriteAttributeDouble("X", coordinates.X);
            writer.WriteAttributeDouble("Y", coordinates.Y);
        }

        /// <summary>
        /// Perf optimization
        /// We know it exists, no need to call base
        /// </summary>
        public override void UpdateExistence()
        {
        }

        [PropertyGridVisible]
        public override double X
        {
            get
            {
                return base.X;
            }
            set
            {
                this.MoveTo(new Point(value, Y));
                this.RecalculateAllDependents();
            }
        }

        [PropertyGridVisible]
        public override double Y
        {
            get
            {
                return base.Y;
            }
            set
            {
                this.MoveTo(new Point(X, value));
                this.RecalculateAllDependents();
            }
        }
    }
}
