using System.Collections.Generic;
using System.Linq;
using Avalonia.Controls.Shapes;

namespace DynamicGeometry
{
    /// <summary>
    /// A point turned about a center by an angle. The angle is a figure it depends on: a
    /// <see cref="Number"/> holding a typed value, shared by every point of a rotated figure,
    /// or anything with an angle (an angle measurement, its arc, a slider), which the point
    /// then follows (<see cref="ITiedValues"/>). Dependencies: the source, the center, the angle.
    /// </summary>
    public class RotatedPoint : PointBase, IPoint, ITiedValues, IConditionalProperties
    {
        public IPoint Source
        {
            get
            {
                return Dependencies.Count >= 1 ? Dependencies[0] as IPoint : null;
            }
        }

        public new IPoint Center
        {
            get
            {
                return Dependencies.Count >= 2 ? Dependencies[1] as IPoint : null;
            }
        }

        /// <summary>A Number or an angle provider</summary>
        public IFigure AngleSource
        {
            get
            {
                return Dependencies.Count >= 3 ? Dependencies[2] : null;
            }
        }

        /// <summary>In degrees, counterclockwise; editable when it is a typed value, which lives in the Number</summary>
        [PropertyGridVisible]
        [PropertyGridName("Angle (degrees)")]
        [PropertyGridGroup("Angle")]
        [PropertyGridPreferredEditor("UpDown")]
        [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
        public double Angle
        {
            get
            {
                switch (AngleSource)
                {
                    case Number number:
                        return number.Value;
                    case IAngleProvider provider:
                        return provider.Angle.ToDegrees();
                    default:
                        return 0;
                }
            }
            set
            {
                if (AngleSource is Number number)
                {
                    number.Value = value;
                }
            }
        }

        /// <summary>The grid's button back from a tied angle (<see cref="Detach"/>), shown while the angle is a figure's</summary>
        [PropertyGridVisible]
        [PropertyGridName("Type the angle")]
        [PropertyGridGroup("Angle")]
        [PropertyGridIcon(PropertyGridIcon.Pencil)]
        public void UntieAngle()
        {
            if (Detach(nameof(Angle)))
            {
                Drawing.RaiseDisplayProperties(this);
            }
        }

        public bool CanEdit(string propertyName)
        {
            switch (propertyName)
            {
                case nameof(Angle):
                    return AngleSource is Number;
                case nameof(UntieAngle):
                    return this.IsTied(nameof(Angle));
                default:
                    return true;
            }
        }

        /// <summary>An angle taken from a figure says which</summary>
        public string Caption(string propertyName, string defaultCaption)
        {
            return propertyName == nameof(Angle) && this.IsTied(propertyName) ? "Angle = " + TiedValues.SourceName(AngleSource) : defaultCaption;
        }

        protected override Shape CreateShape()
        {
            return Factory.CreateDependentPointShape();
        }

        public override void Recalculate()
        {
            if (Source != null && Center != null)
            {
                Coordinates = Math.GetRotationPoint(Source.Coordinates, Center.Coordinates, Math.ToRadians(Angle));
            }

            Exists = Coordinates.Exists();
        }

        #region Tied values

        public IEnumerable<string> TiedValueNames
        {
            get { yield return nameof(Angle); }
        }

        public IFigure GetSource(string name)
        {
            return AngleSource;
        }

        public bool Accepts(string name, IFigure figure)
        {
            // (DynamicGeometry.Label: here Label is the point's own name label)
            return figure is IAngleProvider && !(figure is IPoint) && DynamicGeometry.Label.GivesNumber(figure);
        }

        /// <summary>The points turned about the same center by the same angle: the vertices of one rotated figure, tied and detached together</summary>
        IList<IFigure> Siblings
        {
            get
            {
                var source = AngleSource;
                if (source == null)
                {
                    return new IFigure[] { this };
                }

                return source.Dependents
                    .OfType<RotatedPoint>()
                    .Where(point => point.Center == Center)
                    .Cast<IFigure>()
                    .ToList();
            }
        }

        public bool TieTo(string name, IFigure source)
        {
            return TiedValues.Tie(Siblings, AngleSource, source);
        }

        public bool Detach(string name)
        {
            return this.IsTied(name) && TieTo(name, Number.CreateAuxiliary(Drawing, Angle));
        }

        #endregion
    }
}
