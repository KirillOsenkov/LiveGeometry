using System.Collections.Generic;
using System.Linq;
using Avalonia.Controls.Shapes;

namespace DynamicGeometry
{
    /// <summary>
    /// A point stretched from a center by a factor. The factor is a figure it depends on: a
    /// <see cref="Number"/> holding a typed value, shared by every point of a dilated figure,
    /// or anything with a length (a segment, a slider), which the point then follows
    /// (<see cref="ITiedValues"/>). Dependencies: the source, the center, the factor - and
    /// a fourth when the factor is the ratio of two lengths (no tool makes one; the reader
    /// of GeoGebra files could).
    /// </summary>
    public class DilatedPoint : PointBase, IPoint, ITiedValues, IConditionalProperties
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

        /// <summary>A Number or a length provider</summary>
        public IFigure FactorSource
        {
            get
            {
                return Dependencies.Count >= 3 ? Dependencies[2] : null;
            }
        }

        /// <summary>The factor is the ratio of two lengths</summary>
        bool IsRatio
        {
            get { return Dependencies.Count >= 4; }
        }

        /// <summary>Editable when it is a typed value, which lives in the Number</summary>
        [PropertyGridVisible]
        [PropertyGridGroup("Factor")]
        [PropertyGridPreferredEditor("UpDown")]
        [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
        public double Factor
        {
            get
            {
                if (IsRatio)
                {
                    double denominator = (Dependencies[3] as ILengthProvider)?.Length ?? 1;
                    return denominator != 0 ? ((Dependencies[2] as ILengthProvider)?.Length ?? 1) / denominator : 99999;
                }

                switch (FactorSource)
                {
                    case Number number:
                        return number.Value;
                    case ILengthProvider provider:
                        return provider.Length;
                    default:
                        return 1;
                }
            }
            set
            {
                if (FactorSource is Number number)
                {
                    number.Value = value;
                }
            }
        }

        /// <summary>The grid's button back from a tied factor (<see cref="Detach"/>), shown while the factor is a figure's</summary>
        [PropertyGridVisible]
        [PropertyGridName("Type the factor")]
        [PropertyGridGroup("Factor")]
        [PropertyGridIcon(PropertyGridIcon.Pencil)]
        public void UntieFactor()
        {
            if (Detach(nameof(Factor)))
            {
                Drawing.RaiseDisplayProperties(this);
            }
        }

        public bool CanEdit(string propertyName)
        {
            switch (propertyName)
            {
                case nameof(Factor):
                    return FactorSource is Number;
                case nameof(UntieFactor):
                    return this.IsTied(nameof(Factor));
                default:
                    return true;
            }
        }

        /// <summary>A factor taken from a figure says which</summary>
        public string Caption(string propertyName, string defaultCaption)
        {
            if (propertyName != nameof(Factor) || !this.IsTied(propertyName))
            {
                return defaultCaption;
            }

            return IsRatio
                ? "Factor = " + TiedValues.SourceName(Dependencies[2]) + " / " + TiedValues.SourceName(Dependencies[3])
                : "Factor = " + TiedValues.SourceName(FactorSource);
        }

        protected override Shape CreateShape()
        {
            return Factory.CreateDependentPointShape();
        }

        public override void Recalculate()
        {
            if (Source != null && Center != null)
            {
                Coordinates = Math.GetDilationPoint(Source.Coordinates, Center.Coordinates, Factor);
            }

            // (see ReflectedPoint: no image of what is not there)
            Exists = Dependencies.Exists() && Coordinates.Exists();
        }

        #region Tied values

        public IEnumerable<string> TiedValueNames
        {
            get { yield return nameof(Factor); }
        }

        public IFigure GetSource(string name)
        {
            return FactorSource;
        }

        public bool Accepts(string name, IFigure figure)
        {
            // (DynamicGeometry.Label: here Label is the point's own name label)
            return !(figure is IPoint) && figure.GivesLength();
        }

        /// <summary>The points stretched from the same center by the same factor: the vertices of one dilated figure, tied and detached together</summary>
        IList<IFigure> Siblings
        {
            get
            {
                var source = FactorSource;
                if (source == null)
                {
                    return new IFigure[] { this };
                }

                return source.Dependents
                    .OfType<DilatedPoint>()
                    .Where(point => point.Center == Center)
                    .Cast<IFigure>()
                    .ToList();
            }
        }

        // a ratio of two lengths is left as it is
        public bool TieTo(string name, IFigure source)
        {
            return !IsRatio && TiedValues.Tie(Siblings, FactorSource, source);
        }

        public bool Detach(string name)
        {
            return this.IsTied(name) && TieTo(name, Number.CreateAuxiliary(Drawing, Factor));
        }

        #endregion
    }
}
