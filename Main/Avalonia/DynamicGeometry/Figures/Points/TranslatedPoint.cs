using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls.Shapes;

namespace DynamicGeometry
{
    /// <summary>
    /// A point at a magnitude and a direction from its source point. Each of the two is either
    /// tied to a figure it depends on - a vector (both at once), anything with a length or an
    /// angle, a <see cref="Number"/> holding a typed value - or free: then it is the parameter
    /// that dragging the point changes, kept in the file. A point with a free direction slides
    /// on the circle around its source, one with a free magnitude on the line through it
    /// (signed, so it passes through the source to the other side). Both free would be a free
    /// point on a leash; the property grid doesn't allow it.
    /// </summary>
    public class TranslatedPoint : PointBase, IPoint, IConditionalProperties
    {
        /// <summary>
        /// One of the two quantities: the index of its source in the dependency list (-1 when
        /// free), the parameter used while free (magnitude in units, direction in radians) and
        /// the Number that held it before it was freed, so that fixing it again - undo included -
        /// brings the same one back under the same name.
        /// </summary>
        class Quantity
        {
            public int SourceIndex = -1;
            public double Parameter;
            public Number Retired;
        }

        readonly Quantity magnitudeQuantity = new Quantity();
        readonly Quantity directionQuantity = new Quantity();

        public IPoint Source
        {
            get
            {
                return Dependencies.Count >= 1 ? Dependencies[0] as IPoint : null;
            }
        }

        /// <summary>A vector, a length provider or a Number; null while the magnitude is free</summary>
        public IFigure MagnitudeSource
        {
            get { return SourceOf(magnitudeQuantity); }
        }

        /// <summary>A vector, an angle provider or a Number; null while the direction is free</summary>
        public IFigure DirectionSource
        {
            get { return SourceOf(directionQuantity); }
        }

        IFigure SourceOf(Quantity quantity)
        {
            int index = quantity.SourceIndex;
            return index >= 0 && index < Dependencies.Count ? Dependencies[index] : null;
        }

        public bool IsMagnitudeFree
        {
            get { return magnitudeQuantity.SourceIndex < 0; }
        }

        public bool IsDirectionFree
        {
            get { return directionQuantity.SourceIndex < 0; }
        }

        /// <summary>Dragging changes something: this is a draggable point, styled like one</summary>
        public bool HasFreedom
        {
            get { return IsMagnitudeFree || IsDirectionFree; }
        }

        /// <summary>
        /// The dependencies: the source point, then the magnitude source and the direction
        /// source when there are any (one entry when a vector is both). A null source leaves
        /// that quantity free, at whatever value it has.
        /// </summary>
        public void SetSources(IPoint source, IFigure magnitudeSource, IFigure directionSource)
        {
            var dependencies = new List<IFigure>() { source };
            magnitudeQuantity.SourceIndex = -1;
            directionQuantity.SourceIndex = -1;
            if (magnitudeSource != null)
            {
                magnitudeQuantity.SourceIndex = dependencies.Count;
                dependencies.Add(magnitudeSource);
            }

            if (directionSource != null)
            {
                if (directionSource == magnitudeSource)
                {
                    directionQuantity.SourceIndex = magnitudeQuantity.SourceIndex;
                }
                else
                {
                    directionQuantity.SourceIndex = dependencies.Count;
                    dependencies.Add(directionSource);
                }
            }

            Dependencies = dependencies;
        }

        double MagnitudeValue
        {
            get
            {
                var source = MagnitudeSource;
                if (source is Vector vector)
                {
                    return vector.Magnitude;
                }

                if (source is ILengthProvider length)
                {
                    return length.Length;
                }

                return magnitudeQuantity.Parameter;
            }
        }

        double DirectionRadians
        {
            get
            {
                var source = DirectionSource;
                if (source is Vector vector)
                {
                    return vector.Direction;
                }

                if (source is IAngleProvider angle)
                {
                    return angle.Angle;
                }

                return directionQuantity.Parameter;
            }
        }

        #region Property grid

        // a dragged parameter has all the digits of a double; the grid shows a few
        const int ShownDecimals = 4;

        [PropertyGridVisible]
        [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
        public double Magnitude
        {
            get
            {
                return System.Math.Round(MagnitudeValue, ShownDecimals);
            }
            set
            {
                SetQuantity(magnitudeQuantity, value);
            }
        }

        /// <summary>In degrees, counterclockwise from the x axis</summary>
        [PropertyGridVisible]
        [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
        public double Direction
        {
            get
            {
                return System.Math.Round(DirectionRadians.ToDegrees(), ShownDecimals);
            }
            set
            {
                SetQuantity(directionQuantity, value.ToRadians());
            }
        }

        /// <summary>
        /// A free quantity takes the value as its parameter; one tied to a Number writes it
        /// there (the Number recalculates its dependents, this point among them); one tied to
        /// anything else can't be set, and the grid knows.
        /// </summary>
        void SetQuantity(Quantity quantity, double value)
        {
            var source = SourceOf(quantity);
            if (source is Number number)
            {
                number.Value = quantity == directionQuantity ? value.ToDegrees() : value;
                return;
            }

            if (source != null)
            {
                return;
            }

            quantity.Parameter = value;
            Recalculate();
            UpdateVisual();
            this.RecalculateAllDependents();
        }

        [PropertyGridVisible]
        [PropertyGridName("Free magnitude")]
        [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
        public bool FreeMagnitude
        {
            get
            {
                return IsMagnitudeFree;
            }
            set
            {
                if (value)
                {
                    Free(magnitudeQuantity, MagnitudeValue);
                }
                else
                {
                    Fix(magnitudeQuantity, MagnitudeValue);
                }
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Free direction")]
        [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
        public bool FreeDirection
        {
            get
            {
                return IsDirectionFree;
            }
            set
            {
                if (value)
                {
                    Free(directionQuantity, DirectionRadians);
                }
                else
                {
                    Fix(directionQuantity, DirectionRadians);
                }
            }
        }

        /// <summary>
        /// Which rows the grid lets the user edit: a value when it is free or held by a Number,
        /// a Free checkbox unless ticking it would free both quantities.
        /// </summary>
        public bool CanEdit(string propertyName)
        {
            switch (propertyName)
            {
                case "Magnitude":
                    return IsMagnitudeFree || MagnitudeSource is Number;
                case "Direction":
                    return IsDirectionFree || DirectionSource is Number;
                case "FreeMagnitude":
                    return IsMagnitudeFree || !IsDirectionFree;
                case "FreeDirection":
                    return IsDirectionFree || !IsMagnitudeFree;
                default:
                    return true;
            }
        }

        /// <summary>A value row tied to a figure other than a Number says which</summary>
        public string Caption(string propertyName, string defaultCaption)
        {
            var source = propertyName == "Magnitude" ? MagnitudeSource : propertyName == "Direction" ? DirectionSource : null;
            return source == null || source is Number ? defaultCaption : defaultCaption + " = " + source.Name;
        }

        /// <summary>
        /// The quantity's current value becomes its parameter and its source is dropped. A
        /// Number that held a typed value for this point alone leaves the drawing with it,
        /// kept aside for <see cref="Fix"/>. Nothing moves on screen.
        /// </summary>
        void Free(Quantity quantity, double currentValue)
        {
            int index = quantity.SourceIndex;
            if (index < 0)
            {
                return;
            }

            var source = Dependencies[index];
            var other = quantity == magnitudeQuantity ? directionQuantity : magnitudeQuantity;
            quantity.Parameter = currentValue;
            quantity.SourceIndex = -1;
            if (other.SourceIndex != index)
            {
                if (other.SourceIndex > index)
                {
                    other.SourceIndex--;
                }

                this.RemoveDependencyCore(index, source);
                if (source is Number number && number.Auxiliary && number.Dependents.IsEmpty() && Drawing != null)
                {
                    Drawing.Figures.Remove(number);
                    quantity.Retired = number;
                }
            }

            OnFreedomChanged();
        }

        /// <summary>
        /// The quantity gets a Number holding its current value: the one it had before being
        /// freed, if any, otherwise a new auxiliary one, placed in the figure list before this
        /// point. Nothing moves on screen.
        /// </summary>
        void Fix(Quantity quantity, double currentValue)
        {
            if (quantity.SourceIndex >= 0 || Drawing == null)
            {
                return;
            }

            var number = quantity.Retired ?? Number.CreateAuxiliary(Drawing, currentValue);
            quantity.Retired = null;
            number.Value = quantity == directionQuantity ? currentValue.ToDegrees() : currentValue;
            if (!Drawing.Figures.Contains(number))
            {
                int place = Drawing.Figures.IndexOf(this);
                Drawing.Figures.Insert(place < 0 ? Drawing.Figures.Count : place, number);
            }

            quantity.SourceIndex = Dependencies.Count;
            this.InsertDependencyCore(quantity.SourceIndex, number);
            OnFreedomChanged();
        }

        void OnFreedomChanged()
        {
            Recalculate();
            UpdateVisual();
            this.RecalculateAllDependents();
            UpdateStyleForFreedom();
            RaisePropertyChanged("Magnitude");
            RaisePropertyChanged("Direction");
            RaisePropertyChanged("FreeMagnitude");
            RaisePropertyChanged("FreeDirection");
        }

        /// <summary>
        /// A point that can be dragged looks like one (the green PointOnFigure style), a fully
        /// determined one like a dependent point - unless the user gave it a style of their own.
        /// </summary>
        void UpdateStyleForFreedom()
        {
            var manager = Drawing?.StyleManager;
            if (manager == null)
            {
                return;
            }

            if (Style == null
                || Style.Name == StyleManager.FreePointStyleName
                || Style.Name == StyleManager.PointOnFigureStyleName
                || Style.Name == StyleManager.DependentPointStyleName)
            {
                Style = manager.AssignDefaultStyle(this);
            }
        }

        #endregion

        #region Dragging

        public override bool AllowMove()
        {
            return !Locked && HasFreedom;
        }

        /// <summary>
        /// The free quantity follows the cursor: the direction as the angle from the source, the
        /// magnitude as the projection onto the line through the source, signed.
        /// </summary>
        public override void MoveToCore(Point newPosition)
        {
            var source = Source;
            if (source == null)
            {
                return;
            }

            var origin = source.Coordinates;
            if (IsDirectionFree)
            {
                directionQuantity.Parameter = Math.GetAngle(origin, newPosition);
                RaisePropertyChanged("Direction");
            }

            if (IsMagnitudeFree)
            {
                double direction = DirectionRadians;
                magnitudeQuantity.Parameter =
                    (newPosition.X - origin.X) * System.Math.Cos(direction)
                    + (newPosition.Y - origin.Y) * System.Math.Sin(direction);
                RaisePropertyChanged("Magnitude");
            }

            Recalculate();
        }

        #endregion

        protected override Shape CreateShape()
        {
            return Factory.CreateDependentPointShape();
        }

        public override void Recalculate()
        {
            var source = Source;
            if (source == null || !Dependencies.Exists())
            {
                Exists = false;
                return;
            }

            Coordinates = Math.GetTranslationPoint(source.Coordinates, MagnitudeValue, DirectionRadians);
            Exists = Coordinates.Exists();
        }

        #region Serialization

        // Files from before 2026-09-25 said which dependency was which by type and position, a
        // typed value was an attribute, and Direction was in radians. Read that way when none of
        // the new attributes is there; the typed values then become Numbers once the drawing is
        // loaded (UpgradeLegacyValues).
        bool needsLegacyUpgrade;

        public override void ReadXml(XElement element)
        {
            base.ReadXml(element);
            var magnitudeSource = element.ReadString("MagnitudeSource");
            var directionSource = element.ReadString("DirectionSource");
            bool newFormat = magnitudeSource != null
                || directionSource != null
                || element.Attribute("FreeMagnitude") != null
                || element.Attribute("FreeDirection") != null;
            if (newFormat)
            {
                magnitudeQuantity.SourceIndex = IndexOfDependency(magnitudeSource);
                directionQuantity.SourceIndex = IndexOfDependency(directionSource);
                magnitudeQuantity.Parameter = element.ReadDouble("Magnitude");
                directionQuantity.Parameter = element.ReadDouble("Direction").ToRadians();
            }
            else
            {
                if (Dependencies.Count > 1 && Dependencies[1] is Vector)
                {
                    magnitudeQuantity.SourceIndex = 1;
                    directionQuantity.SourceIndex = 1;
                }
                else
                {
                    magnitudeQuantity.SourceIndex = Dependencies.Count > 1 && Dependencies[1] is ILengthProvider ? 1 : -1;
                    directionQuantity.SourceIndex = Dependencies.Count > 2 && Dependencies[2] is IAngleProvider ? 2 : -1;
                }

                magnitudeQuantity.Parameter = element.ReadDouble("Magnitude");
                directionQuantity.Parameter = element.ReadDouble("Direction");
                needsLegacyUpgrade = HasFreedom;
            }

            Recalculate();
        }

        int IndexOfDependency(string name)
        {
            if (name == null)
            {
                return -1;
            }

            for (int i = 0; i < Dependencies.Count; i++)
            {
                if (Dependencies[i].Name == name)
                {
                    return i;
                }
            }

            return -1;
        }

        /// <summary>
        /// A typed value from an old file was fixed, so it gets a Number like a typed value
        /// does now. Called by the deserializer once the point is in the drawing.
        /// </summary>
        public void UpgradeLegacyValues()
        {
            if (!needsLegacyUpgrade)
            {
                return;
            }

            needsLegacyUpgrade = false;
            if (IsMagnitudeFree)
            {
                Fix(magnitudeQuantity, MagnitudeValue);
            }

            if (IsDirectionFree)
            {
                Fix(directionQuantity, DirectionRadians);
            }
        }

        public override void WriteXml(XmlWriter writer)
        {
            base.WriteXml(writer);
            var magnitudeSource = MagnitudeSource;
            if (magnitudeSource != null)
            {
                writer.WriteAttributeString("MagnitudeSource", magnitudeSource.Name);
            }
            else
            {
                writer.WriteAttributeBool("FreeMagnitude", true);
                writer.WriteAttributeDouble("Magnitude", magnitudeQuantity.Parameter);
            }

            var directionSource = DirectionSource;
            if (directionSource != null)
            {
                writer.WriteAttributeString("DirectionSource", directionSource.Name);
            }
            else
            {
                writer.WriteAttributeBool("FreeDirection", true);
                writer.WriteAttributeDouble("Direction", directionQuantity.Parameter.ToDegrees());
            }
        }

        #endregion
    }
}
