using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls.Shapes;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    /// <summary>
    /// A point at a distance and a direction from its source point. Each of the two is either
    /// tied to a figure it depends on - a vector (both at once), anything with a length or an
    /// angle, a <see cref="Number"/> holding a typed value - or free: then it is the parameter
    /// that dragging the point changes, kept in the file. A point with a free direction slides
    /// on the circle around its source, one with a free distance on the line through it
    /// (signed, so it passes through the source to the other side). Both free would be a free
    /// point on a leash; the property grid doesn't allow it.
    /// </summary>
    public class TranslatedPoint : PointBase, IPoint, ITiedValues, IConditionalProperties
    {
        /// <summary>
        /// One of the two quantities: the index of its source in the dependency list (-1 when
        /// free), the parameter used while free (distance in units, direction in radians) and
        /// the Number that held it before it was freed, so that fixing it again - undo included -
        /// brings the same one back under the same name.
        /// </summary>
        class Quantity
        {
            public int SourceIndex = -1;
            public double Parameter;
            public Number Retired;
        }

        readonly Quantity distanceQuantity = new Quantity();
        readonly Quantity directionQuantity = new Quantity();

        public IPoint Source
        {
            get
            {
                return Dependencies.Count >= 1 ? Dependencies[0] as IPoint : null;
            }
        }

        /// <summary>A vector, a length provider or a Number; null while the distance is free</summary>
        public IFigure DistanceSource
        {
            get { return SourceOf(distanceQuantity); }
        }

        /// <summary>A vector, a line (its oriented direction), an angle provider or a Number; null while the direction is free</summary>
        public IFigure DirectionSource
        {
            get { return SourceOf(directionQuantity); }
        }

        IFigure SourceOf(Quantity quantity)
        {
            int index = quantity.SourceIndex;
            return index >= 0 && index < Dependencies.Count ? Dependencies[index] : null;
        }

        public bool IsDistanceFree
        {
            get { return distanceQuantity.SourceIndex < 0; }
        }

        public bool IsDirectionFree
        {
            get { return directionQuantity.SourceIndex < 0; }
        }

        /// <summary>Dragging changes something: this is a draggable point, styled like one</summary>
        public bool HasFreedom
        {
            get { return IsDistanceFree || IsDirectionFree; }
        }

        /// <summary>
        /// The dependencies: the source point, then the distance source and the direction
        /// source when there are any (one entry when a vector is both). A null source leaves
        /// that quantity free, at whatever value it has.
        /// </summary>
        public void SetSources(IPoint source, IFigure distanceSource, IFigure directionSource)
        {
            var dependencies = new List<IFigure>() { source };
            distanceQuantity.SourceIndex = -1;
            directionQuantity.SourceIndex = -1;
            if (distanceSource != null)
            {
                distanceQuantity.SourceIndex = dependencies.Count;
                dependencies.Add(distanceSource);
            }

            if (directionSource != null)
            {
                if (directionSource == distanceSource)
                {
                    directionQuantity.SourceIndex = distanceQuantity.SourceIndex;
                }
                else
                {
                    directionQuantity.SourceIndex = dependencies.Count;
                    dependencies.Add(directionSource);
                }
            }

            Dependencies = dependencies;
        }

        double DistanceValue
        {
            get
            {
                var source = DistanceSource;
                if (source is Vector vector)
                {
                    return vector.Length;
                }

                if (source is ILengthProvider length)
                {
                    return length.Length;
                }

                return distanceQuantity.Parameter;
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

                // a segment, ray or line: the way it points, from its first point to its second
                if (source is ILine line)
                {
                    return Math.GetAngle(line.Coordinates.P1, line.Coordinates.P2);
                }

                if (source is IAngleProvider angle)
                {
                    return angle.Angle;
                }

                return directionQuantity.Parameter;
            }
        }

        #region Property grid

        [PropertyGridVisible]
        [PropertyGridGroup("Distance")]
        [PropertyGridPreferredEditor("UpDown")]
        [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
        public double Distance
        {
            get
            {
                return DistanceValue;
            }
            set
            {
                SetQuantity(distanceQuantity, value);
            }
        }

        /// <summary>In degrees, counterclockwise from the x axis</summary>
        [PropertyGridVisible]
        [PropertyGridGroup("Direction")]
        [PropertyGridPreferredEditor("UpDown")]
        [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
        public double Direction
        {
            get
            {
                return DirectionRadians.ToDegrees();
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

        /// <summary>
        /// The Free distance and Free direction rows. Undo of a tick puts back the source the
        /// quantity had, whatever it was - unticking the box gives it a Number, which is not
        /// what a distance taken from a vector had.
        /// </summary>
        public class FreedomValue : ConditionalPropertyValue, IRestorableValue
        {
            class SavedSource
            {
                public IFigure Source;
                public int Index;
            }

            TranslatedPoint Owner
            {
                get { return (TranslatedPoint)Parent; }
            }

            Quantity OwnerQuantity
            {
                get { return Property.Name == nameof(FreeDistance) ? Owner.distanceQuantity : Owner.directionQuantity; }
            }

            /// <summary>The source of the quantity and its place among the dependencies; no source while it is free</summary>
            public object CaptureState()
            {
                var quantity = OwnerQuantity;
                return new SavedSource() { Source = Owner.SourceOf(quantity), Index = quantity.SourceIndex };
            }

            public void RestoreState(object state)
            {
                var saved = (SavedSource)state;
                var owner = Owner;
                var quantity = OwnerQuantity;
                if (saved.Source == null)
                {
                    owner.Free(quantity, quantity == owner.distanceQuantity ? owner.DistanceValue : owner.DirectionRadians);
                }
                else if (quantity.SourceIndex < 0)
                {
                    owner.Reattach(quantity, saved.Source, saved.Index);
                }
            }
        }

        /// <summary>
        /// The quantity follows the source again, which is listed among the dependencies
        /// where it was; a Number that left the drawing when the quantity was freed comes
        /// back to its place. For undo of a tick of a Free box.
        /// </summary>
        void Reattach(Quantity quantity, IFigure source, int index)
        {
            if (source is Number number && Drawing != null && !Drawing.Figures.Contains(number))
            {
                Drawing.Figures.ReturnBefore(number, this);
                if (quantity.Retired == number)
                {
                    quantity.Retired = null;
                }
            }

            SetSource(quantity, source, quantity.Parameter, index);
        }

        [PropertyGridVisible]
        [PropertyGridName("Free distance")]
        [PropertyGridGroup("Distance")]
        [PropertyGridCustomValueProvider(typeof(FreedomValue))]
        public bool FreeDistance
        {
            get
            {
                return IsDistanceFree;
            }
            set
            {
                if (value)
                {
                    Free(distanceQuantity, DistanceValue);
                }
                else
                {
                    Fix(distanceQuantity, DistanceValue);
                }
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Free direction")]
        [PropertyGridGroup("Direction")]
        [PropertyGridCustomValueProvider(typeof(FreedomValue))]
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

        /// <summary>The way back from a distance taken from a figure: a typed value again (<see cref="Detach"/>)</summary>
        [PropertyGridVisible]
        [PropertyGridName("Type the distance")]
        [PropertyGridGroup("Distance")]
        [PropertyGridIcon(PropertyGridIcon.Pencil)]
        public void UntieDistance()
        {
            if (Detach(nameof(Distance)))
            {
                Drawing.RaiseDisplayProperties(this);
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Type the direction")]
        [PropertyGridGroup("Direction")]
        [PropertyGridIcon(PropertyGridIcon.Pencil)]
        public void UntieDirection()
        {
            if (Detach(nameof(Direction)))
            {
                Drawing.RaiseDisplayProperties(this);
            }
        }

        /// <summary>
        /// Which rows the grid lets the user edit: a value when it is free or held by a Number,
        /// a Free checkbox unless ticking it would free both quantities, a "Type the ..."
        /// button while the value is a figure's.
        /// </summary>
        public bool CanEdit(string propertyName)
        {
            switch (propertyName)
            {
                case "Distance":
                    return IsDistanceFree || DistanceSource is Number;
                case "Direction":
                    return IsDirectionFree || DirectionSource is Number;
                case nameof(UntieDistance):
                    return this.IsTied(nameof(Distance));
                case nameof(UntieDirection):
                    return this.IsTied(nameof(Direction));
                case "FreeDistance":
                    return IsDistanceFree || !IsDirectionFree;
                case "FreeDirection":
                    return IsDirectionFree || !IsDistanceFree;
                default:
                    return true;
            }
        }

        /// <summary>A value row tied to a figure other than a Number says which</summary>
        public string Caption(string propertyName, string defaultCaption)
        {
            var source = propertyName == "Distance" ? DistanceSource : propertyName == "Direction" ? DirectionSource : null;
            if (source == null || source is Number)
            {
                return defaultCaption;
            }

            return defaultCaption + " = " + TiedValues.SourceName(source);
        }

        protected override string Kind
        {
            get
            {
                return "Translated point";
            }
        }

        /// <summary>
        /// "of A by vector u", "of A by 3 at 30°", "of A along line g" (a free distance), "of
        /// A by a" (a free direction)
        /// </summary>
        public override string Construction
        {
            get
            {
                if (Source == null)
                {
                    return null;
                }

                var result = "of " + ConstructionText.Of(Source);
                var distance = DistanceSource;
                var direction = DirectionSource;
                if (distance != null && distance == direction)
                {
                    return result + " by " + ConstructionText.Of(distance);
                }

                if (distance != null)
                {
                    result += " by " + (distance is Vector ? "the length of " + ConstructionText.Of(distance) : ConstructionText.Length(distance));
                }

                switch (direction)
                {
                    case null:
                        return result;
                    case Vector:
                        return result + " in the direction of " + ConstructionText.Of(direction);
                    case ILine:
                        return result + " along " + ConstructionText.Of(direction);
                    default:
                        return result + " at " + ConstructionText.AngleValue(direction);
                }
            }
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
            var other = quantity == distanceQuantity ? directionQuantity : distanceQuantity;
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
                    Drawing.Figures.Retire(number);
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
                Drawing.Figures.ReturnBefore(number, this);
            }

            quantity.SourceIndex = Dependencies.Count;
            this.InsertDependencyCore(quantity.SourceIndex, number);
            OnFreedomChanged();
        }

        // the rows of the grid change shape, not just their values: which are editable, what
        // "Direction = ..." names, whether "Type the direction" is there - so the grid rebuilds
        void OnFreedomChanged()
        {
            Recalculate();
            UpdateVisual();
            this.RecalculateAllDependents();
            UpdateStyleForFreedom();
            RaisePropertyChanged(null);
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

        #region Tied values

        public IEnumerable<string> TiedValueNames
        {
            get
            {
                yield return nameof(Distance);
                yield return nameof(Direction);
            }
        }

        Quantity QuantityOf(string name)
        {
            return name == nameof(Distance) ? distanceQuantity : directionQuantity;
        }

        public IFigure GetSource(string name)
        {
            return SourceOf(QuantityOf(name));
        }

        /// <summary>What the Translate tool takes for each: a vector for both, a length for the distance, an angle or a line as it points for the direction</summary>
        public bool Accepts(string name, IFigure figure)
        {
            // (nor a label that says no number: a caption is no length and no angle;
            // DynamicGeometry.Label, since here Label is the point's own name label)
            if (figure is IPoint || !DynamicGeometry.Label.GivesNumber(figure))
            {
                return false;
            }

            return name == nameof(Distance)
                ? figure is Vector || figure.GivesLength()
                : figure is Vector || figure is ILine || figure is IAngleProvider;
        }

        /// <summary>
        /// The points translated by the same sources: the vertices of one translated figure,
        /// tied and detached together. A free quantity is each point's own.
        /// </summary>
        IList<IFigure> Siblings(Quantity quantity)
        {
            var source = SourceOf(quantity);
            if (source == null)
            {
                return new IFigure[] { this };
            }

            var distance = DistanceSource;
            var direction = DirectionSource;
            return source.Dependents
                .OfType<TranslatedPoint>()
                .Where(point => point.DistanceSource == distance && point.DirectionSource == direction)
                .Cast<IFigure>()
                .ToList();
        }

        /// <summary>
        /// One quantity from another source (<see cref="TiedValues.Tie"/>). The swap is
        /// <see cref="SetSource"/> on each sibling, recorded with its undo: a source shared
        /// with the other quantity (a vector) or a free quantity has no dependency to replace
        /// one for one.
        /// </summary>
        public bool TieTo(string name, IFigure source)
        {
            var quantity = QuantityOf(name);
            var old = SourceOf(quantity);
            return TiedValues.Tie(Siblings(quantity), old, source, owner =>
            {
                var point = (TranslatedPoint)owner;
                var pointQuantity = point.QuantityOf(name);
                var before = point.SourceOf(pointQuantity);
                int placeBefore = pointQuantity.SourceIndex;
                double parameter = pointQuantity.Parameter;
                point.Drawing.ActionManager.RecordAction(new CallMethodAction(
                    () => point.SetSource(pointQuantity, source, parameter),
                    () => point.SetSource(pointQuantity, before, parameter, placeBefore)));
            });
        }

        public bool Detach(string name)
        {
            if (!this.IsTied(name))
            {
                return false;
            }

            double value = QuantityOf(name) == distanceQuantity ? DistanceValue : DirectionRadians.ToDegrees();
            return TieTo(name, Number.CreateAuxiliary(Drawing, value));
        }

        /// <summary>
        /// The quantity's source from now on - null leaves it free, at <paramref name="parameter"/> -
        /// with the dependency list kept in step: a source the other quantity has too is shared
        /// (a vector), a dependency only this quantity had is dropped. Not recorded itself.
        /// </summary>
        /// <param name="place">Where among the dependencies a source not listed yet goes; at the end when not given (undo gives where it was)</param>
        void SetSource(
            Quantity quantity,
            IFigure source,
            double parameter,
            int place = -1)
        {
            var other = quantity == distanceQuantity ? directionQuantity : distanceQuantity;
            int index = quantity.SourceIndex;
            if (index >= 0)
            {
                quantity.SourceIndex = -1;
                if (other.SourceIndex != index)
                {
                    if (other.SourceIndex > index)
                    {
                        other.SourceIndex--;
                    }

                    this.RemoveDependencyCore(index, Dependencies[index]);
                }
            }

            quantity.Parameter = parameter;
            if (source != null)
            {
                int existing = Dependencies.IndexOf(source);
                if (existing < 0)
                {
                    existing = place >= 1 && place <= Dependencies.Count ? place : Dependencies.Count;
                    if (other.SourceIndex >= existing)
                    {
                        other.SourceIndex++;
                    }

                    this.InsertDependencyCore(existing, source);
                }

                quantity.SourceIndex = existing;
            }

            OnFreedomChanged();
        }

        #endregion

        #region Dragging

        public override bool AllowMove()
        {
            return !Locked && HasFreedom;
        }

        /// <summary>
        /// The free quantity follows the cursor: the direction as the angle from the source, the
        /// distance as the projection onto the line through the source, signed.
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

            if (IsDistanceFree)
            {
                double direction = DirectionRadians;
                distanceQuantity.Parameter =
                    (newPosition.X - origin.X) * System.Math.Cos(direction)
                    + (newPosition.Y - origin.Y) * System.Math.Sin(direction);
                RaisePropertyChanged("Distance");
            }

            Recalculate();
        }

        /// <summary>
        /// The place is the two parameters: put back as they were, where moving back by the
        /// offset of a drag would project the point onto its circle or line anew, somewhere else
        /// </summary>
        public override object CapturePlace()
        {
            return new Point(distanceQuantity.Parameter, directionQuantity.Parameter);
        }

        public override void RestorePlace(object place)
        {
            var parameters = (Point)place;
            distanceQuantity.Parameter = parameters.X;
            directionQuantity.Parameter = parameters.Y;
            RaisePropertyChanged("Distance");
            RaisePropertyChanged("Direction");
            Recalculate();
            UpdateVisual();
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

            Coordinates = Math.GetTranslationPoint(source.Coordinates, DistanceValue, DirectionRadians);
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
            var distanceSource = element.ReadString("DistanceSource");
            var directionSource = element.ReadString("DirectionSource");
            bool newFormat = distanceSource != null
                || directionSource != null
                || element.Attribute("DistanceSourceIndex") != null
                || element.Attribute("DirectionSourceIndex") != null
                || element.Attribute("FreeDistance") != null
                || element.Attribute("FreeDirection") != null;
            if (newFormat)
            {
                distanceQuantity.SourceIndex = IndexOfSource(element, "Distance", distanceSource);
                directionQuantity.SourceIndex = IndexOfSource(element, "Direction", directionSource);
                distanceQuantity.Parameter = element.ReadDouble("Distance");
                directionQuantity.Parameter = element.ReadDouble("Direction").ToRadians();
            }
            else
            {
                if (Dependencies.Count > 1 && Dependencies[1] is Vector)
                {
                    distanceQuantity.SourceIndex = 1;
                    directionQuantity.SourceIndex = 1;
                }
                else
                {
                    distanceQuantity.SourceIndex = Dependencies.Count > 1 && Dependencies[1] is ILengthProvider ? 1 : -1;
                    directionQuantity.SourceIndex = Dependencies.Count > 2 && Dependencies[2] is IAngleProvider ? 2 : -1;
                }

                distanceQuantity.Parameter = element.ReadDouble("Magnitude");
                directionQuantity.Parameter = element.ReadDouble("Direction");
                needsLegacyUpgrade = HasFreedom;
            }

            Recalculate();
        }

        // A source is named, except a part of a figure (a side of a regular polygon), which
        // has no name: that one is said by its place among the dependencies
        // (DistanceSourceIndex="1"). Written by its empty name, it was read as no source
        // at all, and the point came back free.
        int IndexOfSource(XElement element, string quantity, string name)
        {
            var place = element.Attribute(quantity + "SourceIndex");
            if (place != null)
            {
                int index = (int)element.ReadDouble(quantity + "SourceIndex");
                return index >= 1 && index < Dependencies.Count ? index : -1;
            }

            return IndexOfDependency(element, name);
        }

        void WriteSource(XmlWriter writer, string quantity, IFigure source)
        {
            if (string.IsNullOrEmpty(source.Name))
            {
                writer.WriteAttributeString(
                    quantity + "SourceIndex",
                    Dependencies.IndexOf(source).ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
            else
            {
                writer.WriteAttributeString(quantity + "Source", source.Name);
            }
        }

        /// <summary>
        /// The place of the named source among the figure's own Dependency elements: the
        /// name is the one the file says, which the figure read under it need not have any
        /// more. (A paste gives a copy whose name is taken another one: the copy of a
        /// fixed segment's Number n1 is n2, and the copied end, looking for n1 among
        /// figures that were now called n2, came back with no sources - free, at distance
        /// 0, on its pivot.)
        /// </summary>
        int IndexOfDependency(XElement element, string name)
        {
            if (name == null)
            {
                return -1;
            }

            int index = 0;
            foreach (var dependency in element.Elements("Dependency"))
            {
                if (dependency.ReadString("Name") == name && dependency.Attribute("Part") == null)
                {
                    return index < Dependencies.Count ? index : -1;
                }

                index++;
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
            if (IsDistanceFree)
            {
                Fix(distanceQuantity, DistanceValue);
            }

            if (IsDirectionFree)
            {
                Fix(directionQuantity, DirectionRadians);
            }
        }

        public override void WriteXml(XmlWriter writer)
        {
            base.WriteXml(writer);
            var distanceSource = DistanceSource;
            if (distanceSource != null)
            {
                WriteSource(writer, "Distance", distanceSource);
            }
            else
            {
                writer.WriteAttributeBool("FreeDistance", true);
                writer.WriteAttributeDouble("Distance", distanceQuantity.Parameter);
            }

            var directionSource = DirectionSource;
            if (directionSource != null)
            {
                WriteSource(writer, "Direction", directionSource);
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
