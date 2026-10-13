using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Linq;
using Avalonia.Controls;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    public class Segment : LineBase, ILengthProvider, ILine, IFixableLength, IFixableDirection
    {
        #region Decoration

        SegmentDecoration decoration;
        SegmentDecorationMark decorationMark;

        /// <summary>The mark at the middle: ticks, chevrons or a wave (<see cref="SegmentDecorationMark"/>)</summary>
        [PropertyGridVisible]
        [PropertyGridName("Mark")]
        [PropertyGridPreferredEditor("SegmentDecoration")]
        public SegmentDecoration Decoration
        {
            get
            {
                return decoration;
            }
            set
            {
                decoration = value;
                if (Drawing != null)
                {
                    UpdateVisual();
                }

                RaisePropertyChanged("Decoration");
            }
        }

        // not a field initializer: virtual members get called from the base constructor
        SegmentDecorationMark Mark
        {
            get
            {
                if (decorationMark == null)
                {
                    decorationMark = new SegmentDecorationMark();
                }

                return decorationMark;
            }
        }

        public override void UpdateVisual()
        {
            base.UpdateVisual();
            if (decoration == SegmentDecoration.None || !IsShown)
            {
                decorationMark?.Hide();
                return;
            }

            // in the segment's own stroke, as it is drawn now (thicker when selected)
            var coordinates = OnScreenCoordinates;
            Mark.Show(
                ToPhysical(coordinates.P1),
                ToPhysical(coordinates.P2),
                Shape.Stroke,
                Shape.StrokeThickness,
                ZIndex,
                decoration);
        }

        public override void OnAddingToCanvas(Canvas newContainer)
        {
            base.OnAddingToCanvas(newContainer);
            Mark.OnAddingToCanvas(newContainer);
        }

        public override void OnRemovingFromCanvas(Canvas leavingContainer)
        {
            base.OnRemovingFromCanvas(leavingContainer);
            Mark.OnRemovingFromCanvas(leavingContainer);
        }

        public override void ReadXml(XElement element)
        {
            base.ReadXml(element);
            var name = element.ReadString("Decoration");
            if (name != null && Enum.TryParse(name, out SegmentDecoration read))
            {
                decoration = read;
            }
        }

        public override void WriteXml(XmlWriter writer)
        {
            base.WriteXml(writer);
            if (decoration != SegmentDecoration.None)
            {
                writer.WriteAttributeString("Decoration", decoration.ToString());
            }
        }

        #endregion

        /// <summary>
        /// Setting it on a segment with a fixed length (<see cref="FixLength"/>) changes that
        /// length. Otherwise it stretches the segment once, to that length, by moving an end
        /// away from the other: the second end if it can take it, else the first - not a
        /// constraint, dragging changes the length again. Read-only when neither applies.
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridGroup("Length")]
        [PropertyGridPreferredEditor("UpDown")]
        [PropertyGridCustomValueProvider(typeof(LengthPropertyValue))]
        public double Length
        {
            get
            {
                return Coordinates.Length;
            }
            set
            {
                int end = LengthEnd();
                if (end >= 0)
                {
                    LengthConstraint.SetDistance(End(end), End(1 - end), value);
                }
            }
        }

        public IPoint LengthEndpoint
        {
            get
            {
                int end = LengthEnd();
                return end >= 0 ? End(end) : null;
            }
        }

        /// <summary>
        /// The way from the first point to the second, in degrees counterclockwise from the x
        /// axis. Setting it on a segment with a fixed direction (<see cref="FixDirection"/>)
        /// changes that direction. Otherwise it turns an end about the other once: the second
        /// end if it can take it, else the first - not a constraint, dragging changes the
        /// direction again. Read-only when neither applies.
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridName("Direction")]
        [PropertyGridGroup("Direction")]
        [PropertyGridPreferredEditor("UpDown")]
        [PropertyGridCustomValueProvider(typeof(DirectionPropertyValue))]
        public override double Angle
        {
            get
            {
                // a full turn but for rounding is no turn: a first end fixed at 180° from the
                // second, a hair below it (sin 180° is no 0 in doubles), said 360
                double angle = base.Angle;
                return 360 - angle < 1e-9 ? 0 : angle;
            }
            set
            {
                int end = DirectionEnd();
                if (end >= 0)
                {
                    // the first end is turned about the second, the opposite way
                    DirectionConstraint.SetDirection(End(end), End(1 - end), end == 1 ? value : value + 180);
                }
            }
        }

        public IPoint DirectionEndpoint
        {
            get
            {
                int end = DirectionEnd();
                return end >= 0 ? End(end) : null;
            }
        }

        IPoint End(int index)
        {
            return (IPoint)Dependencies[index];
        }

        public IList<IFigure> MeasuredFigures
        {
            get { return new IFigure[] { this }; }
        }

        /// <summary>The end holding a fixed length from the other, or null (see <see cref="LengthConstraint.FixedEnd"/>)</summary>
        TranslatedPoint FixedEnd()
        {
            int end = FixedEndIndex();
            return end >= 0 ? (TranslatedPoint)Dependencies[end] : null;
        }

        int FixedEndIndex()
        {
            return IndexOfEnd((end, pivot) => LengthConstraint.FixedEnd(end, pivot) != null);
        }

        /// <summary>The end a new length moves (see <see cref="LengthConstraint.CanStretch"/>), the second preferred; -1 when neither can</summary>
        int EndToStretch()
        {
            return IndexOfEnd(LengthConstraint.CanStretch);
        }

        /// <summary>The end that answers for the length: the fixed one, else the one that can be stretched</summary>
        int LengthEnd()
        {
            int fixedEnd = FixedEndIndex();
            return fixedEnd >= 0 ? fixedEnd : EndToStretch();
        }

        /// <summary>The end holding a fixed direction from the other (see <see cref="DirectionConstraint.FixedEnd"/>); -1 when none does</summary>
        int FixedDirectionEndIndex()
        {
            return IndexOfEnd((end, pivot) => DirectionConstraint.FixedEnd(end, pivot) != null);
        }

        /// <summary>The end a new direction turns (see <see cref="DirectionConstraint.CanTurn"/>), the second preferred; -1 when neither can</summary>
        int EndToTurn()
        {
            return IndexOfEnd(DirectionConstraint.CanTurn);
        }

        /// <summary>The end that answers for the direction: the fixed one, else the one that can be turned</summary>
        int DirectionEnd()
        {
            int fixedEnd = FixedDirectionEndIndex();
            return fixedEnd >= 0 ? fixedEnd : EndToTurn();
        }

        bool CanFreeDirection()
        {
            int end = FixedDirectionEndIndex();
            return end >= 0 && DirectionConstraint.CanFree((TranslatedPoint)Dependencies[end]);
        }

        /// <summary>Whether Horizontal (axis 0) or Vertical (axis 90) would change something: not on a direction fixed along that axis already</summary>
        bool CanAlign(double axis)
        {
            if (DirectionEnd() < 0)
            {
                return false;
            }

            int fixedEnd = FixedDirectionEndIndex();
            return fixedEnd < 0 || !DirectionConstraint.IsAlong((TranslatedPoint)Dependencies[fixedEnd], axis);
        }

        int IndexOfEnd(System.Func<IFigure, IFigure, bool> qualifies)
        {
            if (Dependencies == null || Dependencies.Count < 2)
            {
                return -1;
            }

            for (int end = 1; end >= 0; end--)
            {
                if (qualifies(Dependencies[end], Dependencies[1 - end]))
                {
                    return end;
                }
            }

            return -1;
        }

        public bool CanEdit(string propertyName)
        {
            switch (propertyName)
            {
                case "Length":
                    return LengthEnd() >= 0;
                case "FixLength":
                    return FixedEndIndex() < 0 && EndToStretch() >= 0;
                case "FreeLength":
                    return FixedEndIndex() >= 0;
                case "Angle":
                    return DirectionEnd() >= 0;
                case "FixDirection":
                    return FixedDirectionEndIndex() < 0 && EndToTurn() >= 0;
                case "FreeDirection":
                    return CanFreeDirection();
                case "MakeHorizontal":
                    return CanAlign(axis: 0);
                case "MakeVertical":
                    return CanAlign(axis: 90);
                // a line or a ray has no length for what measures this one, nor for an
                // expression that reads it ([AB.Length])
                case "ConvertToLine":
                    return !this.IsUsedForLength() && !this.IsReadByExpressionsBeyond(typeof(LineTwoPoints));
                case "ConvertToRay":
                    return !this.IsUsedForLength() && !this.IsReadByExpressionsBeyond(typeof(Ray));
                default:
                    return true;
            }
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            return defaultCaption;
        }

#if !PLAYER && !TABULA

        /// <summary>The segment keeps its length from now on (<see cref="LengthConstraint.Fix"/>)</summary>
        [PropertyGridVisible]
        [PropertyGridName("Fix length")]
        [PropertyGridGroup("Length")]
        [PropertyGridIcon(PropertyGridIcon.Lock)]
        public void FixLength()
        {
            int end = EndToStretch();
            if (end >= 0)
            {
                LengthConstraint.Fix(End(end), End(1 - end));
                ShowProperties();
            }
        }

        /// <summary>The opposite of <see cref="FixLength"/> (<see cref="LengthConstraint.Free"/>)</summary>
        [PropertyGridVisible]
        [PropertyGridName("Free length")]
        [PropertyGridGroup("Length")]
        [PropertyGridIcon(PropertyGridIcon.Unlock)]
        public void FreeLength()
        {
            var fixedEnd = FixedEnd();
            if (fixedEnd != null)
            {
                LengthConstraint.Free(fixedEnd);
                ShowProperties();
            }
        }

        /// <summary>The segment keeps its direction from now on (<see cref="DirectionConstraint.Fix"/>)</summary>
        [PropertyGridVisible]
        [PropertyGridName("Fix direction")]
        [PropertyGridGroup("Direction")]
        [PropertyGridIcon(PropertyGridIcon.Lock)]
        public void FixDirection()
        {
            int end = EndToTurn();
            if (FixedDirectionEndIndex() < 0 && end >= 0)
            {
                DirectionConstraint.Fix(End(end), End(1 - end));
                ShowProperties();
            }
        }

        /// <summary>The opposite of <see cref="FixDirection"/> (<see cref="DirectionConstraint.Free"/>)</summary>
        [PropertyGridVisible]
        [PropertyGridName("Free direction")]
        [PropertyGridGroup("Direction")]
        [PropertyGridIcon(PropertyGridIcon.Unlock)]
        public void FreeDirection()
        {
            if (CanFreeDirection())
            {
                DirectionConstraint.Free((TranslatedPoint)Dependencies[FixedDirectionEndIndex()]);
                ShowProperties();
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Horizontal")]
        [PropertyGridGroup("Direction")]
        [PropertyGridIcon(PropertyGridIcon.Horizontal)]
        public void MakeHorizontal()
        {
            Align(axis: 0);
        }

        [PropertyGridVisible]
        [PropertyGridName("Vertical")]
        [PropertyGridGroup("Direction")]
        [PropertyGridIcon(PropertyGridIcon.Vertical)]
        public void MakeVertical()
        {
            Align(axis: 90);
        }

        /// <summary>
        /// The direction fixed along the axis, the nearer way (rightward or leftward, up or
        /// down: the second point stays on its side of the first). One undo step: the fix,
        /// then the direction set through the row, so that its undo puts the Number back.
        /// </summary>
        void Align(double axis)
        {
            if (!CanAlign(axis))
            {
                return;
            }

            double direction = DirectionConstraint.NearestAlong(Angle, axis);
            var actionManager = Drawing.ActionManager;
            using (Transaction.Create(actionManager, false))
            {
                if (FixedDirectionEndIndex() < 0)
                {
                    int end = EndToTurn();
                    DirectionConstraint.Fix(End(end), End(1 - end));
                }

                Actions.SetProperty(actionManager, PropertyDiscoveryStrategy.CreateValueProvider(this, nameof(Angle)), direction);
            }

            ShowProperties();
        }

        /// <summary>
        /// The grid again, with the other verb. Not for the segment inside a vector, which
        /// is no figure of the drawing: the vector shows its own (the segment's has a style
        /// from no style list, and the style picker threw on it).
        /// </summary>
        void ShowProperties()
        {
            if (Drawing.Figures.Contains(this))
            {
                Drawing.RaiseDisplayProperties(this);
            }
        }

#endif

        public override double GetNearestParameterFromPoint(Avalonia.Point point)
        {
            var parameter = base.GetNearestParameterFromPoint(point);
            if (parameter < 0)
            {
                parameter = 0;
            }
            else if (parameter > 1)
            {
                parameter = 1;
            }
            return parameter;
        }

        public override IFigure HitTest(Avalonia.Point point)
        {
            var epsilon = ToLogical(this.Shape.StrokeThickness) / 2 + CursorTolerance;
            if (Math.IsPointOnSegment(Coordinates, point, epsilon))
            {
                return this;
            }
            return null;
        }

        public override Tuple<double, double> GetParameterDomain()
        {
            return Tuple.Create(0.0, 1.0);
        }

        protected override IReadOnlyList<string> NamesFromDependencies()
        {
            return NamesFromPoints(PointOrder.Reversible);
        }

        protected override string Kind
        {
            get
            {
                return "Segment";
            }
        }

        public override string Noun
        {
            get
            {
                return "segment";
            }
        }

        public override string ToString()
        {
            // I think it is confusing to the user when the title of the property grid for a segment is different than the name.
            // For example, the user might use the ToString() instead of the Name when referring to the segment in an expression. - D.H.
            return base.ToString();
            //return "Segment " + Dependencies[0].ToString() + Dependencies[1].ToString();
        }

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert to line")]
        [PropertyGridIcon(PropertyGridIcon.Line)]
        public void ConvertToLine()
        {
            LineTwoPoints.Convert(this, Factory.CreateLineTwoPoints(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert to ray")]
        [PropertyGridIcon(PropertyGridIcon.Ray)]
        public void ConvertToRay()
        {
            LineTwoPoints.Convert(this, Factory.CreateRay(this.Drawing, this.Dependencies));
        }

#endif
    }
}
