using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Linq;
using Avalonia.Controls;

namespace DynamicGeometry
{
    public class Segment : LineBase, ILengthProvider, ILine, IFixableLength
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
            Mark.Show(ToPhysical(coordinates.P1), ToPhysical(coordinates.P2), Shape.Stroke, Shape.StrokeThickness, decoration);
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
        [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
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
                Drawing.RaiseDisplayProperties(this);
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
