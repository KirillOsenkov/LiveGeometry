using System.Collections.Generic;

namespace DynamicGeometry
{
    public class LineTwoPoints : LineBase, ILine, IConditionalProperties
    {
        /// <summary>
        /// Whether the line runs through its two points, so that a ray or a segment can be
        /// drawn on them instead. Not a line built otherwise - a parallel, a perpendicular,
        /// a bisector: on those the Convert buttons made a ray out of a line and a point
        /// (and threw), or the segment whose bisector it was.
        /// </summary>
        protected virtual bool IsThroughTwoPoints
        {
            get
            {
                return true;
            }
        }

        public virtual bool CanEdit(string propertyName)
        {
            if (propertyName == nameof(ConvertToRay) || propertyName == nameof(ConvertToSegment))
            {
                return IsThroughTwoPoints;
            }

            return true;
        }

        public virtual string Caption(string propertyName, string defaultCaption)
        {
            return defaultCaption;
        }

        public override PointPair OnScreenCoordinates
        {
            get
            {
                return Math.GetLineFromSegment(Coordinates, CanvasLogicalBorders);
            }
        }

        protected override IReadOnlyList<string> NamesFromDependencies()
        {
            return NamesFromPoints(PointOrder.Reversible);
        }

        protected override string Kind
        {
            get
            {
                return "Line";
            }
        }

        public static void Convert(ILine oldLine, ILine newLine)
        {
            var drawing = oldLine.Drawing;
            newLine.Style = oldLine.Style;

            // a hidden helper (converted from the Figure List) stays hidden
            newLine.Visible = oldLine.Visible;
            newLine.Locked = oldLine.Locked;
            Actions.ReplaceWithNew(oldLine, newLine);
            drawing.RaiseUserIsAddingFigures(new Drawing.UIAFEventArgs() { Figures = newLine.AsEnumerable<IFigure>() });
        }

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert to ray")]
        [PropertyGridIcon(PropertyGridIcon.Ray)]
        public void ConvertToRay()
        {
            LineTwoPoints.Convert(this, Factory.CreateRay(this.Drawing, this.Dependencies));
        }

        [PropertyGridVisible]
        [PropertyGridName("Convert to segment")]
        [PropertyGridIcon(PropertyGridIcon.Segment)]
        public void ConvertToSegment()
        {
            LineTwoPoints.Convert(this, Factory.CreateSegment(this.Drawing, this.Dependencies));
        }

#endif

    }
}