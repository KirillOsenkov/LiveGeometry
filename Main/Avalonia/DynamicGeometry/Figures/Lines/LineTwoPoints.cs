using System.Collections.Generic;

namespace DynamicGeometry
{
    public class LineTwoPoints : LineBase, ILine, IFixableDirection
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

            // (a segment bisector is built on two points too, and points across them)
            if (DirectionConstraint.IsMember(propertyName))
            {
                return IsThroughTwoPoints && DirectionConstraint.CanEdit(this, propertyName);
            }

            return true;
        }

        /// <summary>
        /// The way from the first point to the second, in degrees: typed, fixed and freed as a
        /// segment's (<see cref="Segment.Direction"/>, <see cref="DirectionConstraint"/>). Not
        /// a row of a line built otherwise, which hides it.
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridGroup("Direction")]
        [PropertyGridPreferredEditor("UpDown")]
        [PropertyGridCustomValueProvider(typeof(DirectionPropertyValue))]
        public virtual double Direction
        {
            get
            {
                return DirectionConstraint.Of(Coordinates);
            }
            set
            {
                if (IsThroughTwoPoints)
                {
                    DirectionConstraint.SetDirection(this, value);
                }
            }
        }

        public IPoint DirectionEndpoint => IsThroughTwoPoints ? DirectionConstraint.Endpoint(this) : null;

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

        [PropertyGridVisible]
        [PropertyGridName("Fix direction")]
        [PropertyGridGroup("Direction")]
        [PropertyGridIcon(PropertyGridIcon.Lock)]
        public void FixDirection()
        {
            if (IsThroughTwoPoints)
            {
                DirectionConstraint.FixDirection(this);
                Drawing.RaiseDisplayProperties(this);
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Free direction")]
        [PropertyGridGroup("Direction")]
        [PropertyGridIcon(PropertyGridIcon.Unlock)]
        public void FreeDirection()
        {
            if (IsThroughTwoPoints)
            {
                DirectionConstraint.FreeDirection(this);
                Drawing.RaiseDisplayProperties(this);
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Horizontal")]
        [PropertyGridGroup("Direction")]
        [PropertyGridIcon(PropertyGridIcon.Horizontal)]
        public void MakeHorizontal()
        {
            if (IsThroughTwoPoints)
            {
                DirectionConstraint.Align(this, axis: 0);
                Drawing.RaiseDisplayProperties(this);
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Vertical")]
        [PropertyGridGroup("Direction")]
        [PropertyGridIcon(PropertyGridIcon.Vertical)]
        public void MakeVertical()
        {
            if (IsThroughTwoPoints)
            {
                DirectionConstraint.Align(this, axis: 90);
                Drawing.RaiseDisplayProperties(this);
            }
        }

#endif

    }
}