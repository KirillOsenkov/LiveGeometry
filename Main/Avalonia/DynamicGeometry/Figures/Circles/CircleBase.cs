using System.Collections.Generic;
using Avalonia;

namespace DynamicGeometry
{
    public abstract partial class CircleBase : EllipseBase, ICircle, IFixableLength
    {

        public abstract double Radius
        {
            get;
        }

        #region Setting and fixing the radius

        /// <summary>
        /// The two points whose distance is the radius, when the radius can be set: the center
        /// and the point on the circle (Circle), the two radius points (Circle by radius).
        /// Null when it can't (a circle by equation, a radius taken from a figure): the
        /// Radius row is then read-only and the verbs are hidden.
        /// </summary>
        protected virtual IPoint RadiusPivot
        {
            get { return null; }
        }

        protected virtual IPoint RadiusEnd
        {
            get { return null; }
        }

        /// <summary>The radius, settable through the two points (<see cref="LengthConstraint.SetDistance"/>)</summary>
        [PropertyGridVisible]
        [PropertyGridName("Radius")]
        [PropertyGridGroup("Radius")]
        [PropertyGridPreferredEditor("UpDown")]
        [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
        public double Length
        {
            get
            {
                return Radius;
            }
            set
            {
                if (RadiusEnd != null)
                {
                    LengthConstraint.SetDistance(RadiusEnd, RadiusPivot, value);
                }
            }
        }

        /// <summary>A measurement of the radius sits between the two radius points</summary>
        public IList<IFigure> MeasuredFigures
        {
            get { return RadiusEnd == null ? null : new IFigure[] { RadiusPivot, RadiusEnd }; }
        }

        public bool CanEdit(string propertyName)
        {
            var end = RadiusEnd;
            bool isFixed = end != null && LengthConstraint.FixedEnd(end, RadiusPivot) != null;
            bool canStretch = end != null && LengthConstraint.CanStretch(end, RadiusPivot);
            switch (propertyName)
            {
                case "Length":
                    return isFixed || canStretch;
                case "FixLength":
                    return !isFixed && canStretch;
                case "FreeLength":
                    return isFixed;
                default:
                    return true;
            }
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            switch (propertyName)
            {
                case "Length":
                    return "Radius";
                case "FixLength":
                    return "Fix radius";
                case "FreeLength":
                    return "Free radius";
                default:
                    return defaultCaption;
            }
        }

        /// <summary>The circle keeps its radius (<see cref="LengthConstraint.Fix"/>)</summary>
        [PropertyGridVisible]
        [PropertyGridName("Fix radius")]
        [PropertyGridGroup("Radius")]
        [PropertyGridIcon(PropertyGridIcon.Lock)]
        public void FixLength()
        {
            if (CanEdit("FixLength"))
            {
                LengthConstraint.Fix(RadiusEnd, RadiusPivot);
                Drawing.RaiseDisplayProperties(this);
            }
        }

        [PropertyGridVisible]
        [PropertyGridName("Free radius")]
        [PropertyGridGroup("Radius")]
        [PropertyGridIcon(PropertyGridIcon.Unlock)]
        public void FreeLength()
        {
            if (CanEdit("FreeLength"))
            {
                LengthConstraint.Free(LengthConstraint.FixedEnd(RadiusEnd, RadiusPivot));
                Drawing.RaiseDisplayProperties(this);
            }
        }

        #endregion

        public override double Inclination
        {
            get { return 0; }
        }

        public override double SemiMajor
        {
            get { return Radius; }
        }

        public override double SemiMinor
        {
            get { return Radius; }
        }

        public override void UpdateVisual()
        {
            var center = ToPhysical(Center);
            var diameter = ToPhysical(Radius * 2) + shape.StrokeThickness;
            if (shape.Width != diameter)
            {
                shape.Width = diameter;
            }

            if (shape.Height != diameter)
            {
                shape.Height = diameter;
            }

            shape.CenterAt(center);
        }

        public override Point GetPointFromParameter(double parameter)
        {
            if (Settings.PointsOnEllipticalsUseAbsoluteAngle)
            {
                var center = Center;
                var radius = Radius;
                return new Point(
                    center.X + radius * System.Math.Cos(parameter),
                    center.Y + radius * System.Math.Sin(parameter));
            }
            else
            {
                return base.GetPointFromParameter(parameter);
            }
        }
    }
}
