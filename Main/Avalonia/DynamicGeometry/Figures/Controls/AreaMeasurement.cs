using Avalonia;
using System.Linq;

namespace DynamicGeometry
{
    public class AreaMeasurement : Measurement
    {
        protected override string Kind
        {
            get
            {
                return "Area";
            }
        }

        /// <summary>"of triangle ABC", "of circle c", "of ABCD" (points)</summary>
        public override string Construction
        {
            get
            {
                if (Dependencies.Count == 0)
                {
                    return null;
                }

                return "of " + (Dependencies[0] is IShapeWithInterior
                    ? ConstructionText.Of(Dependencies[0])
                    : ConstructionText.Points(Dependencies.ToArray()));
            }
        }

        public double Measure
        {
            get
            {
                if (Dependencies[0] is IShapeWithInterior)
                {
                    return (Dependencies[0] as IShapeWithInterior).Area;
                }
                return Dependencies.ToPoints().Area();
            }
        }

        public override Point Anchor
        {
            get
            {
                return Origin;
            }
        }

        public override void UpdateVisual()
        {
            base.UpdateVisual();
            Text = Math.Round(Measure, DecimalsToShow).ToString();
        }

        private Point Origin
        {
            get
            {
                if (Dependencies[0] is PointBase)
                {
                    return Dependencies.ToPoints().Midpoint();
                }
                else
                {
                    return Dependencies[0].Center;
                }               
            }
        }
    }
}
