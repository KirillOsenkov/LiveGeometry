using System.Linq;
using Avalonia;

namespace DynamicGeometry
{
    public class MidPoint : PointBase, IPoint, IConditionalProperties
    {
        /// <summary>No "Free point" while a locus is drawn from the point</summary>
        public bool CanEdit(string propertyName)
        {
            return propertyName != nameof(Release) || PointSnapping.CanRelease(this);
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            return defaultCaption;
        }

        protected override string Kind
        {
            get
            {
                return "Midpoint";
            }
        }

        /// <summary>"of CD"</summary>
        public override string Construction
        {
            get
            {
                return "of " + ConstructionText.Points(Dependencies.ToArray());
            }
        }

        protected override Avalonia.Controls.Shapes.Shape CreateShape()
        {
            return Factory.CreateDependentPointShape();
        }

        public override void Recalculate()
        {
            Coordinates = new Point(
                (Point(0).X + Point(1).X) / 2,
                (Point(0).Y + Point(1).Y) / 2);
        }

        /// <summary>Lets go of the two points: a free point where it is (<see cref="PointSnapping.Release"/>)</summary>
        [PropertyGridVisible]
        [PropertyGridName("Free point")]
        [PropertyGridIcon(PropertyGridIcon.Unlock)]
        public void Release()
        {
            PointSnapping.Release(this);
        }
    }
}