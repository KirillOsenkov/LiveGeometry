using Avalonia;

namespace DynamicGeometry
{
    public class MidPoint : PointBase, IPoint
    {
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
    }
}