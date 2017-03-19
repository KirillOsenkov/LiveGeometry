using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace DynamicGeometry
{
    public abstract class CoordinatesShapeBase<TShape> : ShapeBase<TShape>, IMovable
        where TShape : FrameworkElement
    {
        public override void MoveToCore(Point newLocation)
        {
            Coordinates = newLocation;
        }

        public override void UpdateVisual()
        {
            if (!Visible || !Exists)
            {
                return;
            }

            shape.CenterAt(ToPhysical(Coordinates));
        }

        public Point Coordinates { get; set; }

        public override Point Center
        {
            get
            {
                return Coordinates;
            }
        }
    }
}
