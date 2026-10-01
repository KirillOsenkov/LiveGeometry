using Avalonia;

namespace DynamicGeometry
{
    public abstract class CoordinatesShapeBase<TShape> : ShapeBase<TShape>, IMovable, IRestorablePlace
        where TShape : FrameworkElement
    {
        public override void MoveToCore(Point newLocation)
        {
            Coordinates = newLocation;
        }

        /// <summary>Where the figure is, for undo of a move: its coordinates, unless a subclass knows better</summary>
        public virtual object CapturePlace()
        {
            return Coordinates;
        }

        public virtual void RestorePlace(object place)
        {
            MoveTo((Point)place);
        }

        public override void UpdateVisual()
        {
            if (!IsShown)
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
