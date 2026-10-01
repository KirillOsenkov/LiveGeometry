using Avalonia;

namespace DynamicGeometry
{
    public interface IMovable
    {
        void MoveTo(Point position);
        bool AllowMove();
        Point Coordinates { get; }
    }

    /// <summary>
    /// A movable that can say where it is and be put back there exactly, for undo and redo
    /// of a move (<see cref="MoveAction"/>): its coordinates, or what its place really is - a
    /// label's offset in pixels from what it labels, the distance and direction of a
    /// translated point. Moving back by the offset it was moved by is not the same: a
    /// point on a circle doesn't land where the offset says.
    /// </summary>
    public interface IRestorablePlace
    {
        object CapturePlace();
        void RestorePlace(object place);
    }

    /// <summary>
    /// A figure dragged by parts: the part under the press is what moves. A slider's knob
    /// changes its value, anywhere else on the slider moves the whole of it.
    /// </summary>
    public interface IMovableParts
    {
        /// <returns>The part a press here drags; null when nothing on the figure moves</returns>
        IMovable FindMovablePart(Point point);
    }

    public static class IMovableExtensions
    {
        public static void MoveTo(this IMovable movable, double x, double y)
        {
            movable.MoveTo(new Point(x, y));
        }
    }
}
