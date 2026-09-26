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
