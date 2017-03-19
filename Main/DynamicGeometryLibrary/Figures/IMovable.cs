using System.Linq;
using System.Windows;

namespace DynamicGeometry
{
    public interface IMovable
    {
        void MoveTo(Point position);
        bool AllowMove();
        Point Coordinates { get; }
    }

    public static class IMovableExtensions
    {
        public static void MoveTo(this IMovable movable, double x, double y)
        {
            movable.MoveTo(new Point(x, y));
        }
    }
}
