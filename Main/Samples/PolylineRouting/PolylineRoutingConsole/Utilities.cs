using System.Collections.Generic;

namespace DynamicGeometry
{
    public static class Extensions
    {
        public static PointPair GetSegment(this IList<Point> points, int startIndex)
        {
            return new PointPair(points[startIndex], points[(startIndex + 1) % points.Count]);
        }

        public static int RotateNext(this int index, int count)
        {
            return (index + 1) % count;
        }

        public static int RotatePrevious(this int index, int count)
        {
            return index > 0 ? index - 1 : count - 1;
        }

        public static void RemoveLast<T>(this IList<T> list)
        {
            list.RemoveAt(list.Count - 1);
        }
    }
}
