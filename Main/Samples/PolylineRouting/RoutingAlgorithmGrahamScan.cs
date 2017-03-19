using System.Collections.Generic;
using System.Linq;
using System.Windows;
using DynamicGeometry;

namespace PolylineRouting
{
    public class RoutingAlgorithmGrahamScan : RoutingAlgorithm
    {
        public RoutingAlgorithmGrahamScan(
            Point start, Point end, List<Point> polygon)
            : base(start, end, polygon) { }

        public RoutingAlgorithmGrahamScan() { }

        int nearestIntersection;
        int farthestIntersection;
        int intersectionsCount;
        List<IntersectionData> intersectionData = new List<IntersectionData>();

        public override List<Point> ShortestRoute()
        {
            var result = new List<Point>();

            FindIntersections();
            if (intersectionsCount == 0 || intersectionsCount % 2 == 1)
            {
                result.Add(Start);
                result.Add(End);
                return result;
            }
            FindNearestIntersection();

            var route1 = ShortestRouteOriented(false);
            var route2 = ShortestRouteOriented(true);
            var length1 = route1.PolylineLength();
            var length2 = route2.PolylineLength();
            result = length1 < length2 ? route1 : route2;

            return result;
        }

        void FindNearestIntersection()
        {
            nearestIntersection = intersectionData[0].Index;
            farthestIntersection = nearestIntersection;
            int other = nearestIntersection;
            for (int i = 0; i < N; i++)
            {
                farthestIntersection = farthestIntersection.RotateNext(N);
                other = other.RotatePrevious(N);
                if (intersectionData.FirstOrDefault(d => d.Index == farthestIntersection) != null)
                {
                    break;
                }
                if (intersectionData.FirstOrDefault(d => d.Index == other) != null)
                {
                    farthestIntersection = other;
                    break;
                }
            }
        }

        List<Point> ShortestRouteOriented(bool clockwise)
        {
            var result = new List<Point>();
            var path = PreparePath(clockwise);
            int sign = 1;
            if (path.Count > 2)
            {
                sign = Math.VectorProduct(path[0], path[1], End).Sign();
                result = GrahamScan(path, sign);
                result = CutOutLoops(result);
            }

            if (result.Count == 2
                && ((result[0] == Start && result[1] == End)
                    || (result[0] == End && result[1] == Start)))
            {
                return result;
            }

            result = ExpandRouteRecursive(result);
            if (result.Count > 2)
            {
                result = GrahamScan(result, sign);
            }
            result = ExpandRouteRecursive(result);

            return result;
        }

        int SegmentIntersectsSide(Point a, Point b)
        {
            var segment = new PointPair(a, b);
            for (int i = 0; i < N; i++)
            {
                Point intersection = Math.GetIntersectionOfSegments(segment, Polygon.GetSegment(i));
                if (intersection.Exists()
                    && Polygon[i] != Start
                    && Polygon[i] != End
                    && Polygon[i.RotateNext(N)] != Start
                    && Polygon[i.RotateNext(N)] != End)
                {
                    return i;
                }
            }
            return -1;
        }

        void FindIntersections()
        {
            intersectionsCount = 0;
            int currentSign = 1;
            intersectionData.Clear();
            var segment = new PointPair(Start, End);

            for (int i = 0; i < N; i++)
            {
                var current = Polygon.GetSegment(i);
                Point intersection = Math.GetIntersectionOfSegments(
                    segment, current);
                if (intersection.Exists()
                    && Polygon[i] != Start
                    && Polygon[i] != End
                    && Polygon[i.RotateNext(N)] != Start
                    && Polygon[i.RotateNext(N)] != End)
                {
                    intersectionsCount++;
                    int orientation = Math.VectorProduct(Start, Polygon[i.RotateNext(N)], End).Sign();

                    intersectionData.Add(new IntersectionData()
                    {
                        Index = i,
                        IntersectionPoint = intersection,
                        DistanceToStart = intersection.Distance(Start),
                        Orientation = orientation
                    });

                    currentSign = -currentSign;
                }
            }

            intersectionData = new List<IntersectionData>(
                intersectionData.OrderBy(d => d.DistanceToStart));
        }

        class IntersectionData
        {
            public int Index { get; set; }
            public double DistanceToStart { get; set; }
            public Point IntersectionPoint { get; set; }
            public int Orientation { get; set; }

            public override string ToString()
            {
                return Index.ToString()
                    + ": "
                    + IntersectionPoint.ToString()
                    + ", Distance = "
                    + DistanceToStart.ToString()
                    + ", Orientation = "
                    + Orientation.ToString();
            }
        }

        List<Point> PreparePath(bool clockwise)
        {
            List<Point> path = new List<Point>();
            System.Func<int, int> next = i => i.RotateNext(N);
            int j = nearestIntersection;
            int farthest = farthestIntersection;
            if (clockwise)
            {
                next = i => i.RotatePrevious(N);
            }
            else
            {
                farthest = next(farthest);
                j = next(j);
            }

            path.Add(Start);
            while (true)
            {
                var current = Polygon[j];
                if (current != Start && current != End)
                {
                    path.Add(current);
                }
                j = next(j);
                if (j == farthest)
                {
                    break;
                }
            }
            path.Add(End);
            return path;
        }

        static List<Point> GrahamScan(List<Point> path, int sign)
        {
            path = new List<Point>(path);
            if (path.Count < 3)
            {
                return path;
            }
            List<Point> result = new List<Point>();

            result.Add(path[0]);
            result.Add(path[1]);

            for (int i = 2; i <= path.Count - 1; i++)
            {
                while (result.Count >= 2
                    && Math.VectorProduct(
                        result[result.Count - 2],
                        result[result.Count - 1],
                        path[i]).Sign() != sign)
                {
                    result.RemoveLast();
                }
                result.Add(path[i]);
            }

            return result;
        }

        List<Point> CutOutLoops(List<Point> path)
        {
            List<Point> result = new List<Point>(path);
            for (int i = 3; i < path.Count; i++)
            {
                for (int j = 1; j < i - 1; j++)
                {
                    var currentSegment = new PointPair(path[i - 1], path[i]);
                    var previousSegment = new PointPair(path[j - 1], path[j]);
                    if (Math.GetIntersectionOfSegments
                        (currentSegment, previousSegment)
                        .Exists())
                    {
                        if (SegmentIntersectsSide(path[i - 1], path[j - 1]) == -1)
                        {
                            result.RemoveRange(j, i - j);
                            return CutOutLoops(result);
                        }
                    }
                }
            }
            return result;
        }

        List<Point> ExpandRouteRecursive(List<Point> route)
        {
            route = new List<Point>(route);
            List<Point> result = new List<Point>();
            result.Add(route[0]);

            for (int i = 0; i < route.Count - 1; i++)
            {
                var routeForSegment = RoutingAlgorithm.GrahamScan(
                    route[i],
                    route[i + 1],
                    Polygon);
                if (routeForSegment.Count == 2)
                {
                    result.Add(route[i + 1]);
                }
                else if (routeForSegment.Count > 2)
                {
                    for (int j = 1; j < routeForSegment.Count; j++)
                    {
                        result.Add(routeForSegment[j]);
                    }
                }
                else
                {
                    result.Add(route[i + 1]);
                }
            }
            return result;
        }
    }
}