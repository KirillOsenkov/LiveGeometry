using System.Collections.Generic;
using System.Windows;
using System;
using System.Linq;
using DynamicGeometry;

namespace PolylineRouting
{
    public class RoutingAlgorithm
    {
        public static List<Point> GrahamScan
            (Point start, Point end, List<Point> polygon)
        {
            var algorithm = new RoutingAlgorithmGrahamScan(start, end, polygon);
            return Apply(algorithm);
        }

        public static List<Point> Dijkstra
            (Point start, Point end, List<Point> polygon)
        {
            var algorithm = new RoutingAlgorithmDijkstra(start, end, polygon);
            return Apply(algorithm);
        }

        public static List<Point> Apply(RoutingAlgorithm algorithm)
        {
            var result = algorithm.ShortestRoute();
            result = RemoveRedundantSegments(result);
            return result;
        }

        public RoutingAlgorithm()
        {
        }

        public RoutingAlgorithm(Point start, Point end, List<Point> polygon)
        {
            Start = start;
            End = end;
            Polygon = polygon;
        }

        Point _start;
        public Point Start
        {
            get
            {
                return _start;
            }
            set
            {
                _start = value.RoundToEpsilon();
            }
        }
        Point _end;
        public Point End
        {
            get
            {
                return _end;
            }
            set
            {
                _end = value.RoundToEpsilon();
            }
        }
        List<Point> _polygon;
        public List<Point> Polygon
        {
            get
            {
                return _polygon;
            }
            set
            {
                _polygon = value.RoundToEpsilon().ToList();
            }
        }
        public int N { get { return Polygon.Count; } }

        public virtual List<Point> ShortestRoute()
        {
            return new List<Point>() 
            { 
                Start, End 
            };
        }

        public static List<Point> RemoveRedundantSegments(List<Point> points)
        {
            points = new List<Point>(points);
            for (int i = 1; i < points.Count - 1; i++)
            {
                if (DynamicGeometry.Math.VectorProduct
                    (points[i - 1], points[i], points[i + 1]) == 0)
                {
                    points.RemoveAt(i--);
                }
            }
            return points;
        }

        public void ParseInput(string segment, string countLines, string polygonCoordinates)
        {
            string[] segmentLine = segment.Split(',');
            Start = new Point(Convert.ToDouble(segmentLine[0]), Convert.ToDouble(segmentLine[1]));
            End = new Point(Convert.ToDouble(segmentLine[2]), Convert.ToDouble(segmentLine[3]));

            int count = Convert.ToInt32(countLines);

            double[] polygon = polygonCoordinates.Split(',').Select(s => Convert.ToDouble(s)).ToArray();
            Polygon = new List<Point>();
            for (int i = 0; i * 2 + 1 < polygon.Length; i++)
            {
                Polygon.Add(new Point(polygon[i * 2], polygon[i * 2 + 1]));
            }
        }
    }
}