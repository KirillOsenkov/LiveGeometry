using System.Collections.Generic;
using System;
using System.Linq;
using DynamicGeometry;

namespace PolylineRouting
{
    public abstract class RoutingAlgorithm
    {
        public static List<Point> GrahamScan
            (Point start, Point end, List<Point> polygon)
        {
            var algorithm = new RoutingAlgorithmGrahamScan(start, end, polygon);
            var result = algorithm.ShortestRoute();
            return result;
        }

        public static List<Point> Dijkstra
            (Point start, Point end, List<Point> polygon)
        {
            var algorithm = new RoutingAlgorithmDijkstra(start, end, polygon);
            var result = algorithm.ShortestRoute();
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

        protected Point Start;
        protected Point End;
        protected List<Point> Polygon;
        protected int N { get { return Polygon.Count; } }

        public abstract List<Point> ShortestRoute();

        public void ParseInput(string segment, string countLines, string polygonCoordinates)
        {
            string[] segmentLine = segment.Split(',');
            Start = new Point(Convert.ToDouble(segmentLine[0]), Convert.ToDouble(segmentLine[1]));
            End = new Point(Convert.ToDouble(segmentLine[2]), Convert.ToDouble(segmentLine[3]));

            int count = Convert.ToInt32(countLines);

            double[] polygon = polygonCoordinates.Split(',').Select(s => Convert.ToDouble(s)).ToArray();
            Polygon = new List<Point>();
            for (int i = 0; i < count; i++)
            {
                Polygon.Add(new Point(polygon[i * 2], polygon[i * 2 + 1]));
            }
        }
   }
}