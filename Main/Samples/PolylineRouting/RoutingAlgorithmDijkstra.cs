using System.Collections.Generic;
using System.Linq;
using System.Windows;
using DynamicGeometry;

namespace PolylineRouting
{
    public class RoutingAlgorithmDijkstra : RoutingAlgorithm
    {
        public RoutingAlgorithmDijkstra(Point start, Point end, List<Point> polygon)
            : base(start, end, polygon) { }

        public RoutingAlgorithmDijkstra() { }

        class Node
        {
            public Point Point;
            public List<int> Neighbors = new List<int>();
        }

        public override List<Point> ShortestRoute()
        {
            Graph graph = ConstructGraph();
            int n = graph.Nodes.Count;
            List<Node> shortestPath = graph.ShortestPath(n - 2, n - 1);
            return (from node in shortestPath
                    select node.Point).ToList();
        }

        class Graph
        {
            public List<Node> Nodes = new List<Node>();
            public double[] Distance;
            public int[] Previous;

            public List<Node> ShortestPath(int start, int end)
            {
                int n = Nodes.Count;
                Distance = new double[n];
                Previous = new int[n];
                for (int i = 0; i < n; i++)
                {
                    Distance[i] = double.PositiveInfinity;
                    Previous[i] = -1;
                }
                Distance[start] = 0;

                var queue = Enumerable.Range(0, n).ToList();
                while (queue.Count > 0)
                {
                    int min = FindMinimalDistance(queue);
                    int current = queue[min];
                    if (Distance[current] == double.PositiveInfinity)
                    {
                        break;
                    }
                    queue.RemoveAt(min);

                    foreach (int neighbor in Nodes[current].Neighbors)
                    {
                        double distanceToNeighbor =
                            Nodes[current].Point.Distance(Nodes[neighbor].Point);
                        double newDistance = Distance[current] + distanceToNeighbor;
                        if (Distance[neighbor] == double.PositiveInfinity
                            || newDistance <= Distance[neighbor])
                        {
                            Distance[neighbor] = newDistance;
                            Previous[neighbor] = current;
                        }
                    }
                }

                List<Node> result = new List<Node>();
                result.Add(Nodes[end]);
                while (Previous[end] != -1)
                {
                    end = Previous[end];
                    if (result.Contains(Nodes[end]))
                    {
                        break;
                    }
                    result.Add(Nodes[end]);
                }

                result.Reverse();
                return result;
            }

            int FindMinimalDistance(List<int> queue)
            {
                double minimum = double.PositiveInfinity;
                int minIndex = 0;
                for (int i = 0; i < queue.Count; i++)
                {
                    if (Distance[queue[i]] < minimum)
                    {
                        minimum = Distance[queue[i]];
                        minIndex = i;
                    }
                }
                return minIndex;
            }
        }

        Graph ConstructGraph()
        {
            Graph result = new Graph();
            foreach (var vertex in Polygon)
            {
                result.Nodes.Add(new Node() { Point = vertex });
            }
            result.Nodes.Add(new Node() { Point = Start });
            result.Nodes.Add(new Node() { Point = End });
            AddEdges(result);
            return result;
        }

        void AddEdges(Graph graph)
        {
            int n = graph.Nodes.Count;
            for (int i = 0; i < n - 1; i++)
            {
                for (int j = i + 1; j < n; j++)
                {
                    if ((j == i + 1 && j < n - 2) || (i == 0 && j == n - 3))
                    {
                        AddEdge(graph, i, j);
                    }
                    else
                    {
                        var a = graph.Nodes[i].Point;
                        var b = graph.Nodes[j].Point;
                        PointPair segment = new PointPair(a, b);
                        if (!SegmentIntersectsPolygon(segment))
                        {
                            AddEdge(graph, i, j);
                        }
                    }
                }
            }
        }

        void AddEdge(Graph graph, int i, int j)
        {
            graph.Nodes[i].Neighbors.Add(j);
            graph.Nodes[j].Neighbors.Add(i);
        }

        bool SegmentIntersectsPolygon(PointPair segment)
        {
            if (Math.GetIntersections(Polygon, segment).Count != 0)
            {
                return true;
            }

            if (Math.IsPointInPolygon(Polygon, segment.Midpoint))
            {
                return true;
            }

            for (int k = 0; k < Polygon.Count; k++)
            {
                var projection = Math.GetProjection(Polygon[k], segment);
                if (projection.Ratio > 0 
                    && projection.Ratio < 1
                    && projection.DistanceToLine < 0.01)
                {
                    var point1 = Math.Midpoint(segment.P1, projection.Point);
                    var point2 = Math.Midpoint(segment.P2, projection.Point);
                    if (Math.IsPointInPolygon(Polygon, point1)
                        || Math.IsPointInPolygon(Polygon, point2))
                    {
                        return true;
                    }
                }
            }

            return false;
        }
    }
}