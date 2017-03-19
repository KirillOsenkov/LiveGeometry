using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using DynamicGeometry;

namespace PolylineRouting
{
    public class Route : Curve
    {
        public override void GetPoints(List<Point> list)
        {
            var start = Point(0);
            var end = Point(1);
            Polygon polygon = (Polygon)Dependencies.ElementAt(2);
            var points = polygon.Dependencies.ToPoints();

            var polyline = RoutingAlgorithm.Dijkstra(start, end, new List<Point>(points));
            list.AddRange(polyline);
        }

        protected override void ConstructPolyline(List<Point> points)
        {
            PolylineRounding(points, pathSegments);
        }
    }
}