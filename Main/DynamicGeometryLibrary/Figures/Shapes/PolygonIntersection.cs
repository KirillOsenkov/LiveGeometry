using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace DynamicGeometry
{
    public class PolygonIntersection : CompositeFigure
    {
        public override void Recalculate()
        {
            var first = this.Dependencies.Polygon(0);
            var second = this.Dependencies.Polygon(1);

            var intersections = Intersect(first, second);

            ClearChildren();

            foreach (var intersection in intersections)
            {
                var vertices = new List<PointBase>(intersection.Length);

                foreach (var vertex in intersection)
                {
                    var point = new PointBase();
                    point.MoveToCore(vertex);
                    point.Dependencies.Add(this);
                    vertices.Add(point);
                    AddChild(point);
                }

                for (int i = 0; i < intersection.Length; i++)
                {
                    var side = new Segment();
                    if (i == 0)
                    {
                        side.Dependencies.Add(vertices[intersection.Length - 1], vertices[0]);
                    }
                    else
                    {
                        side.Dependencies.Add(vertices[i - 1], vertices[i]);
                    }

                    AddChild(side);
                }

                var polygon = new Polygon();
                polygon.Dependencies.AddRange(vertices);
                AddChild(polygon);
            }

            UpdateVisual();
        }

        public Point[][] Intersect(Point[] first, Point[] second)
        {
            var list = new List<Point>();

            for (int i = 0; i < first.Length && i < second.Length; i++)
            {
                var mid = Math.Midpoint(first[i], second[i]);
                list.Add(mid);
            }

            return new[] { list.ToArray() };
        }

        public override string ToString()
        {
            return $"Intersection of {Dependencies[0]} and {Dependencies[1]}";
        }
    }
}