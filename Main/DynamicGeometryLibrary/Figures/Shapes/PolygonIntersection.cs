using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using XP;

namespace DynamicGeometry
{
    public class PolygonIntersection : CompositeFigure
    {
        private Point[][] intersections;

        public static bool AreEqual(Point[][] left, Point[][] right)
        {
            if ((left == null) != (right == null))
            {
                return false;
            }

            if (left.Length != right.Length)
            {
                return false;
            }

            for (int i = 0; i < left.Length; i++)
            {
                if (!AreEqual(left[i], right[i]))
                {
                    return false;
                }
            }

            return true;
        }

        public static bool AreEqual(Point[] left, Point[] right)
        {
            if ((left == null) != (right == null))
            {
                return false;
            }

            if (left.Length != right.Length)
            {
                return false;
            }

            for (int i = 0; i < left.Length; i++)
            {
                if (left[i] != right[i])
                {
                    return false;
                }
            }

            return true;
        }

        public override void Recalculate()
        {
            var first = this.Dependencies.Polygon(0);
            var second = this.Dependencies.Polygon(1);

            var newIntersections = Intersect(first, second);
            if (AreEqual(intersections, newIntersections))
            {
                return;
            }

            intersections = newIntersections;

            ClearChildren();

            foreach (var intersection in newIntersections)
            {
                var vertices = new List<PointBase>(intersection.Length);

                foreach (var vertex in intersection)
                {
                    var point = new PointBase();
                    point.MoveToCore(vertex);
                    point.Dependencies.Add(this);
                    vertices.Add(point);
                    AddChild(point);
                    Drawing.StyleManager.SetStyleIfAvailable(point, "DependentPointStyle");
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
                    Drawing.StyleManager.SetStyleIfAvailable(side, "OtherLine");
                }

                var polygon = new Polygon();
                polygon.Dependencies.AddRange(vertices);
                AddChild(polygon);
                Drawing.StyleManager.SetStyleIfAvailable(polygon, "OtherShape");
            }

            UpdateVisual();
        }

        public Point[][] Intersect(Point[] first, Point[] second)
        {
            GbPolyg gbPolyg1 = new GbPolyg();
            GbPolyg gbPolyg2 = new GbPolyg();
            foreach (Point point in first)
                gbPolyg1.Add((point.X, point.Y));
            foreach (Point point in second)
                gbPolyg2.Add((point.X, point.Y));
            Polyg polyg1 = gbPolyg1.Polyg();
            Polyg polyg2 = gbPolyg2.Polyg();
            List<Point> list = new List<Point>();
            foreach (Pt pt in polyg1.PolygIntersect(polyg2))
                list.Add(new Point(pt.x, pt.y));

            return new[] { list.ToArray() };
        }

        public override string ToString()
        {
            return $"Intersection of {Dependencies[0]} and {Dependencies[1]}";
        }
    }
}
