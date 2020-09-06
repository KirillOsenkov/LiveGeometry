using System;
using System.Collections.Generic;
using System.Windows;

namespace DynamicGeometry
{
    public class PolygonIntersection : DependentPolygonBase, IPolygon
    {
        public override void Recalculate()
        {
            var first = this.Dependencies.Polygon(0);
            var second = this.Dependencies.Polygon(1);

            var intersection = Intersect(first, second);

            Recreate(intersection.Length, recalculate: false);

            for (int i = 0; i < vertices.Count; i++)
            {
                vertices[i].MoveTo(intersection[i]);
            }

            UpdateVisual();
        }

        protected override void AddSide(int sideCount)
        {
            var side = new Segment();
            side.Drawing = Drawing;

            var index = sides.Count;
            var NumberOfSides = sideCount;

            if (index > 2)
            {
                sides[0].Dependencies[0] = vertices[NumberOfSides - 1];
            }

            if (index == 0)
            {
                side.Dependencies = new[] { vertices[NumberOfSides - 1], vertices[0] };
            }
            else
            {
                side.Dependencies = new[] { vertices[index - 1], vertices[index] };
            }

            sides.Add(side);
            Children.Add(side);
            if (Drawing != null)
            {
                side.OnAddingToCanvas(Drawing.Canvas);
            }

            side.RegisterWithDependencies();
        }

        protected override void RemoveSide()
        {
            var index = sides.Count - 1;
            if (index > 2)
            {
                sides[0].Dependencies[0] = vertices[vertices.Count - 1];
            }

            var side = sides[index];

            side.UnregisterFromDependencies();

            sides.RemoveLast();
            Children.Remove(side);
            if (Drawing != null)
            {
                side.OnRemovingFromCanvas(Drawing.Canvas);
            }
        }

        protected override void AdjustVerticesList(int sideCount)
        {
            if (vertices.Count < sideCount)
            {
                int requiredNumber = sideCount - vertices.Count;
                for (int i = 0; i < requiredNumber; i++)
                {
                    AddVertex();
                }
            }
            else if (vertices.Count > sideCount)
            {
                int requiredNumber = vertices.Count - sideCount;
                for (int i = 0; i < requiredNumber; i++)
                {
                    RemoveVertex();
                }
            }
        }

        public Point[] Intersect(Point[] first, Point[] second)
        {
            var list = new List<Point>();

            for (int i = 0; i < first.Length && i < second.Length; i++)
            {
                var mid = Math.Midpoint(first[i], second[i]);
                list.Add(mid);
            }

            return list.ToArray();
        }

        public override string ToString()
        {
            return $"Intersection of {Dependencies[0]} and {Dependencies[1]}";
        }
    }
}