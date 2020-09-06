using System;
using System.Collections.Generic;
using System.Windows;

namespace DynamicGeometry
{
    public class RegularPolygon : DependentPolygonBase, IPolygon
    {
        private int numberOfSides = 5;
        [PropertyGridVisible]
        [PropertyGridName("Number of sides")]
        public int NumberOfSides
        {
            get
            {
                return numberOfSides;
            }
            set
            {
                if (value < 3 || value > 500)
                {
                    return;
                }

                numberOfSides = value;
                Recreate(numberOfSides);
                this.RecalculateAllDependents();
            }
        }

        public override Point Center
        {
            get { return this.Dependencies.Point(0); }
        }

        public Point Vertex
        {
            get { return this.Dependencies.Point(1); }
        }

        public override void Recalculate()
        {
            if (sides.Count != NumberOfSides)
            {
                Recreate(NumberOfSides);
                return;
            }

            var center = Center;
            var vertex = Vertex;

            double initialAngle = Math.GetAngle(center, vertex);
            double radius = center.Distance(vertex);
            double increment = Math.DOUBLEPI / NumberOfSides;

            for (int i = 0; i < NumberOfSides - 1; i++)
            {
                double angle = initialAngle + (i + 1) * increment;

                double X = center.X + radius * System.Math.Cos(angle);
                double Y = center.Y + radius * System.Math.Sin(angle);

                vertices[i].MoveTo(new Point(X, Y));
            }

            this.UpdateVisual();
        }

        protected override void CollectPolygonDependencies(Action<IFigure> callback)
        {
            callback(this.Dependencies[1]);
            base.CollectPolygonDependencies(callback);
        }

        protected override void AdjustVerticesList(int sideCount)
        {
            if (vertices.Count < sideCount - 1)
            {
                int requiredNumber = sideCount - vertices.Count - 1;
                for (int i = 0; i < requiredNumber; i++)
                {
                    AddVertex();
                }
            }
            else if (vertices.Count >= sideCount)
            {
                int requiredNumber = vertices.Count - sideCount;
                for (int i = 0; i <= requiredNumber; i++)
                {
                    RemoveVertex();
                }
            }
        }

        protected override void AddSide(int sideCount)
        {
            var side = new Segment();
            side.Drawing = Drawing;
            var index = sides.Count;
            var NumberOfSides = sideCount;
            if (index > 2)
            {
                var firstSide = sides[index - 1];
                firstSide.UnregisterFromDependencies();
                firstSide.Dependencies[1] = vertices[index - 1];
                firstSide.RegisterWithDependencies();
            }

            if (index == 0)
            {
                side.Dependencies = new[] { this.Dependencies[1], vertices[0] };
            }
            else if (index == NumberOfSides - 1)
            {
                side.Dependencies = new[] { vertices[NumberOfSides - 2], this.Dependencies[1] };
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
                var firstSide = sides[index - 1];
                firstSide.UnregisterFromDependencies();
                firstSide.Dependencies[1] = this.Dependencies[1];
                firstSide.RegisterWithDependencies();
            }

            var side = sides[index];

            side.UnregisterFromDependencies();

            sides.RemoveLast();

            var drawing = Drawing;
            var action = new RemoveFigureAction(drawing, side);
            action.Execute();

            Children.Remove(side);

            if (Drawing != null)
            {
                side.OnRemovingFromCanvas(Drawing.Canvas);
            }
        }

        public override string ToString()
        {
            return NumberOfSides.ToString() + "-gon";
        }
    }
}
