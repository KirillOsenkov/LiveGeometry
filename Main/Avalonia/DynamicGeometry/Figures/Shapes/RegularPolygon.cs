using System;
using Avalonia;

namespace DynamicGeometry
{
    public class RegularPolygon : DependentPolygonBase, IPolygon, IFixableLength
    {
        #region Side

        IPoint CenterPoint
        {
            get { return (IPoint)Dependencies[0]; }
        }

        IPoint VertexPoint
        {
            get { return (IPoint)Dependencies[1]; }
        }

        double RadiusToSide
        {
            get { return 2 * System.Math.Sin(Math.PI / NumberOfSides); }
        }

        /// <summary>
        /// The side, which is the vertex's distance from the center scaled by the number of
        /// sides. Set, it moves the vertex or changes its fixed distance
        /// (<see cref="LengthConstraint.SetDistance"/>).
        /// </summary>
        [PropertyGridVisible]
        [PropertyGridName("Side")]
        [PropertyGridGroup("Side")]
        [PropertyGridPreferredEditor("UpDown")]
        [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
        public double Length
        {
            get
            {
                return Center.Distance(Vertex) * RadiusToSide;
            }
            set
            {
                LengthConstraint.SetDistance(VertexPoint, CenterPoint, value / RadiusToSide);
            }
        }

        public bool CanEdit(string propertyName)
        {
            bool isFixed = LengthConstraint.FixedEnd(VertexPoint, CenterPoint) != null;
            switch (propertyName)
            {
                case "Length":
                    return isFixed || LengthConstraint.CanStretch(VertexPoint, CenterPoint);
                case "FixLength":
                    return !isFixed && LengthConstraint.CanStretch(VertexPoint, CenterPoint);
                case "FreeLength":
                    return isFixed;
                default:
                    return true;
            }
        }

        public string Caption(string propertyName, string defaultCaption)
        {
            return propertyName == "Length" ? "Side" : defaultCaption;
        }

        /// <summary>The polygon keeps its size: the vertex stays at its distance from the center</summary>
        [PropertyGridVisible]
        [PropertyGridName("Fix length")]
        [PropertyGridGroup("Side")]
        [PropertyGridIcon(PropertyGridIcon.Lock)]
        public void FixLength()
        {
            LengthConstraint.Fix(VertexPoint, CenterPoint);
            Drawing.RaiseDisplayProperties(this);
        }

        [PropertyGridVisible]
        [PropertyGridName("Free length")]
        [PropertyGridGroup("Side")]
        [PropertyGridIcon(PropertyGridIcon.Unlock)]
        public void FreeLength()
        {
            var fixedEnd = LengthConstraint.FixedEnd(VertexPoint, CenterPoint);
            if (fixedEnd != null)
            {
                LengthConstraint.Free(fixedEnd);
                Drawing.RaiseDisplayProperties(this);
            }
        }

        #endregion

        private int numberOfSides = 5;
        [PropertyGridVisible]
        [PropertyGridName("Number of sides")]
        [Domain(3, 500)]
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
