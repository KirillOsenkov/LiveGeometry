using System;
using System.Collections.Generic;
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
        [PropertyGridCustomValueProvider(typeof(LengthPropertyValue))]
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

        public IPoint LengthEndpoint => VertexPoint;

        // the sides are the polygon's own children, nothing in the drawing to measure
        public IList<IFigure> MeasuredFigures
        {
            get { return null; }
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

        private int numberOfSides = DefaultNumberOfSides;
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

                // a vertex or a side that something is built on stays: fewer sides would
                // take that figure away with it, from inside this setter, where no undo
                // would bring it back
                int minimum = MinimumNumberOfSides();
                if (value < minimum)
                {
                    value = minimum;
                    Drawing?.RaiseStatusNotification(
                        Name + " keeps " + minimum + " sides: figures are built on its vertices or sides.");
                }

                numberOfSides = value;
                Recreate(numberOfSides);
                this.RecalculateAllDependents();
                // the title says it: Regular pentagon
                RaisePropertyChanged(nameof(NumberOfSides));

                // and the side is another, with the vertex where it was
                RaisePropertyChanged(nameof(Length));
            }
        }

        /// <summary>The fewest sides that keep every vertex and side something outside the polygon is built on</summary>
        int MinimumNumberOfSides()
        {
            int minimum = 3;
            for (int i = 0; i < vertices.Count; i++)
            {
                // vertices[0] is the second vertex
                if (i + 2 > minimum && HasOutsideDependents(vertices[i]))
                {
                    minimum = i + 2;
                }
            }

            for (int i = 0; i < sides.Count; i++)
            {
                if (i + 1 > minimum && HasOutsideDependents(sides[i]))
                {
                    minimum = i + 1;
                }
            }

            return minimum;
        }

        const int DefaultNumberOfSides = 5;

        public override void WriteXml(System.Xml.XmlWriter writer)
        {
            base.WriteXml(writer);
            if (numberOfSides != DefaultNumberOfSides)
            {
                writer.WriteAttributeString("Sides", numberOfSides.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }

            WritePartStyles(writer);
        }

        public override void ReadXml(System.Xml.Linq.XElement element)
        {
            base.ReadXml(element);
            int count = (int)element.ReadDouble("Sides");
            if (count >= 3 && count <= 500)
            {
                numberOfSides = count;
            }

            // the parts now, not at the first recalculation: a figure of the file that is
            // built on a vertex or a side looks for it as soon as it is read
            if (Drawing != null)
            {
                Recreate(numberOfSides, recalculate: false);
                ReadPartStyles(element);
            }
        }

        /// <summary>A side's page has the side's length, which is every side's</summary>
        protected override IEnumerable<IValueProvider> GetPartValues(IFigure part)
        {
            if (part is PolygonSide)
            {
                yield return PropertyDiscoveryStrategy.CreateValueProvider(this, nameof(Length));
            }
        }

        public override Point Center
        {
            get { return Point(0); }
        }

        public Point Vertex
        {
            get { return Point(1); }
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
            bool isNew = retiredSides.Count == 0;
            var side = isNew ? new PolygonSide(this) : retiredSides.Pop();
            var common = isNew ? CommonStyle(sides) : null;
            side.Drawing = Drawing;
            side.Visible = Visible;
            side.Selected = Selected;
            var index = sides.Count;
            var NumberOfSides = sideCount;
            if (index > 2)
            {
                var firstSide = sides[index - 1];
                firstSide.UnregisterFromDependencies();
                firstSide.Dependencies[1] = vertices[index - 1];
                RegisterPart(firstSide);
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
            if (IsOnCanvas)
            {
                side.OnAddingToCanvas(Drawing.Canvas);
            }

            if (common != null)
            {
                side.Style = common;
            }

            RegisterPart(side);
        }

        protected override void RemoveSide()
        {
            var index = sides.Count - 1;
            if (index > 2)
            {
                var firstSide = sides[index - 1];
                firstSide.UnregisterFromDependencies();
                firstSide.Dependencies[1] = this.Dependencies[1];
                RegisterPart(firstSide);
            }

            var side = sides[index];
            sides.RemoveLast();
            RemovePart(side);
            retiredSides.Push(side);
        }

        readonly Stack<Segment> retiredSides = new Stack<Segment>();

        /// <summary>"Regular pentagon", "Square", "Regular 13-gon"</summary>
        protected override string Kind
        {
            get
            {
                switch (NumberOfSides)
                {
                    case 3:
                        return "Equilateral triangle";
                    case 4:
                        return "Square";
                    case 5:
                        return "Regular pentagon";
                    case 6:
                        return "Regular hexagon";
                    case 7:
                        return "Regular heptagon";
                    case 8:
                        return "Regular octagon";
                    case 9:
                        return "Regular nonagon";
                    case 10:
                        return "Regular decagon";
                    case 12:
                        return "Regular dodecagon";
                    default:
                        return "Regular " + NumberOfSides + "-gon";
                }
            }
        }

        /// <summary>"with center F and vertex G"</summary>
        public override string Construction
        {
            get
            {
                if (Dependencies.Count < 2)
                {
                    return null;
                }

                return "with center " + ConstructionText.Of(CenterPoint) + " and vertex " + ConstructionText.Of(VertexPoint);
            }
        }

        /// <summary>The name, not the dump of the parts a composite gives</summary>
        public override string ToString()
        {
            return Name;
        }
    }
}
