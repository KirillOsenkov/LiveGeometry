using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Xml;

namespace DynamicGeometry
{
    public class RegularPolygon : CompositeFigure, IShapeWithInterior
    {
        private readonly List<PointBase> vertices;
        private readonly List<Segment> sides;
        private readonly Polygon polygon;

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
                Recreate();
            }
        }

        public RegularPolygon()
        {
            polygon = new Polygon();
            vertices = new List<PointBase>();
            sides = new List<Segment>();
            Children.Add(polygon);
        }

        public override Point Center
        {
            get { return this.Dependencies.Point(0); }
        }

        public Point Vertex
        {
            get { return this.Dependencies.Point(1); }
        }

        public double Area
        {
            get
            {
                return polygon.Area;
            }
        }
        public void Recreate()
        {
            AdjustVerticesList();
            AdjustSides();
            AdjustPolygon();
            Recalculate();
        }

        public override void Recalculate()
        {
            if (sides.Count != NumberOfSides)
            {
                Recreate();
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

        private void AdjustPolygon()
        {
            List<IFigure> allVertices = new List<IFigure>();
            allVertices.Add(this.Dependencies[1]);
            allVertices.AddRange(this.vertices.Cast<IFigure>());
            polygon.Dependencies = allVertices;
        }

        private void AdjustSides()
        {
            if (sides.Count < NumberOfSides)
            {
                for (int i = 0; i < NumberOfSides - sides.Count; i++)
                {
                    AddSide();
                }
            }
            else if (sides.Count > NumberOfSides)
            {
                for (int i = 0; i < sides.Count - NumberOfSides; i++)
                {
                    RemoveSide();
                }
            }
        }

        private void RemoveSide()
        {
            var index = sides.Count - 1;
            if (index > 2)
            {
                sides[index - 1].Dependencies[1] = this.Dependencies[1];
            }
            var side = sides[index];
            sides.RemoveLast();
            Children.Remove(side);
            if (Drawing != null)
            {
                side.OnRemovingFromCanvas(Drawing.Canvas);
            }
        }

        private void AddSide()
        {
            var side = new Segment();
            side.Drawing = Drawing;
            var index = sides.Count;
            if (index > 2)
            {
                sides[index - 1].Dependencies[1] = vertices[index - 1];
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
        }

        private void AdjustVerticesList()
        {
            if (vertices.Count < NumberOfSides - 1)
            {
                int requiredNumber = NumberOfSides - vertices.Count - 1;
                for (int i = 0; i < requiredNumber; i++)
                {
                    AddVertex();
                }
            }
            else if (vertices.Count >= NumberOfSides)
            {
                int requiredNumber = vertices.Count - NumberOfSides;
                for (int i = 0; i <= requiredNumber; i++)
                {
                    RemoveVertex();
                }
            }
        }

        private void RemoveVertex()
        {
            var vertex = vertices[vertices.Count - 1];
            vertices.RemoveLast();
            Children.Remove(vertex);
            if (Drawing != null)
            {
                vertex.OnRemovingFromCanvas(Drawing.Canvas);
            }
        }

        private void AddVertex()
        {
            var vertex = new PointBase();
            vertex.Drawing = Drawing;
            vertices.Add(vertex);
            Children.Add(vertex);
            if (Drawing != null)
            {
                vertex.OnAddingToCanvas(Drawing.Canvas);
            }
        }

#if !PLAYER

        public override void WriteXml(XmlWriter writer)
        {
            if (!Visible)
            {
                writer.WriteAttributeString("Visible", "false");
            }
            if (polygon.Style != null)
            {
                writer.WriteAttributeString("Style", polygon.Style.Name);
            }
        }

#endif

        public override string ToString()
        {
            return NumberOfSides.ToString() + "-gon";
        }
    }
}
