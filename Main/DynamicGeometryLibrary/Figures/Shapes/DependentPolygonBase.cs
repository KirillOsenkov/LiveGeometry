using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace DynamicGeometry
{
    public class DependentPolygonBase : CompositeFigure, IShapeWithInterior
    {
        protected readonly List<PointBase> vertices = new List<PointBase>();
        protected readonly List<Segment> sides = new List<Segment>();
        protected readonly Polygon polygon = new Polygon();

        public DependentPolygonBase()
        {
            Children.Add(polygon);
        }

        public virtual void Recreate(int sideCount, bool recalculate = true)
        {
            AdjustVerticesList(sideCount);
            AdjustSides(sideCount);
            AdjustPolygon();

            if (recalculate)
            {
                Recalculate();
            }
        }

        protected virtual void AdjustPolygon()
        {
            List<IFigure> allVertices = new List<IFigure>();
            allVertices.AddRange(this.vertices.Cast<IFigure>());
            polygon.Dependencies = allVertices;
        }

        public double Area => polygon.Area;

        private void AdjustSides(int sideCount)
        {
            var NumberOfSides = sideCount;

            if (sides.Count < NumberOfSides)
            {
                var needed = NumberOfSides - sides.Count;
                for (int i = 0; i < needed; i++)
                {
                    AddSide(sideCount);
                }
            }
            else if (sides.Count > NumberOfSides)
            {
                var extra = sides.Count - NumberOfSides;
                for (int i = 0; i < extra; i++)
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

        protected virtual void AddSide(int sideCount)
        {
        }

        protected virtual void AdjustVerticesList(int sideCount)
        {
        }

        protected void RemoveVertex()
        {
            var vertex = vertices[vertices.Count - 1];
            vertices.RemoveLast();
            Children.Remove(vertex);
            if (Drawing != null)
            {
                vertex.OnRemovingFromCanvas(Drawing.Canvas);
            }
        }

        protected void AddVertex()
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
    }
}