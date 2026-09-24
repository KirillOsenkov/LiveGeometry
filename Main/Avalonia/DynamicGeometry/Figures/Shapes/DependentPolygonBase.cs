using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;
using System.Xml;

namespace DynamicGeometry
{
    public class DependentPolygonBase : CompositeFigure, IShapeWithInterior, IPolygonalChain
    {
        protected readonly List<PointBase> vertices = new List<PointBase>();
        protected readonly List<Segment> sides = new List<Segment>();
        protected readonly Polygon polygon = new Polygon();

        public DependentPolygonBase()
        {
            Children.Add(polygon);
        }

        /// <summary>
        /// A click on any child selects the whole figure, so the property grid shows this
        /// composite and its style is the one the user edits: it is the polygon's style (the
        /// one WriteXml saves). Without this the composite kept a style of its own that
        /// nothing painted with, and changing it did nothing.
        /// </summary>
        public override IFigureStyle Style
        {
            get
            {
                return polygon.Style;
            }
            set
            {
                polygon.Style = value;
            }
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

        protected virtual void CollectPolygonDependencies(Action<IFigure> collector)
        {
            foreach (var item in vertices)
            {
                collector(item);
            }

            collector(this);
        }

        protected void AdjustPolygon()
        {
            polygon.UnregisterFromDependencies();
            List<IFigure> allVertices = new List<IFigure>();
            CollectPolygonDependencies(allVertices.Add);
            polygon.Dependencies = allVertices;
            polygon.RegisterWithDependencies();
        }

        public double Area => polygon.Area;

        public Point[] VertexCoordinates => vertices.Select(v => v.Coordinates).ToArray();

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

        protected virtual void RemoveSide()
        {
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
            vertex.UnregisterFromDependencies();
            vertices.RemoveLast();

            var drawing = vertex.Drawing;
            var action = new RemoveFigureAction(drawing, vertex);
            action.Execute();
            Children.Remove(vertex);

            if (Drawing != null)
            {
                vertex.OnRemovingFromCanvas(Drawing.Canvas);
            }
        }

        protected void AddVertex()
        {
            var vertex = new PointBase();
            vertex.Dependencies.Add(this);
            vertex.RegisterWithDependencies();
            vertex.Drawing = Drawing;
            vertices.Add(vertex);
            Children.Add(vertex);
            if (Drawing != null)
            {
                vertex.OnAddingToCanvas(Drawing.Canvas);
            }

            Drawing.StyleManager.SetStyleIfAvailable(vertex, "DependentPointStyle");
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