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
            vertices.RemoveLast();
            RemovePart(vertex);
        }

        /// <summary>
        /// A part the count no longer needs (a vertex, a side) leaves: out of its
        /// dependencies, out of the children, off the canvas. Only a part something outside
        /// the polygon is built on goes through the whole deletion, which takes the
        /// dependents along and tells the selection. Its own parts depending on it (the
        /// polygon, the sides at the vertex) are rewired by whoever adjusts them, so they
        /// don't count: the deletion per part would rebuild the property grid once per
        /// vertex and once per side of a 500-gon becoming a triangle.
        /// </summary>
        protected void RemovePart(IFigure part)
        {
            part.UnregisterFromDependencies();
            if (part.Dependents.Any(dependent => !Children.Contains(dependent)))
            {
                var action = new RemoveFigureAction(Drawing, part);
                action.Execute();
            }

            Children.Remove(part);
            if (Drawing != null)
            {
                part.OnRemovingFromCanvas(Drawing.Canvas);
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

            Drawing.StyleManager.SetStyleIfAvailable(vertex, StyleManager.DependentPointStyleName);
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