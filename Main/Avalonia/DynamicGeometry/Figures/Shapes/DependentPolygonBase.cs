using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using System.Xml;

namespace DynamicGeometry
{
    /// <summary>
    /// A figure with parts that other figures can be built on, though the parts are not
    /// figures of the drawing themselves: the vertices and sides a regular polygon works out.
    /// A part has no name; a file refers to it by its owner's name and the part's
    /// (<c>&lt;Dependency Name="RegularPolygon1" Part="Vertex3" /&gt;</c>).
    /// </summary>
    public interface IFigureParts : IFigure
    {
        /// <summary>What a file calls the part; null if the figure is not a part of this one</summary>
        string GetPartName(IFigure part);

        /// <summary>The part a file calls so; null if there is none</summary>
        IFigure GetPart(string partName);
    }

    public class DependentPolygonBase : CompositeFigure, IShapeWithInterior, IPolygonalChain, IFigureParts
    {
        protected readonly List<PointBase> vertices = new List<PointBase>();
        protected readonly List<Segment> sides = new List<Segment>();
        protected readonly Polygon polygon = new InteriorPolygon();

        public DependentPolygonBase()
        {
            Children.Add(polygon);
        }

        #region Parts

        // The parts are figures of the library that the drawing never sees: no names of
        // their own (a name would take a letter from the points, or a number from the
        // segments, of the drawing), found through the polygon.

        /// <summary>A vertex the polygon works out</summary>
        public class PolygonVertex : PointBase
        {
            public override void OnAddingToDrawing(Drawing drawing)
            {
            }
        }

        /// <summary>A side between two vertices</summary>
        public class PolygonSide : Segment
        {
            public override void OnAddingToDrawing(Drawing drawing)
            {
            }
        }

        /// <summary>The filled inside</summary>
        public class InteriorPolygon : Polygon
        {
            public override void OnAddingToDrawing(Drawing drawing)
            {
            }
        }

        const string VertexPart = "Vertex";
        const string SidePart = "Side";
        const string InteriorPart = "Interior";

        /// <summary>
        /// Vertex2, Vertex3... (Vertex1 is the point the polygon is built on), Side1, Side2...
        /// counted from that point, and Interior
        /// </summary>
        public string GetPartName(IFigure part)
        {
            int index = vertices.IndexOf(part as PointBase);
            if (index >= 0)
            {
                return VertexPart + (index + 2);
            }

            index = sides.IndexOf(part as Segment);
            if (index >= 0)
            {
                return SidePart + (index + 1);
            }

            return part == polygon ? InteriorPart : null;
        }

        public IFigure GetPart(string partName)
        {
            if (partName == InteriorPart)
            {
                return polygon;
            }

            if (TryGetIndex(partName, VertexPart, out int index))
            {
                index -= 2;
                return index >= 0 && index < vertices.Count ? vertices[index] : null;
            }

            if (TryGetIndex(partName, SidePart, out index))
            {
                index -= 1;
                return index >= 0 && index < sides.Count ? sides[index] : null;
            }

            return null;
        }

        static bool TryGetIndex(string partName, string kind, out int index)
        {
            index = 0;
            return partName != null
                && partName.StartsWith(kind, StringComparison.Ordinal)
                && int.TryParse(partName.Substring(kind.Length), out index);
        }

        /// <summary>Whether something that is not the polygon's own is built on the part</summary>
        protected bool HasOutsideDependents(IFigure part)
        {
            return part.Dependents.Any(dependent => !Children.Contains(dependent));
        }

        /// <summary>
        /// Whether the parts' shapes are on a canvas. Parts made before the polygon is on one
        /// (read from a file, where what is built on a vertex needs the vertex at once) get
        /// there with the polygon, not on their own.
        /// </summary>
        protected bool IsOnCanvas { get; private set; }

        public override void OnAddingToCanvas(Canvas newContainer)
        {
            base.OnAddingToCanvas(newContainer);
            IsOnCanvas = true;
        }

        public override void OnRemovingFromCanvas(Canvas leavingContainer)
        {
            base.OnRemovingFromCanvas(leavingContainer);
            IsOnCanvas = false;
        }

        // The parts are built on the polygon and on the point that is its first vertex.
        // While the polygon is out of the drawing (deleted, its making undone) they are not
        // among the dependents of those: a point would be left with dependents that are in
        // no drawing. Back in, they are again.
        bool partsUnregistered;

        public override void OnRemovingFromDrawing(Drawing drawing)
        {
            base.OnRemovingFromDrawing(drawing);
            if (!partsUnregistered)
            {
                partsUnregistered = true;
                foreach (var part in Children)
                {
                    part.UnregisterFromDependencies();
                }
            }
        }

        public override void OnAddingToDrawing(Drawing drawing)
        {
            base.OnAddingToDrawing(drawing);
            if (partsUnregistered)
            {
                partsUnregistered = false;
                foreach (var part in Children)
                {
                    part.RegisterWithDependencies();
                }
            }
        }

        #endregion

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
        /// dependencies, out of the children, off the canvas. Its own parts depending on it
        /// (the polygon, the sides at the vertex) are rewired by whoever adjusts them.
        /// Nothing outside the polygon is built on it: the count doesn't go below what such
        /// figures need (<see cref="RegularPolygon.NumberOfSides"/>) - taking them away here,
        /// from inside the setter, would be a deletion no undo brings back.
        /// </summary>
        protected void RemovePart(IFigure part)
        {
            part.UnregisterFromDependencies();
            Children.Remove(part);
            if (IsOnCanvas)
            {
                part.OnRemovingFromCanvas(Drawing.Canvas);
            }
        }

        protected void AddVertex()
        {
            var vertex = new PolygonVertex();
            vertex.Dependencies.Add(this);
            vertex.RegisterWithDependencies();
            vertex.Drawing = Drawing;
            vertices.Add(vertex);
            Children.Add(vertex);
            if (IsOnCanvas)
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
            if (Locked)
            {
                writer.WriteAttributeString("Locked", "true");
            }
            if (polygon.Style != null)
            {
                writer.WriteAttributeString("Style", polygon.Style.Name);
            }
        }

#endif
    }
}