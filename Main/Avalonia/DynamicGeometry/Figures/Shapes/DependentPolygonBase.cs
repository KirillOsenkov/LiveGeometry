using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;

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

        /// <summary>
        /// The parts a click selects by themselves, to be styled (the vertices and the sides);
        /// a click on any other part (the inside) selects the figure
        /// </summary>
        IEnumerable<IFigure> SelectableParts { get; }
    }

    /// <summary>A part of a figure (<see cref="IFigureParts"/>)</summary>
    public interface IFigurePart : IFigure
    {
        IFigureParts Owner { get; }
    }

    /// <summary>
    /// What a selection means where a figure has parts. A part selected by itself is styled
    /// by itself; everything else done to the selection (delete, hide, lock, copy) is done
    /// to the figure it belongs to.
    /// </summary>
    public static class FigureParts
    {
        /// <summary>What a click on the figure selects: a vertex or a side by itself, but the inside is the figure</summary>
        public static IFigure SelectionTarget(IFigure figure)
        {
            return figure is IFigurePart part && !part.Owner.SelectableParts.Contains(figure)
                ? part.Owner
                : figure;
        }

        /// <summary>The figure of the drawing that the figure is, or is a part of</summary>
        public static IFigure Whole(IFigure figure)
        {
            return figure is IFigurePart part ? part.Owner : figure;
        }

        /// <summary>The figures of the drawing that these are, or are parts of, each once</summary>
        public static IEnumerable<IFigure> Wholes(IEnumerable<IFigure> figures)
        {
            return figures.Select(Whole).Distinct();
        }

        /// <summary>The figure and its parts that can be selected by themselves</summary>
        public static IEnumerable<IFigure> WithParts(IFigure figure)
        {
            yield return figure;
            if (figure is IFigureParts parts)
            {
                foreach (var part in parts.SelectableParts)
                {
                    yield return part;
                }
            }
        }

        /// <summary>Whether the figure is selected, or a part of it by itself</summary>
        public static bool HasSelection(IFigure figure)
        {
            return figure.Selected
                || figure is IFigureParts parts && parts.SelectableParts.Any(part => part.Selected);
        }
    }

    public class DependentPolygonBase : CompositeFigure, IShapeWithInterior, IPerimeter, IPolygonalChain, IFigureParts
    {
        protected readonly List<PointBase> vertices = new List<PointBase>();
        protected readonly List<Segment> sides = new List<Segment>();
        protected readonly Polygon polygon;
        readonly Stack<PointBase> retiredVertices = new Stack<PointBase>();

        public DependentPolygonBase()
        {
            // the polygon itself is drawn by its parts; its layer says the verbs apply to it
            Layer = ZOrder.Polygons;
            polygon = new InteriorPolygon(this);
            Children.Add(polygon);
        }

        #region Parts

        // The parts are figures of the library that the drawing never sees: no names of
        // their own (a name would take a letter from the points, or a number from the
        // segments, of the drawing), found through the polygon. A vertex or a side can be
        // selected by itself and has a style of its own; its page in the grid has that and
        // little else (ICustomPropertyProvider), since the rows it would inherit as a point
        // or a segment (a name, Delete, Convert to line...) are the polygon's to decide.

        /// <summary>A vertex the polygon works out</summary>
        public class PolygonVertex : PointBase, IFigurePart, ICustomPropertyProvider, ICustomMethodProvider
        {
            public PolygonVertex(DependentPolygonBase owner)
            {
                Owner = owner;
            }

            public DependentPolygonBase Owner { get; }

            IFigureParts IFigurePart.Owner => Owner;

            public override void OnAddingToDrawing(Drawing drawing)
            {
            }

            public IEnumerable<IValueProvider> GetProperties()
            {
                return Owner.GetPartProperties(this);
            }

            public IEnumerable<IOperationDescription> GetMethods()
            {
                return Owner.GetPartMethods(this);
            }

            public override string ToString()
            {
                return Owner.DescribePart(this);
            }
        }

        /// <summary>A side between two vertices</summary>
        public class PolygonSide : Segment, IFigurePart, ICustomPropertyProvider, ICustomMethodProvider
        {
            public PolygonSide(DependentPolygonBase owner)
            {
                Owner = owner;
            }

            public DependentPolygonBase Owner { get; }

            IFigureParts IFigurePart.Owner => Owner;

            public override void OnAddingToDrawing(Drawing drawing)
            {
            }

            public IEnumerable<IValueProvider> GetProperties()
            {
                return Owner.GetPartProperties(this);
            }

            public IEnumerable<IOperationDescription> GetMethods()
            {
                return Owner.GetPartMethods(this);
            }

            public override string ToString()
            {
                return Owner.DescribePart(this);
            }
        }

        /// <summary>The filled inside: to the user, the polygon itself</summary>
        public class InteriorPolygon : Polygon, IFigurePart
        {
            public InteriorPolygon(DependentPolygonBase owner)
            {
                Owner = owner;
            }

            public DependentPolygonBase Owner { get; }

            IFigureParts IFigurePart.Owner => Owner;

            public override void OnAddingToDrawing(Drawing drawing)
            {
            }

            public override string ToString()
            {
                return Owner.DescribePart(this);
            }
        }

        public IEnumerable<IFigure> SelectableParts => vertices.Concat<IFigure>(sides);

        // whether the whole polygon is selected: a vertex or a side selected by itself leaves
        // this off (it is the part that the grid shows), while selecting the whole selects
        // every part with it, so that they all show the halo
        bool selectedWhole;

        public override bool Selected
        {
            get
            {
                return selectedWhole;
            }
            set
            {
                selectedWhole = value;
                base.Selected = value;
            }
        }

        /// <summary>"Side 2 of regular pentagon p": the title of a part's page</summary>
        public string DescribePart(IFigure part)
        {
            var partName = GetPartName(part);
            if (partName == null)
            {
                return ToString();
            }

            // the inside stands for the polygon itself (FigureParts.SelectionTarget): a
            // perimeter "of regular pentagon p", the hover's "Regular pentagon p" (it read
            // the name as a side's, "Side rior of", and the inside had no name at all)
            if (partName == InteriorPart)
            {
                return Reference;
            }

            var kind = partName.StartsWith(VertexPart, StringComparison.Ordinal) ? VertexPart : SidePart;
            return kind + " " + partName.Substring(kind.Length) + " of " + Reference;
        }

        /// <summary>Polygon p, q... (no points of its own name it)</summary>
        protected override string FirstLetter
        {
            get
            {
                return "p";
            }
        }

        /// <summary>The rows of a part's page: its style, and what the polygon adds (<see cref="GetPartValues"/>)</summary>
        public IEnumerable<IValueProvider> GetPartProperties(IFigure part)
        {
            return GetPartValues(part)
                .Append(PropertyDiscoveryStrategy.CreateValueProvider(part, nameof(StyleDisplay)));
        }

        /// <summary>Rows of the polygon's own to show on a part's page (a side's length)</summary>
        protected virtual IEnumerable<IValueProvider> GetPartValues(IFigure part)
        {
            return Enumerable.Empty<IValueProvider>();
        }

        public IEnumerable<IOperationDescription> GetPartMethods(IFigure part)
        {
            yield return MethodDescription.Create(typeof(FigureBase).GetMethod(nameof(EditStyleButton)));
            yield return MethodDescription.Create(typeof(FigureBase).GetMethod(nameof(CreateNewStyle)));
            yield return new DelegateOperation(
                "SelectWhole",
                "Select " + Reference,
                PropertyGridIcon.Polygon,
                SelectWhole);
        }

        /// <summary>The polygon instead of a part of it, in the selection and in the grid</summary>
        public void SelectWhole()
        {
            if (Drawing == null)
            {
                return;
            }

            Drawing.Figures.ClearSelection();
            Selected = true;
            Drawing.RaiseSelectionChanged(Drawing.GetSelectedFigures());
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
        // While the polygon is not in the drawing (read from a file and not added yet,
        // deleted, its making undone) they are not among the dependents of those: a point
        // would have dependents that are in no drawing, which fails the consistency check
        // (a file with a regular polygon on a point by coordinates did not load). In the
        // drawing, they are.
        bool partsUnregistered = true;

        /// <summary>Lists a part with what it is built on, if the polygon is in the drawing; else that waits until it is</summary>
        protected void RegisterPart(IFigure part)
        {
            if (!partsUnregistered)
            {
                part.RegisterWithDependencies();
            }
        }

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
        /// The style of the inside, which is the polygon's own: a click inside selects the
        /// polygon, and this is what its grid edits as the fill. Without this the composite
        /// kept a style of its own that nothing painted with, and changing it did nothing.
        /// The sides and vertices have styles of their own (<see cref="SideStyles"/>).
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

        [PropertyGridName("Fill")]
        public override IFigureStyle StyleDisplay
        {
            get
            {
                return base.StyleDisplay;
            }
            set
            {
                base.StyleDisplay = value;
            }
        }

        /// <summary>All the sides at once; one of them by itself is selected with a click on it</summary>
        [PropertyGridVisible]
        public PartStylesValue SideStyles
        {
            get
            {
                return sides.Count > 0 ? new PartStylesValue(nameof(SideStyles), "Sides", sides) : null;
            }
        }

        /// <summary>The vertices the polygon works out, all at once (the one it is built on is a point of its own)</summary>
        [PropertyGridVisible]
        public PartStylesValue VertexStyles
        {
            get
            {
                return vertices.Count > 0 ? new PartStylesValue(nameof(VertexStyles), "Vertices", vertices) : null;
            }
        }

        const string SidesElement = "Sides";
        const string VerticesElement = "Vertices";
        const string PartElement = "Part";

        /// <summary>
        /// The styles of the parts, as elements after the attributes: the style most sides
        /// have (<c>&lt;Sides Style="RedLine" /&gt;</c>, left out when it is the default),
        /// likewise for the vertices, and a part with another style by itself
        /// (<c>&lt;Part Name="Side2" Style="BlueLine" /&gt;</c>).
        /// </summary>
        protected void WritePartStyles(XmlWriter writer)
        {
            WritePartStyles(writer, SidesElement, sides);
            WritePartStyles(writer, VerticesElement, vertices);
        }

        void WritePartStyles(XmlWriter writer, string elementName, IEnumerable<IFigure> parts)
        {
            var styled = parts.Where(part => part.Style != null).ToList();
            if (styled.Count == 0 || Drawing == null)
            {
                return;
            }

            var common = styled
                .GroupBy(part => part.Style)
                .OrderByDescending(group => group.Count())
                .First()
                .Key;
            if (common != Drawing.StyleManager.AssignDefaultStyle(styled[0]))
            {
                writer.WriteStartElement(elementName);
                writer.WriteAttributeString("Style", common.Name);
                writer.WriteEndElement();
            }

            foreach (var part in styled.Where(part => part.Style != common))
            {
                writer.WriteStartElement(PartElement);
                writer.WriteAttributeString("Name", GetPartName(part));
                writer.WriteAttributeString("Style", part.Style.Name);
                writer.WriteEndElement();
            }
        }

        /// <summary>The styles <see cref="WritePartStyles(XmlWriter)"/> wrote, once the parts are made</summary>
        protected void ReadPartStyles(XElement element)
        {
            var manager = Drawing?.StyleManager;
            if (manager == null)
            {
                return;
            }

            void Apply(IFigure part, string styleName)
            {
                var style = styleName != null ? manager[styleName] : null;
                if (part != null && style != null && style.GetType().SupportsFigureType(part.GetType()))
                {
                    part.Style = style;
                }
            }

            foreach (var side in sides)
            {
                Apply(side, (string)element.Element(SidesElement)?.Attribute("Style"));
            }

            foreach (var vertex in vertices)
            {
                Apply(vertex, (string)element.Element(VerticesElement)?.Attribute("Style"));
            }

            foreach (var part in element.Elements(PartElement))
            {
                Apply(GetPart((string)part.Attribute("Name")), (string)part.Attribute("Style"));
            }
        }

        public virtual void Recreate(int sideCount, bool recalculate = true)
        {
            AdjustVerticesList(sideCount);
            AdjustSides(sideCount);
            AdjustPolygon();

            // New parts are shown, and a hidden polygon must hide them too: read from a
            // file, it made its parts after it was told it was hidden, and its vertices and
            // sides came back on screen, to be clicked, with the polygon itself not there.
            // (Likewise when its number of sides went up while hidden.)
            if (!Visible)
            {
                foreach (var part in Children)
                {
                    part.Visible = false;
                }
            }

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
            RegisterPart(polygon);
        }

        public double Area => polygon.Area;

        public double Perimeter => polygon.Perimeter;

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
            retiredVertices.Push(vertex);
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

        /// <summary>
        /// The style a new part takes: the one most of the others have (sides made red get a
        /// red side when there are more of them), or none yet. A part that comes back (fewer
        /// sides, then more) keeps its own, or undo of the fewer would not restore it.
        /// </summary>
        protected static IFigureStyle CommonStyle(IEnumerable<IFigure> parts)
        {
            return parts
                .Where(part => part.Style != null)
                .GroupBy(part => part.Style)
                .OrderByDescending(group => group.Count())
                .FirstOrDefault()?
                .Key;
        }

        protected void AddVertex()
        {
            bool isNew = retiredVertices.Count == 0;
            var vertex = isNew ? new PolygonVertex(this) : retiredVertices.Pop();
            var common = isNew ? CommonStyle(vertices) : null;
            vertex.Dependencies = new IFigure[] { this };
            vertex.Visible = Visible;
            vertex.Selected = Selected;
            RegisterPart(vertex);
            vertex.Drawing = Drawing;
            vertices.Add(vertex);
            Children.Add(vertex);
            if (IsOnCanvas)
            {
                vertex.OnAddingToCanvas(Drawing.Canvas);
            }

            if (common != null)
            {
                vertex.Style = common;
            }
            else if (isNew)
            {
                Drawing.StyleManager.SetStyleIfAvailable(vertex, StyleManager.DependentPointStyleName);
            }
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