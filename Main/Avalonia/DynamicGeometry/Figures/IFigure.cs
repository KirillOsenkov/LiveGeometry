using System;
using System.Collections.Generic;
using Avalonia;
using Avalonia.Controls;
using System.Xml;
using System.Xml.Linq;

namespace DynamicGeometry
{
    public partial interface IFigure : IEquatable<IFigure>
    {
        Drawing Drawing { get; set; }

        IList<IFigure> Dependencies { get; set; }
        IList<IFigure> Dependents { get; }

        IFigure Clone();

        string Name { get; set; }

        /// <summary>Nobody has named the figure (see <see cref="FigureBase.HasDefaultName"/>)</summary>
        bool HasDefaultName { get; }

        /// <summary>"Segment AB": what the property grid calls the figure (see <see cref="FigureBase.Title"/>)</summary>
        string Title { get; }

        /// <summary>"of CD": how the figure is built, after its title (see <see cref="FigureBase.Construction"/>)</summary>
        string Construction { get; }

        bool Exists { get; set; }
        bool Selected { get; set; }
        bool Enabled { get; set; }
        bool Locked { get; set; }
        bool IsHitTestVisible { get; set; }
        void UpdateExistence();
        void Recalculate();
        void UpdateVisual();

        IFigureStyle Style { get; set; }
        void ApplyStyle();

        /// <summary>
        /// Determines if a point lies on a figure and returns the figure in this case.
        /// </summary>
        /// <param name="point">Point's logical coordinates</param>
        /// <returns>A figure (usually itself or a child) if a point is on this figure, null otherwise</returns>
        IFigure HitTest(Point point);

        /// <summary>
        /// Usually the geometric centroid. Defined in IFigure for labeling purposes. Figures without geometric centers should return a sensible value. Default is (0,0).
        /// </summary>
        Point Center { get; }

        /// <summary>
        /// Unused in Live Geometry but used in Tabula. This tag is used for example when a figure having a PointOnFigure is reflected.
        /// </summary>
        bool Flipped { get; set; }

        void OnAddingToCanvas(Canvas canvas);
        void OnRemovingFromCanvas(Canvas canvas);

        void OnAddingToDrawing(Drawing drawing);
        void OnRemovingFromDrawing(Drawing drawing);

        /// <summary>The layer the figure is drawn in, its kind's (<see cref="ZOrder"/>)</summary>
        ZOrder Layer { get; set; }

        /// <summary>
        /// Where the figure is among those of its band of layers: 0 unless Bring to front or
        /// Send to back changed it (<see cref="ZOrders"/>); saved as <c>Z</c>
        /// </summary>
        int Z { get; set; }

        /// <summary>What the layer and the Z come to: the order the figures are drawn and hit in</summary>
        int ZIndex { get; }

        bool Visible { get; set; }

        /// <summary>
        /// Created on demand for another figure (a Number holding a typed value): removed
        /// along with its last dependent.
        /// </summary>
        bool Auxiliary { get; set; }

        string GenerateFigureName();

        /// <param name="blacklist">A list of names to exclude. Can be null.</param>
        string GenerateFigureName(List<string> blacklist);

#if !PLAYER
        void WriteXml(XmlWriter writer);
#endif
        void ReadXml(XElement element);

        /// <summary>
        /// Should a figure be serialized?  In the DG Library only CartesianGrid returns false.
        /// </summary>
        bool Serializable { get; }
    }
}