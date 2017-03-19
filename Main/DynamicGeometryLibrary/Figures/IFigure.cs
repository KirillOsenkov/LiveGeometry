using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
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

        int ZIndex { get; set; }
        bool Visible { get; set; }

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