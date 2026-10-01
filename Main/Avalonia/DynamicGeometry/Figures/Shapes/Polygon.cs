using System.Collections.Generic;
using GuiLabs.Undo;

namespace DynamicGeometry
{
    public partial class Polygon : PolygonBase, IPolygon
    {
        /// <summary>Triangle ABC, square ABCD; one with more vertices than that is numbered</summary>
        public const int MaxVerticesInName = 10;

        protected override IReadOnlyList<string> NamesFromDependencies()
        {
            return NamesFromPoints(PointOrder.Cyclic, MaxVerticesInName);
        }

        protected override string Kind
        {
            get
            {
                switch (Dependencies.Count)
                {
                    case 3:
                        return "Triangle";
                    case 4:
                        return Quadrilaterals.Classify(Point(0), Point(1), Point(2), Point(3));
                    case 5:
                        return "Pentagon";
                    case 6:
                        return "Hexagon";
                    default:
                        return "Polygon";
                }
            }
        }

#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert to polyline")]
        [PropertyGridIcon(PropertyGridIcon.Polyline)]
        public void ConvertToPolyline()
        {
            // The outline through the same vertices and back to the first: ABCA. (It used to
            // be drawn on new free points, and the vertices deleted - with everything else
            // built on them, which in a construction is most of the drawing; redo of that threw.)
            var drawing = Drawing;
            var points = new List<IFigure>(Dependencies);
            points.Add(points[0]);
            var polyline = Factory.CreatePolyline(drawing, points);

            using (Transaction.Create(drawing.ActionManager, delayed: false))
            {
                Actions.Add(drawing, polyline);

                // in the polygon's place in the list
                Actions.MoveBefore(drawing, polyline, this);
                Actions.Remove(this);
            }

            drawing.RaiseUserIsAddingFigures(new Drawing.UIAFEventArgs() { Figures = polyline.AsEnumerable<IFigure>() });
        }

#endif
    }
}
