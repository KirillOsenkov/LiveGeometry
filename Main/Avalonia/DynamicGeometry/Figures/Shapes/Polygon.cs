using System.Collections.Generic;

namespace DynamicGeometry
{
    public partial class Polygon : PolygonBase, IPolygon
    {
        /// <summary>Triangle ABC, square ABCD; one with more vertices than that is numbered</summary>
        public const int MaxVerticesInName = 10;

        protected override string NameFromDependencies()
        {
            return NameFromPoints(MaxVerticesInName);
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
            List<IFigure> newPolyLinePoints = new List<IFigure>();
            List<IFigure> verticesToDelete = new List<IFigure>();

            using (Drawing.ActionManager.CreateTransaction())
            {
                foreach (var vertex in this.Dependencies)
                {
                    IPoint vertexPoint = vertex as IPoint;
                    FreePoint newVertexPoint = Factory.CreateFreePoint(this.Drawing, vertexPoint.Coordinates);
                    Actions.Add(Drawing, newVertexPoint);
                    verticesToDelete.Add(vertexPoint);
                    newPolyLinePoints.Add(newVertexPoint);
                }

                // add last point
                newPolyLinePoints.Add(newPolyLinePoints[0]);

                Polyline newPolyline = Factory.CreatePolyline(this.Drawing, newPolyLinePoints);
                Actions.Add(Drawing, newPolyline);

                // delete main shape
                Actions.Remove(this);

                foreach (var vertexToDelete in verticesToDelete)
                {
                    Actions.Remove(vertexToDelete);
                }
            }
        }

#endif
    }
}
