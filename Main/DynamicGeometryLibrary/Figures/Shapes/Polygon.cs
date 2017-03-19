using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace DynamicGeometry
{
    public partial class Polygon : PolygonBase
    {
#if !PLAYER && !TABULA

        [PropertyGridVisible]
        [PropertyGridName("Convert to Polyline")]
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
