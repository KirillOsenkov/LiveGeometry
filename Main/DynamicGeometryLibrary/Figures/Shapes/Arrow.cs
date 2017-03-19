using System.Linq;
using System.Windows;
using System.Windows.Media;

namespace DynamicGeometry
{
    public class Arrow : Polygon
    {
        public Arrow()
        {
            Shape.Stroke = null;
            pointCache.Add(new Point(), new Point(), new Point(), new Point(), new Point(), new Point(), new Point());
            Shape.Points = pointCache;
        }

        protected override void OnDependenciesChanged()
        {
        }

        PointCollection pointCache = new PointCollection();

        public override void UpdateVisual()
        {
            if (Drawing == null)
            {
                return;
            }

            if (vertexCoordinates == null)
            {
                vertexCoordinates = new Point[7];
            }

            PointPair line = Dependencies.Line(0);
            LineBase parentLine = Dependencies.ElementAt(0) as LineBase;
            if (parentLine != null)
            {
                line = parentLine.OnScreenCoordinates;
            }

            Point p1 = line.P1;
            Point p2 = line.P2;
            double d = p1.Distance(p2);
            var size = ToLogical((Style as LineStyle).StrokeWidth) + .5;
            double arrowLength = size;// ToLogical(16);
            double arrowWidth = size;// 1.0 / 3;
            double shaftWidth = arrowWidth / 3;
            Point triangleBase = new Point(
                p2.X + (p1.X - p2.X) * arrowLength / d,
                p2.Y + (p1.Y - p2.Y) * arrowLength / d);

            // TODO: need to measure performance - I don't know what's 
            // faster - creating a new collection or emptying and
            // refilling the existing collection.
            // Gut feeling is that emptying and refilling should be faster
            // but I need to measure to confirm that.
            VertexCoordinates[0] = new Point(
                    triangleBase.X + (p2.Y - triangleBase.Y) * arrowWidth,
                    triangleBase.Y + (triangleBase.X - p2.X) * arrowWidth);
            VertexCoordinates[1] = p2;
            VertexCoordinates[2] = new Point(
                triangleBase.X + (triangleBase.Y - p2.Y) * arrowWidth,
                triangleBase.Y + (p2.X - triangleBase.X) * arrowWidth);
            VertexCoordinates[3] = new Point(
                triangleBase.X + (triangleBase.Y - p2.Y) * shaftWidth,
                triangleBase.Y + (p2.X - triangleBase.X) * shaftWidth);
            VertexCoordinates[4] = new Point(
                p1.X + (triangleBase.Y - p2.Y) * shaftWidth,
                p1.Y + (p2.X - triangleBase.X) * shaftWidth);
            VertexCoordinates[5] = new Point(
                    p1.X + (p2.Y - triangleBase.Y) * shaftWidth,
                    p1.Y + (triangleBase.X - p2.X) * shaftWidth);
            VertexCoordinates[6] = new Point(
                    triangleBase.X + (p2.Y - triangleBase.Y) * shaftWidth,
                    triangleBase.Y + (triangleBase.X - p2.X) * shaftWidth);

            pointCache[0] = ToPhysical(VertexCoordinates[0]);
            pointCache[1] = ToPhysical(VertexCoordinates[1]);
            pointCache[2] = ToPhysical(VertexCoordinates[2]);
            pointCache[3] = ToPhysical(VertexCoordinates[3]);
            pointCache[4] = ToPhysical(VertexCoordinates[4]);
            pointCache[5] = ToPhysical(VertexCoordinates[5]);
            pointCache[6] = ToPhysical(VertexCoordinates[6]);
        }
    }
}
