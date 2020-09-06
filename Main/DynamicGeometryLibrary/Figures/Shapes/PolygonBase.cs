using System.Linq;
using System.Windows;

namespace DynamicGeometry
{
    public interface IPolygonalChain
    {
        Point[] VertexCoordinates { get; }
    }

    public interface IPolygon : IPolygonalChain { }

    public abstract class PolygonBase : ShapeBase<System.Windows.Shapes.Polygon>, IPolygonalChain, IShapeWithInterior
    {
        /// <summary>
        /// Just for caching purposes, to avoid array allocations on a hotpath
        /// </summary>
        protected IPoint[] vertices;
        protected Point[] vertexCoordinates;
        public Point[] VertexCoordinates
        {
            get
            {
                return vertexCoordinates;
            }
        }

        protected override void OnDependenciesChanged()
        {
            UpdatePointCache();
        }

        protected void UpdatePointCache()
        {
            vertices = Dependencies.Where(f => f is IPoint).Cast<IPoint>().ToArray();   // Tolerates non-IPoint dependencies.
            vertexCoordinates = new Point[vertices.Length];
            var cache = new System.Windows.Media.PointCollection();
            Shape.Points = cache;

            for (int i = 0; i < vertices.Length; i++)
            {
                cache.Add(new Point());
            }
        }

        public override void UpdateVisual()
        {
            if (vertices == null)
            {
                UpdatePointCache();
            }

            var points = Shape.Points;
            var coordinateSystem = Drawing.CoordinateSystem;
            for (int i = 0; i < vertices.Length; i++)
            {
                vertexCoordinates[i] = vertices[i].Coordinates;
                points[i] = coordinateSystem.ToPhysical(vertexCoordinates[i]);
            }
        }

        protected override int DefaultZOrder()
        {
            return (int)ZOrder.Polygons;
        }

        public override IFigure HitTest(Point point)
        {
            bool isInside = vertexCoordinates.IsPointInPolygon(point);
            return isInside ? this : null;
        }

        public double Area => VertexCoordinates.Area();

        public double Perimeter
        {
            get
            {
                double sum = 0;
                int vertexCount = VertexCoordinates.Count();
                for (int i = 0; i < vertexCount; i++)
                {
                    sum += VertexCoordinates[i].Distance(VertexCoordinates[i.RotateNext(vertexCount)]);
                }
                return sum;
            }
        }

        public override Point Center => VertexCoordinates.Midpoint();

        protected override System.Windows.Shapes.Polygon CreateShape()
        {
            return Factory.CreatePolygonShape();
        }
    }
}
