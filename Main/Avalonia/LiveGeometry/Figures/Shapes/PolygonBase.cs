using System.Linq;
using Avalonia;

namespace LiveGeometry;

public interface IPolygonalChain
{
    Point[] VertexCoordinates { get; }
}

public interface IPolygon : IPolygonalChain { }

public abstract class PolygonBase : ShapeBase<Avalonia.Controls.Shapes.Polygon>, IPolygonalChain, IShapeWithInterior, IPerimeter, IRegion
{
    /// <summary>
    /// Just for caching purposes, to avoid array allocations on a hotpath
    /// </summary>
    protected IPoint[] vertices;
    protected Point[] vertexCoordinates;

    /// <summary>Polygon p, q... when no points name it (more than ten vertices)</summary>
    protected override string FirstLetter => "p";
    public Point[] VertexCoordinates => vertexCoordinates;

    protected override void OnDependenciesChanged()
    {
        UpdatePointCache();
    }

    protected void UpdatePointCache()
    {
        vertices = Dependencies.Where(f => f is IPoint).Cast<IPoint>().ToArray();   // Tolerates non-IPoint dependencies.
        vertexCoordinates = new Point[vertices.Length];
        var cache = new PointCollection();
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

        Shape.PointsChanged();
    }

    protected override ZOrder DefaultLayer()
    {
        return ZOrder.Polygons;
    }

    /// <summary>
    /// The outline where it draws something, the inside where it is filled: on a line
    /// style, or a shape style not filled, a polygon is an outline, and a click inside it
    /// goes through, as inside a circle
    /// </summary>
    public override IFigure HitTest(Point point)
    {
        if (DrawsOutline
            && Math.IsPointOnPolygonalChain(vertexCoordinates, point, ToLogical(Shape.StrokeThickness / 2 + Math.CursorTolerance), true))
        {
            return this;
        }

        bool isInside = Shape.Fill != null && vertexCoordinates.IsPointInPolygon(point);
        return isInside ? this : null;
    }

    bool DrawsOutline
    {
        get
        {
            return Shape.StrokeThickness > 0
                && Shape.Stroke != null
                && !(Shape.Stroke is Avalonia.Media.ISolidColorBrush brush && brush.Color.A == 0);
        }
    }

    public double Area => VertexCoordinates.Area();

    public bool HasInside => vertices != null && vertices.Length >= 3;

    /// <summary>The vertices as they are right now (the cache of <see cref="VertexCoordinates"/> is the last visual's)</summary>
    Point[] CurrentVertexCoordinates()
    {
        if (vertices == null)
        {
            UpdatePointCache();
        }

        var result = new Point[vertices.Length];
        for (int i = 0; i < vertices.Length; i++)
        {
            result[i] = vertices[i].Coordinates;
        }

        return result;
    }

    public bool ContainsPoint(Point point)
    {
        return HasInside && CurrentVertexCoordinates().IsPointInPolygon(point);
    }

    public Point NearestPointInside(Point point)
    {
        if (!HasInside)
        {
            return Math.InfinitePoint;
        }

        var current = CurrentVertexCoordinates();
        return current.IsPointInPolygon(point) ? point : Math.GetNearestPointOnPolygonalChain(current, point, isClosed: true);
    }

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

    /// <summary>No fill of its own: a shape style fills it, a line style leaves it an outline</summary>
    protected override Avalonia.Controls.Shapes.Polygon CreateShape()
    {
        return new PolygonShape();
    }
}
