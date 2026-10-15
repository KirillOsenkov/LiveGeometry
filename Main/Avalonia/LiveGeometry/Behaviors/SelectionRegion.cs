using System.Collections.Generic;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using AvaloniaShapes = Avalonia.Controls.Shapes;

namespace DynamicGeometry;

/// <summary>
/// What a drag selects with (<see cref="RegionSelection"/>): a rectangle, or a lasso - the path
/// the pointer went, closed by a straight line back to where it began. In canvas pixels, and
/// judged by what is on the screen: a figure is taken by the shapes that draw it, the ones a
/// selection's halo is drawn from (<see cref="SelectionHalo"/>), so nothing is needed per kind
/// of figure.
/// </summary>
public class SelectionRegion
{
    /// <summary>
    /// Whether the region takes every figure it touches (true), or only those wholly inside it
    /// (false). Lines, rays and graphs of functions go on past the window, and are taken when
    /// touched either way: what they are built on may be far away.
    /// </summary>
    public const bool TakesTouchedFigures = true;

    /// <summary>In pixels: how far the pointer goes before the lasso takes in another point of its path</summary>
    public static double LassoStep = 2;

    readonly List<Point> points = new List<Point>();

    SelectionRegion(bool isRectangle, Point start)
    {
        IsRectangle = isRectangle;
        points.Add(start);
    }

    public static SelectionRegion Rectangle(Point corner)
    {
        return new SelectionRegion(isRectangle: true, corner);
    }

    public static SelectionRegion Lasso(Point start)
    {
        return new SelectionRegion(isRectangle: false, start);
    }

    public bool IsRectangle { get; }

    /// <summary>The rectangle's corner across from the first, or the lasso's next point; false when the region stays as it was</summary>
    public bool Extend(Point point)
    {
        if (IsRectangle)
        {
            if (points.Count == 1)
            {
                points.Add(point);
                return true;
            }

            if (points[1] == point)
            {
                return false;
            }

            points[1] = point;
            return true;
        }

        if (points[points.Count - 1].Distance(point) < LassoStep)
        {
            return false;
        }

        points.Add(point);
        return true;
    }

    /// <summary>The corners of the region in order round it, the last joined to the first</summary>
    public IReadOnlyList<Point> Outline
    {
        get
        {
            if (!IsRectangle)
            {
                return points;
            }

            var first = points[0];
            var second = points[points.Count - 1];
            return new[] { first, new Point(second.X, first.Y), second, new Point(first.X, second.Y) };
        }
    }

    public Rect Bounds
    {
        get
        {
            var outline = Outline;
            double left = outline[0].X;
            double top = outline[0].Y;
            double right = left;
            double bottom = top;
            foreach (var point in outline)
            {
                left = System.Math.Min(left, point.X);
                top = System.Math.Min(top, point.Y);
                right = System.Math.Max(right, point.X);
                bottom = System.Math.Max(bottom, point.Y);
            }

            return new Rect(left, top, right - left, bottom - top);
        }
    }

    /// <summary>
    /// The outline as a closed geometry, filled by the nonzero rule: a lasso drawn round twice,
    /// or into a figure eight, holds all that it goes round
    /// </summary>
    /// <param name="transform">From canvas pixels to the coordinates wanted, or null for canvas pixels</param>
    public Geometry CreateGeometry(Matrix? transform = null)
    {
        var geometry = new StreamGeometry();
        using (var context = geometry.Open())
        {
            context.SetFillRule(FillRule.NonZero);
            bool isFirst = true;
            foreach (var point in Outline)
            {
                var transformed = transform != null ? point.Transform(transform.Value) : point;
                if (isFirst)
                {
                    context.BeginFigure(transformed, isFilled: true);
                    isFirst = false;
                }
                else
                {
                    context.LineTo(transformed);
                }
            }

            context.EndFigure(isClosed: true);
        }

        return geometry;
    }

    /// <summary>The figures of the drawing the region takes, each a whole: a regular polygon, not its side</summary>
    public List<IFigure> FindFigures(Drawing drawing)
    {
        var result = new List<IFigure>();
        var canvas = drawing.Canvas;
        if (canvas == null)
        {
            return result;
        }

        var bounds = Bounds;
        foreach (var figure in drawing.Figures)
        {
            if (CanBeTaken(drawing, figure) && Takes(figure, canvas, bounds))
            {
                result.Add(figure);
            }
        }

        return result;
    }

    /// <summary>
    /// Not a hidden figure, nor what a click can't select: the grid and its axes, a Number (it
    /// has no place), a name (it goes with its figure, as Delete takes it), the text a gallery
    /// drawing came with (paper to the Drag tool)
    /// </summary>
    static bool CanBeTaken(Drawing drawing, IFigure figure)
    {
        return figure.Visible
            && figure.Exists
            && figure.IsHitTestVisible
            && !(figure is CartesianGrid)
            && !(figure is AxisLine)
            && !(figure is Number)
            && !(figure is PointLabel)
            && !(figure is FigureLabel)
            && !(figure is ControlBase label && drawing.FixedLabels.Contains(label));
    }

    /// <summary>A figure that goes on past the window, which no region holds whole</summary>
    static bool IsEndless(IFigure figure)
    {
        return (figure is LineBase && !(figure is Segment)) || figure is FunctionGraph;
    }

    bool Takes(IFigure figure, Canvas canvas, Rect bounds)
    {
        var shapes = new List<Control>();
        CollectShapes(figure, shapes);
        bool takesTouched = TakesTouchedFigures || IsEndless(figure);
        bool isShown = false;
        foreach (var shape in shapes)
        {
            var overlap = FindOverlap(shape, canvas, bounds, needsInside: !takesTouched);
            if (overlap == null)
            {
                continue;
            }

            isShown = true;
            if (takesTouched && overlap != Overlap.Apart)
            {
                return true;
            }

            if (!takesTouched && overlap != Overlap.Inside)
            {
                return false;
            }
        }

        return !takesTouched && isShown;
    }

    /// <summary>The controls that draw the figure: its shape, or the shapes of its parts (a slider, a vector, a regular polygon, a path)</summary>
    static void CollectShapes(IFigure figure, List<Control> shapes)
    {
        if (figure is IHasShape owner && owner.Shape != null)
        {
            shapes.Add(owner.Shape);
        }

        if (figure is CompositeFigure composite)
        {
            foreach (var child in composite.Children)
            {
                // (a handle shows only next to an anchor that is selected: it says nothing of the path)
                if (!(child is BezierPath.BezierPathHandle))
                {
                    CollectShapes(child, shapes);
                }
            }
        }
    }

    enum Overlap
    {
        Apart,
        Touching,
        Inside
    }

    /// <summary>
    /// How the shape lies to the region: its fill and its stroke as drawn, or the box of a
    /// control without a geometry (a label, a show/hide box, an emoji); null for a shape not on
    /// the screen. Inside is only told apart from touching when <paramref name="needsInside"/>.
    /// </summary>
    Overlap? FindOverlap(Control shape, Canvas canvas, Rect regionBounds, bool needsInside)
    {
        if (!shape.IsEffectivelyVisible)
        {
            return null;
        }

        var transform = shape.TransformToVisual(canvas);
        if (transform == null)
        {
            return null;
        }

        var matrix = transform.Value;
        Geometry fill = null;
        Geometry stroke = null;
        Rect bounds;
        if (shape is AvaloniaShapes.Shape geometryShape && geometryShape.RenderedGeometry != null)
        {
            var geometry = geometryShape.RenderedGeometry;
            bool isStroked = IsPainted(geometryShape.Stroke) && geometryShape.StrokeThickness > 0;
            bool isFilled = IsPainted(geometryShape.Fill);
            if (!isStroked && !isFilled)
            {
                return null;
            }

            Pen pen = null;
            if (isStroked)
            {
                pen = new Pen(
                    Brushes.Black,
                    geometryShape.StrokeThickness,
                    lineCap: geometryShape.StrokeLineCap,
                    lineJoin: geometryShape.StrokeJoin);
            }

            bounds = (pen != null ? geometry.GetRenderBounds(pen) : geometry.Bounds).TransformToAABB(matrix);
            if (!bounds.Intersects(regionBounds))
            {
                return Overlap.Apart;
            }

            if (IsRectangle && IsAxisAligned(matrix) && regionBounds.Contains(bounds))
            {
                return Overlap.Inside;
            }

            // (copies: a geometry combined with another is watched by the combination, and
            // the shape's own would be kept alive by every one made)
            if (isFilled)
            {
                fill = geometry.Clone();
            }

            if (pen != null)
            {
                stroke = geometry.GetWidenedGeometry(pen);
            }
        }
        else
        {
            var box = new Rect(shape.Bounds.Size);
            bounds = box.TransformToAABB(matrix);
            if (!bounds.Intersects(regionBounds))
            {
                return Overlap.Apart;
            }

            if (IsRectangle && IsAxisAligned(matrix) && regionBounds.Contains(bounds))
            {
                return Overlap.Inside;
            }

            fill = new RectangleGeometry(box);
        }

        if (!matrix.TryInvert(out var inverse))
        {
            return null;
        }

        // the region in the shape's own coordinates, where its geometry is
        var region = CreateGeometry(inverse);
        bool touches = Intersects(region, fill) || Intersects(region, stroke);
        if (!touches)
        {
            return Overlap.Apart;
        }

        if (needsInside && IsWithin(fill, region) && IsWithin(stroke, region))
        {
            return Overlap.Inside;
        }

        return Overlap.Touching;
    }

    static bool IsPainted(IBrush brush)
    {
        return brush != null
            && brush.Opacity > 0
            && !(brush is ISolidColorBrush solid && solid.Color.A == 0);
    }

    /// <summary>Whether the matrix only moves and scales, so that a box stays a box</summary>
    static bool IsAxisAligned(Matrix matrix)
    {
        return matrix.M12 == 0 && matrix.M21 == 0;
    }

    static bool Intersects(Geometry region, Geometry geometry)
    {
        return geometry != null
            && !IsEmpty(new CombinedGeometry(GeometryCombineMode.Intersect, region, geometry).Bounds);
    }

    static bool IsWithin(Geometry geometry, Geometry region)
    {
        return geometry == null
            || IsEmpty(new CombinedGeometry(GeometryCombineMode.Exclude, geometry, region).Bounds);
    }

    static bool IsEmpty(Rect bounds)
    {
        return bounds.Width <= 0 || bounds.Height <= 0;
    }
}
