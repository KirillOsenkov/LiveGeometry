using System.Collections.Generic;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using AvaloniaShapes = Avalonia.Controls.Shapes;

namespace LiveGeometry;

/// <summary>
/// What a drag selects with (<see cref="RegionSelection"/>): a rectangle, or a lasso - the path
/// the pointer went, closed by a straight line back to where it began. In canvas pixels, and
/// judged by what is on the screen: a figure is taken by the shapes that draw it, the ones a
/// selection's halo is drawn from (<see cref="SelectionHalo"/>), so nothing is needed per kind
/// of figure.
/// </summary>
public class SelectionRegion
{
    /// <summary>In pixels: how far the pointer goes before the lasso takes in another point of its path</summary>
    public static double LassoStep = 2;

    /// <summary>
    /// In pixels: of a figure that goes on past the window or under the side panel, what lies
    /// this far in from the window's edges and from the panel must be inside the region (see
    /// <see cref="IsInside"/>). A drag right along the edge is too much to ask.
    /// </summary>
    public static double EdgeMargin = 50;

    readonly List<Point> points = new List<Point>();

    // the way a lasso takes figures, decided by the way it first goes; null until then
    bool? lassoTakesTouched;

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

    /// <summary>
    /// Whether the region takes every figure it touches (true), or only those wholly inside it
    /// (false), as AutoCAD's crossing and window selections go: a rectangle dragged to the left
    /// of where it began touches, one dragged to the right holds; a lasso goes by the way it
    /// first went.
    /// </summary>
    public bool TakesTouchedFigures => IsRectangle
        ? points.Count > 1 && points[1].X < points[0].X
        : lassoTakesTouched == true;

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

        bool isDecided = false;
        if (lassoTakesTouched == null && System.Math.Abs(point.X - points[0].X) >= Dragger.DragThreshold)
        {
            lassoTakesTouched = point.X < points[0].X;
            isDecided = true;
        }

        if (points[points.Count - 1].Distance(point) < LassoStep)
        {
            return isDecided;
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
        return CreatePolygon(Outline, transform);
    }

    static Geometry CreatePolygon(IEnumerable<Point> corners, Matrix? transform)
    {
        var geometry = new StreamGeometry();
        using (var context = geometry.Open())
        {
            context.SetFillRule(FillRule.NonZero);
            bool isFirst = true;
            foreach (var point in corners)
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
        var area = new VisibleArea(canvas);
        var taken = new HashSet<IFigure>();
        foreach (var figure in drawing.Figures)
        {
            if (CanBeTaken(drawing, figure)
                && ((!TakesTouchedFigures && IsTakenWithItsPoints(figure, taken))
                    || Takes(figure, canvas, bounds, area)))
            {
                result.Add(figure);
                taken.Add(figure);
            }
        }

        return result;
    }

    /// <summary>
    /// Not a hidden figure, nor what a click can't select: the grid and its axes, a name (it
    /// goes with its figure, as Delete takes it), the text a gallery drawing came with (paper
    /// to the Drag tool)
    /// </summary>
    static bool CanBeTaken(Drawing drawing, IFigure figure)
    {
        return figure.Visible
            && figure.Exists
            && figure.IsHitTestVisible
            && !(figure is CartesianGrid)
            && !(figure is AxisLine)
            && !(figure is NameLabel)
            && !(figure is ControlBase label && drawing.FixedLabels.Contains(label));
    }

    /// <summary>A figure that goes on past the window, which no region holds whole</summary>
    static bool IsEndless(IFigure figure)
    {
        return (figure is LineBase && !(figure is Segment)) || figure is FunctionGraph;
    }

    /// <summary>
    /// A line or a ray goes with the points it is drawn through when the region takes all of
    /// them: line AB with A and B, a parallel with the point it goes through. (The drawing's
    /// list is in dependency order, so its points have been judged already.)
    /// </summary>
    static bool IsTakenWithItsPoints(IFigure figure, HashSet<IFigure> taken)
    {
        if (!IsEndless(figure))
        {
            return false;
        }

        bool hasPoints = false;
        foreach (var dependency in figure.Dependencies)
        {
            if (dependency is IPoint)
            {
                if (!taken.Contains(dependency))
                {
                    return false;
                }

                hasPoints = true;
            }
        }

        return hasPoints;
    }

    bool Takes(IFigure figure, Canvas canvas, Rect bounds, VisibleArea area)
    {
        var shapes = new List<Control>();
        CollectShapes(figure, shapes);
        bool isEndless = IsEndless(figure);
        bool isShown = false;
        foreach (var shape in shapes)
        {
            var drawn = DrawnShape.Find(shape, canvas);
            if (drawn == null)
            {
                continue;
            }

            if (TakesTouchedFigures)
            {
                if (Touches(drawn, bounds))
                {
                    return true;
                }

                continue;
            }

            var inside = IsInside(drawn, bounds, area, isEndless);
            if (inside == null)
            {
                continue;
            }

            if (inside == false)
            {
                return false;
            }

            isShown = true;
        }

        return isShown;
    }

    /// <summary>
    /// Where the figure is drawn on the canvas, in pixels, the stroke included; null for one
    /// not drawn, or one that goes on past the window (a line has no place to be shown whole)
    /// </summary>
    public static Rect? FindShownBounds(IFigure figure, Canvas canvas)
    {
        if (IsEndless(figure) || !figure.Visible || !figure.Exists)
        {
            return null;
        }

        var shapes = new List<Control>();
        CollectShapes(figure, shapes);
        Rect? bounds = null;
        foreach (var shape in shapes)
        {
            var drawn = DrawnShape.Find(shape, canvas);
            if (drawn != null)
            {
                bounds = bounds?.Union(drawn.Bounds) ?? drawn.Bounds;
            }
        }

        return bounds;
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

    bool Touches(DrawnShape drawn, Rect regionBounds)
    {
        if (!drawn.Bounds.Intersects(regionBounds))
        {
            return false;
        }

        if (IsRectangle && drawn.IsAxisAligned && regionBounds.Contains(drawn.Bounds))
        {
            return true;
        }

        // the region in the shape's own coordinates, where its geometry is
        var region = CreateGeometry(drawn.Inverse);
        drawn.CreateGeometries();
        return Intersects(region, drawn.Fill) || Intersects(region, drawn.Stroke);
    }

    /// <summary>
    /// Whether the shape is inside the region: the whole of it where the window shows it whole,
    /// else what of it lies <see cref="EdgeMargin"/> in from the window's edges and the side
    /// panel - or, where nothing of it does, what shows at all. A line is never shown whole.
    /// Null for a shape not on the screen.
    /// </summary>
    bool? IsInside(DrawnShape drawn, Rect regionBounds, VisibleArea area, bool isEndless)
    {
        var shown = drawn.Bounds.Intersect(area.Window);
        if (shown.Width <= 0 && shown.Height <= 0)
        {
            return null;
        }

        if (IsRectangle && drawn.IsAxisAligned && regionBounds.Contains(drawn.Bounds))
        {
            return true;
        }

        if (!shown.Intersects(regionBounds))
        {
            return false;
        }

        var region = CreateGeometry(drawn.Inverse);
        drawn.CreateGeometries();
        if (!isEndless && area.ShowsWhole(drawn.Bounds))
        {
            return IsWithin(drawn.Fill, region) && IsWithin(drawn.Stroke, region);
        }

        var reach = area.CreateGeometry(EdgeMargin, drawn.Inverse);
        var fill = Clip(drawn.Fill, reach);
        var stroke = Clip(drawn.Stroke, reach);
        if (fill == null && stroke == null)
        {
            var screen = area.CreateGeometry(margin: 0, drawn.Inverse);
            fill = Clip(drawn.Fill, screen);
            stroke = Clip(drawn.Stroke, screen);
            if (fill == null && stroke == null)
            {
                return null;
            }
        }

        return IsWithin(fill, region) && IsWithin(stroke, region);
    }

    /// <summary>What of the geometry lies in the area, or null for nothing</summary>
    static Geometry Clip(Geometry geometry, Geometry area)
    {
        if (geometry == null)
        {
            return null;
        }

        var clipped = new CombinedGeometry(GeometryCombineMode.Intersect, geometry, area);
        return IsEmpty(clipped.Bounds) ? null : clipped;
    }

    static bool IsPainted(IBrush brush)
    {
        return brush != null
            && brush.Opacity > 0
            && !(brush is ISolidColorBrush solid && solid.Color.A == 0);
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

    /// <summary>
    /// A shape as it is drawn on the canvas: its fill and its stroke, or the box of a control
    /// without a geometry (a label, a show/hide box, an emoji). The geometries are made only
    /// when the bounds don't decide.
    /// </summary>
    class DrawnShape
    {
        Geometry geometry;
        Pen pen;
        bool isFilled;
        Rect box;
        bool hasGeometries;

        /// <summary>From the shape's coordinates to the canvas</summary>
        public Matrix Matrix { get; private set; }

        /// <summary>From the canvas to the shape's coordinates</summary>
        public Matrix Inverse { get; private set; }

        /// <summary>On the canvas, the stroke included</summary>
        public Rect Bounds { get; private set; }

        public Geometry Fill { get; private set; }

        public Geometry Stroke { get; private set; }

        /// <summary>Whether the matrix only moves and scales, so that a box stays a box</summary>
        public bool IsAxisAligned => Matrix.M12 == 0 && Matrix.M21 == 0;

        /// <summary>The shape as drawn, or null for one not on the screen or that paints nothing</summary>
        public static DrawnShape Find(Control shape, Canvas canvas)
        {
            if (!shape.IsEffectivelyVisible)
            {
                return null;
            }

            var transform = shape.TransformToVisual(canvas);
            if (transform == null || !transform.Value.TryInvert(out var inverse))
            {
                return null;
            }

            var drawn = new DrawnShape() { Matrix = transform.Value, Inverse = inverse };
            if (shape is AvaloniaShapes.Shape geometryShape && geometryShape.RenderedGeometry != null)
            {
                var geometry = geometryShape.RenderedGeometry;
                bool isStroked = IsPainted(geometryShape.Stroke) && geometryShape.StrokeThickness > 0;
                drawn.isFilled = IsPainted(geometryShape.Fill);
                if (!isStroked && !drawn.isFilled)
                {
                    return null;
                }

                drawn.geometry = geometry;
                if (isStroked)
                {
                    drawn.pen = new Pen(
                        Brushes.Black,
                        geometryShape.StrokeThickness,
                        lineCap: geometryShape.StrokeLineCap,
                        lineJoin: geometryShape.StrokeJoin);
                }

                drawn.Bounds = (drawn.pen != null ? geometry.GetRenderBounds(drawn.pen) : geometry.Bounds)
                    .TransformToAABB(drawn.Matrix);
            }
            else
            {
                drawn.box = new Rect(shape.Bounds.Size);
                drawn.Bounds = drawn.box.TransformToAABB(drawn.Matrix);
            }

            return drawn;
        }

        public void CreateGeometries()
        {
            if (hasGeometries)
            {
                return;
            }

            hasGeometries = true;
            if (geometry == null)
            {
                Fill = new RectangleGeometry(box);
                return;
            }

            // (copies: a geometry combined with another is watched by the combination, and
            // the shape's own would be kept alive by every one made)
            if (isFilled)
            {
                Fill = geometry.Clone();
            }

            if (pen != null)
            {
                Stroke = geometry.GetWidenedGeometry(pen);
            }
        }
    }

    /// <summary>What of the canvas shows: the window, less the side panel where it lies over the canvas</summary>
    class VisibleArea
    {
        public VisibleArea(Canvas canvas)
        {
            Window = new Rect(canvas.Bounds.Size);
            Panel = DrawingHost.FindSidePanelBounds(canvas);
        }

        public Rect Window { get; }

        public Rect? Panel { get; }

        public bool ShowsWhole(Rect bounds)
        {
            return Window.Contains(bounds) && !(Panel != null && Panel.Value.Intersects(bounds));
        }

        /// <summary>
        /// The area <paramref name="margin"/> pixels in from the window's edges and out from the
        /// panel, in the coordinates <paramref name="transform"/> takes canvas pixels to
        /// </summary>
        public Geometry CreateGeometry(double margin, Matrix transform)
        {
            double inset = System.Math.Min(margin, System.Math.Min(Window.Width, Window.Height) / 2);
            var inner = Window.Deflate(inset);
            Geometry area = CreatePolygon(Corners(inner), transform);
            if (Panel != null)
            {
                area = new CombinedGeometry(
                    GeometryCombineMode.Exclude,
                    area,
                    CreatePolygon(Corners(Panel.Value.Inflate(margin)), transform));
            }

            return area;
        }

        static Point[] Corners(Rect rect)
        {
            return [rect.TopLeft, rect.TopRight, rect.BottomRight, rect.BottomLeft];
        }
    }
}
