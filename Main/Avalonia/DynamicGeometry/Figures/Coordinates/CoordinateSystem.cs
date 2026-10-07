using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using M = System.Math;

namespace DynamicGeometry
{
    public partial class CoordinateSystem : IMovable
    {
        public CoordinateSystem(Drawing drawing)
        {
            Check.NotNull(drawing);

            Drawing = drawing;
            Drawing.SizeChanged += Drawing_SizeChanged;
            Origin = PhysicalSize.Scale(0.5).SnapToIntegers().Minus(0.5);
        }

        public Point PhysicalSize
        {
            get
            {
                return new Point(Canvas.ActualWidth, Canvas.ActualHeight);
            }
        }

        public Drawing Drawing { get; set; }

        /// <summary>
        /// Sets the unitLength and origin of the coordinate system so that the rectangle shows.
        /// A rectangle without a size (a file saved from a window that had none, a file
        /// edited by hand), or a canvas without one, leaves the view as it is: it threw, and
        /// the whole file was refused for the view it asked for.
        /// </summary>
        public void SetViewport(double minX, double maxX, double minY, double maxY)
        {
            var logicalWidth = maxX - minX;
            var logicalHeight = maxY - minY;
            var physicalSize = PhysicalSize;
            if (!logicalWidth.IsValidPositiveValue()
                || !logicalHeight.IsValidPositiveValue()
                || !physicalSize.X.IsValidPositiveValue()
                || !physicalSize.Y.IsValidPositiveValue())
            {
                return;
            }

            Fit(new Rect(minX, minY, logicalWidth, logicalHeight), marginPixels: 0);
        }

        public const double MinUnitLength = 1;
        public const double MaxUnitLength = 1000;
        const double zoomFactor = 1.2;

        /// <summary>What "zoom to fit" leaves free around the drawing</summary>
        public const double FitMarginPixels = 40;

        /// <summary>"Zoom to fit" on a tiny drawing (two points next to each other) stops here</summary>
        public const double MaxFitUnitLength = 200;

        /// <summary>
        /// The zoom within its limits; a zoom that is no number (worked out from a size of 0
        /// by 0) is the default one. Math.Min and Math.Max hand NaN on, and a NaN zoom put
        /// every figure nowhere, which layout throws on.
        /// </summary>
        public static double ClampUnitLength(double value)
        {
            if (double.IsNaN(value))
            {
                return Settings.DefaultUnitLength;
            }

            return M.Max(MinUnitLength, M.Min(MaxUnitLength, value));
        }

        /// <summary>
        /// Shows the logical rectangle as large as the canvas allows, centered.
        /// </summary>
        void Fit(Rect logicalBounds, double marginPixels, double maxUnitLength = MaxUnitLength, double zoom = 1)
        {
            var physicalSize = PhysicalSize;
            var availableWidth = M.Max(physicalSize.X - 2 * marginPixels, physicalSize.X / 2);
            var availableHeight = M.Max(physicalSize.Y - 2 * marginPixels, physicalSize.Y / 2);

            // a single point has no size: nothing to derive the zoom from, keep it
            var newUnitLength = unitLength;
            if (logicalBounds.Width > 0 || logicalBounds.Height > 0)
            {
                newUnitLength = M.Min(
                    logicalBounds.Width > 0 ? availableWidth / logicalBounds.Width : double.MaxValue,
                    logicalBounds.Height > 0 ? availableHeight / logicalBounds.Height : double.MaxValue);
                newUnitLength = M.Min(newUnitLength, maxUnitLength) * zoom;
            }

            SetView(logicalBounds.Center, ClampUnitLength(newUnitLength));
        }

        /// <summary>A suggested view of the drawing, edge to edge (see <see cref="Drawing.Scenes"/>)</summary>
        public void FitScene(Rect scene)
        {
            Fit(scene, marginPixels: 0);
        }

        /// <summary>
        /// Puts the logical point into the middle of the canvas at the given zoom.
        /// </summary>
        public void SetView(Point logicalCenter, double newUnitLength)
        {
            SetView(logicalCenter, newUnitLength, PhysicalSize.Scale(0.5));
        }

        /// <summary>
        /// Puts the logical point under the given pixel at the given zoom: the middle of the
        /// room that is left beside a caption, for instance.
        /// </summary>
        public void SetView(Point logicalPoint, double newUnitLength, Point physicalPoint)
        {
            unitLength = newUnitLength;
            scale = unitLength / Settings.DefaultUnitLength;
            origin = SnapOrigin(new Point(
                physicalPoint.X - logicalPoint.X * unitLength,
                physicalPoint.Y + logicalPoint.Y * unitLength));
            Recalculate();
        }

        // the origin sits in the middle of a pixel so that the axes are crisp
        static Point SnapOrigin(Point origin)
        {
            return origin.SnapToIntegers().Minus(0.5);
        }

        /// <summary>The logical point in the middle of the canvas</summary>
        public Point ViewCenter
        {
            get
            {
                return ToLogical(PhysicalSize.Scale(0.5));
            }
        }

        public void ZoomIn()
        {
            Zoom(zoomFactor, PhysicalSize.Scale(0.5));
        }

        public void ZoomOut()
        {
            Zoom(1 / zoomFactor, PhysicalSize.Scale(0.5));
        }

        /// <summary>
        /// Zooms so that whatever is under the focus stays under it: the cursor for the
        /// wheel, the middle of the canvas for the keyboard and the menu.
        /// </summary>
        /// <param name="focus">Physical (canvas) coordinates</param>
        public void Zoom(double factor, Point focus)
        {
            var newUnitLength = ClampUnitLength(unitLength * factor);
            if (newUnitLength == unitLength)
            {
                return;
            }

            var ratio = newUnitLength / unitLength;
            unitLength = newUnitLength;
            scale = unitLength / Settings.DefaultUnitLength;

            // not snapped: a rounded origin makes the point under the cursor creep while zooming
            origin = new Point(
                focus.X - (focus.X - origin.X) * ratio,
                focus.Y - (focus.Y - origin.Y) * ratio);
            Recalculate();
        }

        /// <summary>
        /// Two fingers on the paper: what was under <paramref name="from"/> goes under
        /// <paramref name="to"/> (the middle between the fingers, before and now), zoomed
        /// about it by how much the fingers came apart. At the limit of the zoom the paper
        /// still follows the fingers.
        /// </summary>
        /// <param name="from">Physical (canvas) coordinates</param>
        /// <param name="to">Physical (canvas) coordinates</param>
        public void PanAndZoom(Point from, Point to, double factor)
        {
            var newUnitLength = ClampUnitLength(unitLength * factor);
            var ratio = newUnitLength / unitLength;
            if (ratio == 1 && from == to)
            {
                return;
            }

            unitLength = newUnitLength;
            scale = unitLength / Settings.DefaultUnitLength;
            origin = new Point(
                to.X - (from.X - origin.X) * ratio,
                to.Y - (from.Y - origin.Y) * ratio);
            Recalculate();
        }

        /// <summary>
        /// Zoom to fit: everything visible, as large as possible. An empty drawing goes back
        /// to the default view.
        /// </summary>
        /// <param name="alsoShow">
        /// A part of the plane to keep in view whatever is in it: graphs and lines have no
        /// bounds of their own
        /// </param>
        /// <param name="zoom">A factor on the zoom that fits (a tile of the gallery is fitted a little closer)</param>
        public void ZoomExtend(Rect? alsoShow = null, double zoom = 1)
        {
            Rect bounds;
            if (!TryGetBoundsToShow(out bounds, alsoShow))
            {
                SetView(new Point(), Settings.DefaultUnitLength);
                return;
            }

            double margin = FitMarginPixels + GetPointReach();
            Fit(bounds, margin, MaxFitUnitLength, zoom);

            // text keeps its size in pixels, so in logical units a label grows as the view
            // zooms out: measure again at the new zoom until it settles. Only while it
            // changes: each fit works the whole drawing out again, and without labels the
            // bounds are the same at every zoom.
            for (int i = 0; i < 3 && TryGetBoundsToShow(out var refitted, alsoShow) && refitted != bounds; i++)
            {
                bounds = refitted;
                Fit(bounds, margin, MaxFitUnitLength, zoom);
            }
        }

        bool TryGetBoundsToShow(out Rect bounds, Rect? alsoShow)
        {
            bool hasContent = TryGetContentBounds(out bounds);
            if (alsoShow != null)
            {
                bounds = hasContent ? bounds.Union(alsoShow.Value) : alsoShow.Value;
                return true;
            }

            return hasContent;
        }

        /// <summary>
        /// Moves the middle of the drawing into the middle of the canvas, keeping the zoom.
        /// </summary>
        public void CenterContent()
        {
            Rect bounds;
            SetView(TryGetContentBounds(out bounds) ? bounds.Center : new Point(), unitLength);
        }

        /// <summary>
        /// The logical box around everything of a finite size that is showing: points, circles,
        /// arcs, labels. Lines, rays and graphs don't end, so they don't count.
        /// </summary>
        /// <param name="include">Which figures count; all of them by default</param>
        /// <param name="includeHidden">Which hidden figures count as if they showed; none by default</param>
        public bool TryGetContentBounds(
            out Rect bounds,
            Func<IFigure, bool> include = null,
            Func<IFigure, bool> includeHidden = null)
        {
            double minX = double.MaxValue;
            double minY = double.MaxValue;
            double maxX = double.MinValue;
            double maxY = double.MinValue;

            void Include(Point point)
            {
                if (!point.Exists())
                {
                    return;
                }

                minX = M.Min(minX, point.X);
                minY = M.Min(minY, point.Y);
                maxX = M.Max(maxX, point.X);
                maxY = M.Max(maxY, point.Y);
            }

            foreach (var figure in Drawing.Figures)
            {
                if ((!figure.Visible && (includeHidden == null || !includeHidden(figure)))
                    || !figure.Exists
                    || (include != null && !include(figure)))
                {
                    continue;
                }

                // a pinned label or box is on the screen, not in the plane: nothing to fit
                if (figure is Label { Pin: not LabelPin.None } || figure is ShowHideControl { Pin: not LabelPin.None })
                {
                    continue;
                }

                if (figure is IPoint point)
                {
                    Include(point.Coordinates);
                }
                else if (figure is Segment || figure is IPolygonalChain || figure is Bezier)
                {
                    // the corners count even when the points themselves are hidden
                    foreach (var vertex in figure.Dependencies.OfType<IPoint>())
                    {
                        Include(vertex.Coordinates);
                    }
                }
                else if (figure is BezierPath path)
                {
                    // the curve as drawn: it bulges past its anchors
                    var curve = path.Bounds;
                    if (curve.Width > 0 || curve.Height > 0)
                    {
                        Include(curve.TopLeft);
                        Include(curve.BottomRight);
                    }
                }
                else if (figure is Slider slider)
                {
                    Include(slider.Anchor.Coordinates);
                    Include(slider.Knob.Coordinates);
                }
                else if (figure is IArc arc && arc.SemiMajor.EqualsWithPrecision(arc.SemiMinor))
                {
                    // a circular arc reaches its ends and whichever of the four compass points
                    // lie on it - not the whole circle, which for the big arc of a spiral is
                    // many times the drawing
                    Include(arc.BeginLocation);
                    Include(arc.EndLocation);
                    var center = arc.Center;
                    double radius = arc.SemiMajor;
                    for (int quarter = 0; quarter < 4; quarter++)
                    {
                        double angle = quarter * M.PI / 2;
                        if (Math.IsAngleBetweenAngles(angle, arc.StartAngle, arc.EndAngle, arc.Clockwise))
                        {
                            Include(new Point(center.X + radius * M.Cos(angle), center.Y + radius * M.Sin(angle)));
                        }
                    }
                }
                else if (figure is IEllipse ellipse)
                {
                    // the box of the whole ellipse, turned or not, even for an arc of one
                    var reach = M.Max(ellipse.SemiMajor, ellipse.SemiMinor);
                    Include(ellipse.Center.Plus(new Point(reach, reach)));
                    Include(ellipse.Center.Minus(new Point(reach, reach)));
                }
                else if (figure is ControlBase control)
                {
                    // measured here and now: right after loading a drawing there was no layout
                    // pass yet, and the bounds of a label whose text just changed are stale
                    control.Shape.Measure(Size.Infinity);
                    var size = control.Shape.DesiredSize;
                    var topLeft = control.Coordinates;
                    Include(topLeft);
                    Include(new Point(topLeft.X + ToLogical(size.Width), topLeft.Y - ToLogical(size.Height)));
                }
            }

            if (minX > maxX)
            {
                bounds = default(Rect);
                return false;
            }

            bounds = new Rect(minX, minY, maxX - minX, maxY - minY);
            return true;
        }

        /// <summary>
        /// How far the biggest visible point reaches out from its place, in pixels: an emoji can
        /// be 100 px across, and fitting only its center would cut half of it off at the edge
        /// </summary>
        /// <param name="include">Which figures count; all of them by default</param>
        public double GetPointReach(Func<IFigure, bool> include = null)
        {
            double reach = 0;
            foreach (var figure in Drawing.Figures)
            {
                if (figure is PointBase point
                    && point.Visible
                    && point.Exists
                    && (include == null || include(figure))
                    && point.Shape.Width.IsValidValue())
                {
                    reach = M.Max(reach, point.Shape.Width / 2);
                }
            }

            return reach;
        }

        /// <summary>
        /// The zoom, at most <paramref name="unitLength"/>, at which every visible point, shape
        /// and all, stays inside a room of this size around <paramref name="center"/>. Each
        /// point limits it by its own size and place: a room shrunk on every side by the
        /// biggest point's reach (<see cref="GetPointReach"/>) left a figure with a big emoji
        /// in the middle much smaller than the room.
        /// </summary>
        /// <param name="include">Which figures count; all of them by default</param>
        public double LimitZoomByPointReach(
            double unitLength,
            Point center,
            Size room,
            Func<IFigure, bool> include = null)
        {
            foreach (var figure in Drawing.Figures)
            {
                if (figure is PointBase point
                    && point.Visible
                    && point.Exists
                    && (include == null || include(figure))
                    && point.Shape.Width.IsValidValue())
                {
                    double reach = point.Shape.Width / 2;
                    unitLength = Limit(unitLength, M.Abs(point.Coordinates.X - center.X), room.Width / 2 - reach);
                    unitLength = Limit(unitLength, M.Abs(point.Coordinates.Y - center.Y), room.Height / 2 - reach);
                }
            }

            return unitLength;

            static double Limit(double unitLength, double distance, double pixels)
            {
                return distance > 0 && pixels > 0 ? M.Min(unitLength, pixels / distance) : unitLength;
            }
        }

        //private double mScale = 1.0;
        //public double Scale
        //{
        //    get
        //    {
        //        return mScale;
        //    }
        //    set
        //    {
        //        mScale = value;
        //        UnitLength = value * Settings.DefaultUnitLength;
        //        Drawing.RaiseZoomChanged();
        //    }
        //}

        #region Bounds

        public Canvas Canvas
        {
            get
            {
                return Drawing.Canvas;
            }
        }

        public Point[] LogicalViewportVertices { get; set; }
        public double MinimalVisibleX { get; set; }
        public double MinimalVisibleY { get; set; }
        public double MaximalVisibleX { get; set; }
        public double MaximalVisibleY { get; set; }

        void Drawing_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            // what was in the middle of the canvas stays in the middle, whichever edge moved
            var previous = e.PreviousSize;
            if (previous.Width > 0 && previous.Height > 0)
            {
                origin = new Point(
                    origin.X + (e.NewSize.Width - previous.Width) / 2,
                    origin.Y + (e.NewSize.Height - previous.Height) / 2);
            }

            Recalculate();
        }

        private void Recalculate()
        {
            LogicalViewportVertices = GetViewportVerticesInLogical();
            if (LogicalViewportVertices != null)
            {
                MinimalVisibleX = LogicalViewportVertices.Min(p => p.X);
                MinimalVisibleY = LogicalViewportVertices.Min(p => p.Y);
                MaximalVisibleX = LogicalViewportVertices.Max(p => p.X);
                MaximalVisibleY = LogicalViewportVertices.Max(p => p.Y);
                Drawing.Recalculate();
            }
        }

        public Point[] GetViewportVerticesInLogical()
        {
            Point[] result = new Point[4];
            var physicalSize = PhysicalSize;
            if (!physicalSize.X.IsValidPositiveValue() || !physicalSize.Y.IsValidPositiveValue())
            {
                return LogicalViewportVertices;
            }
            result[0] = ToLogical(new Point());
            result[1] = ToLogical(new Point(physicalSize.X, 0));
            result[2] = ToLogical(new Point(physicalSize.X, physicalSize.Y));
            result[3] = ToLogical(new Point(0, physicalSize.Y));
            return result;
        }

        /// <summary>The x of every labeled grid line in view</summary>
        public IEnumerable<double> GetVisibleXPoints()
        {
            return GridValues(MinimalVisibleX, MaximalVisibleX, MajorGridStep);
        }

        /// <summary>The y of every labeled grid line in view</summary>
        public IEnumerable<double> GetVisibleYPoints()
        {
            return GridValues(MinimalVisibleY, MaximalVisibleY, MajorGridStep);
        }

        /// <summary>The x of every finer grid line in view, the labeled ones left out</summary>
        public IEnumerable<double> GetMinorXPoints()
        {
            return MinorGridValues(MinimalVisibleX, MaximalVisibleX);
        }

        /// <summary>The y of every finer grid line in view, the labeled ones left out</summary>
        public IEnumerable<double> GetMinorYPoints()
        {
            return MinorGridValues(MinimalVisibleY, MaximalVisibleY);
        }

        static IEnumerable<double> GridValues(double min, double max, double step)
        {
            long first = (long)M.Ceiling(min / step);
            long last = (long)M.Floor(max / step);
            for (long index = first; index <= last; index++)
            {
                yield return GridValue(index, step);
            }
        }

        IEnumerable<double> MinorGridValues(double min, double max)
        {
            ChooseGridStep(out double step, out int subdivisions);
            if (subdivisions == 1)
            {
                yield break;
            }

            step /= subdivisions;
            long first = (long)M.Ceiling(min / step);
            long last = (long)M.Floor(max / step);
            for (long index = first; index <= last; index++)
            {
                if (index % subdivisions != 0)
                {
                    yield return GridValue(index, step);
                }
            }
        }

        // index * step without the floating point dust: 3 * 0.2 is 0.6000000000000001,
        // and the label would print it
        static double GridValue(long index, double step)
        {
            return M.Round(index * step, 10);
        }

        #endregion

        #region Grid step

        /// <summary>Labeled grid lines come no closer than this, in pixels</summary>
        public const double MinimumMajorGridSpacing = 40;

        /// <summary>
        /// The finer lines between the labeled ones are left out when they would be closer
        /// than this, in pixels
        /// </summary>
        public const double MinimumMinorGridSpacing = 10;

        /// <summary>
        /// The grid never goes finer than this, in units, whatever the zoom: a drawing that
        /// counts unit squares (Pick's theorem) sets 1. Zero, the default, lets the zoom
        /// decide. Saved as GridStep on the Viewport.
        /// </summary>
        public double GridStep { get; set; }

        /// <summary>
        /// The distance between the labeled grid lines, in units: 1, 2 or 5 times a power of
        /// ten, the smallest that keeps them <see cref="MinimumMajorGridSpacing"/> apart at
        /// the current zoom. Shift-snapping lands on these lines.
        /// </summary>
        public double MajorGridStep
        {
            get
            {
                ChooseGridStep(out double step, out _);
                return step;
            }
        }

        /// <param name="subdivisions">
        /// How many parts the finer lines cut a step into: 5 (a step of 1 or 5), 4 (a step of
        /// 2), or 1 when there is no room for finer lines
        /// </param>
        void ChooseGridStep(out double step, out int subdivisions)
        {
            double minimum = MinimumMajorGridSpacing / unitLength;
            double decade = M.Pow(10, M.Floor(M.Log10(minimum)));
            if (decade >= minimum)
            {
                step = decade;
                subdivisions = 5;
            }
            else if (2 * decade >= minimum)
            {
                step = 2 * decade;
                subdivisions = 4;
            }
            else if (5 * decade >= minimum)
            {
                step = 5 * decade;
                subdivisions = 5;
            }
            else
            {
                step = 10 * decade;
                subdivisions = 5;
            }

            if (GridStep > 0 && step < GridStep)
            {
                step = GridStep;
            }

            double minorStep = step / subdivisions;
            if (minorStep * unitLength < MinimumMinorGridSpacing || (GridStep > 0 && minorStep < GridStep))
            {
                subdivisions = 1;
            }
        }

        #endregion

        #region Coordinate transforms

        private double scale = 1;
        public double Scale
        {
            get
            {
                return scale;
            }
        }

        private double unitLength = Settings.DefaultUnitLength;
        /// <summary>
        /// How many pixels are in a logical unit?
        /// </summary>
        public double UnitLength
        {
            get
            {
                return unitLength;
            }
            set
            {
                if (value < 1 || value > 1000)
                {
                    return;
                }
                unitLength = value;
                scale = unitLength / Settings.DefaultUnitLength;
                Recalculate();
            }
        }

        private Point origin;
        /// <summary>
        /// Origin is in physical coordinates (for 800x600 it will usually be (400;300))
        /// </summary>
        public Point Origin
        {
            get
            {
                return origin;
            }
            private set
            {
                origin = value;
                Recalculate();
            }
        }

        public double CursorTolerance
        {
            get
            {
                return ToLogical(Math.CursorTolerance);
            }
        }

        public virtual Point ToLogical(Point physicalPoint)
        {
            return new Point(
                 (physicalPoint.X - origin.X) / unitLength,
                -(physicalPoint.Y - origin.Y) / unitLength).RoundToEpsilon();
        }

        public IEnumerable<Point> ToLogical(IEnumerable<Point> physicalPoints)
        {
            return physicalPoints.Select(p => ToLogical(p));
        }

        public PointPair ToLogical(PointPair pointPair)
        {
            var result = new PointPair(ToLogical(pointPair.P1), ToLogical(pointPair.P2));
            if (result.P1.X > result.P2.X)
            {
                result = result.Reverse;
            }
            if (result.P1.Y > result.P2.Y)
            {
                var temp = result.P2.Y;
                result.P2 = result.P2.WithY(result.P1.Y);
                result.P1 = result.P1.WithY(temp);
            }
            return result;
        }

        public virtual double ToLogical(double length)
        {
            return length / UnitLength;
        }

        public virtual Avalonia.Rect ToLogical(Avalonia.Rect rect)
        {
            var result = new Avalonia.Rect();
            var origin = new Point(rect.X, rect.Y);
            var logicalOrigin = ToLogical(origin);
            result = result.WithX(logicalOrigin.X);
            result = result.WithY(logicalOrigin.Y);
            result = result.WithWidth(ToLogical(rect.Width));
            result = result.WithHeight(ToLogical(rect.Height));
            return result;
        }

        public virtual Avalonia.Rect ToPhysical(Avalonia.Rect rect)
        {
            var result = new Avalonia.Rect();
            var origin = new Point(rect.X, rect.Y);
            var physicalOrigin = ToPhysical(origin);
            result = result.WithX(physicalOrigin.X);
            result = result.WithY(physicalOrigin.Y);
            result = result.WithWidth(ToPhysical(rect.Width));
            result = result.WithHeight(ToPhysical(rect.Height));
            return result;
        }

        public virtual Point ToPhysical(Point logicalPoint)
        {
            return new Point(
                origin.X + logicalPoint.X * unitLength,
                origin.Y - logicalPoint.Y * unitLength);
        }

        public PointPair ToPhysical(PointPair logicalPointPair)
        {
            return new PointPair(ToPhysical(logicalPointPair.P1), ToPhysical(logicalPointPair.P2));
        }

        public IEnumerable<Point> ToPhysical(IEnumerable<Point> logicalPoints)
        {
            return logicalPoints.Select(p => ToPhysical(p));
        }

        public void ToPhysicalInPlace(List<Point> logicalPoints)
        {
            for (int i = 0; i < logicalPoints.Count; i++)
            {
                logicalPoints[i] = ToPhysical(logicalPoints[i]);
            }
        }

        public virtual double ToPhysical(double length)
        {
            return length * UnitLength;
        }

        #endregion

        #region IMovable Members

        public void MoveTo(Point position)
        {
            // Exactly, no rounding to whole pixels: a drag arrives as many small steps (half a
            // pixel each on a 200% display) and every one of them was rounded up to a full
            // pixel, so the plane ran ahead of the cursor - twice as fast on a slow drag.
            position = ToPhysical(position);
            if (position == origin)
            {
                return;
            }

            Origin = position;
        }

        public bool AllowMove()
        {
            return true;
        }

        public Point Coordinates
        {
            get { return new Point(); }
        }

        #endregion
    }
}
