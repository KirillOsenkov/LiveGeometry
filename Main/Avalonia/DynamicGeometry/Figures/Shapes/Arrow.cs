using System.Collections.Generic;
using System.Linq;
using Avalonia;

namespace DynamicGeometry
{
    public class Arrow : Polygon
    {
        public Arrow()
        {
            Shape.Stroke = null;
            Shape.Points = pointCache;
        }

        protected override void OnDependenciesChanged()
        {
        }

        PointCollection pointCache = new PointCollection();

        /// <summary>
        /// Whether the arrow is the whole of it, shaft and head, in one solid color (an axis),
        /// or only the head, and the shaft is a line of its own that a dash can break (a
        /// vector: <see cref="Vector.VectorShaft"/>)
        /// </summary>
        public bool DrawsShaft { get; set; } = true;

        /// <summary>Where the arrow is, in pixels</summary>
        public struct Outline
        {
            public Point Tail;
            public Point HeadBase;
            public Point Tip;
            public Point Across;
            public double HalfShaft;
            public double HalfHead;
        }

        /// <summary>
        /// All in pixels: an arrow is a line with a head, as wide as its style says and with a
        /// head that goes with that width - the same at any zoom.
        /// </summary>
        public Outline Measure()
        {
            PointPair line = this.Line(0);
            LineBase parentLine = Dependencies.ElementAt(0) as LineBase;
            if (parentLine != null)
            {
                line = parentLine.OnScreenCoordinates;
            }

            Point tail = ToPhysical(line.P1);
            Point tip = ToPhysical(line.P2);
            double length = tail.Distance(tip);

            var lineStyle = Style as LineStyle;
            double width = lineStyle != null ? lineStyle.StrokeWidth : 1;
            double headLength = System.Math.Min(HeadLength + HeadGrowth * width, length);

            var along = length > 1e-9 ? (tip - tail) / length : new Point(1, 0);

            // the head points at the end point, it doesn't hide under it
            var endPoint = parentLine != null && parentLine.Dependencies.Count > 1
                ? parentLine.Dependencies[1] as PointBase
                : null;
            if (endPoint != null && endPoint.Visible && endPoint.Shape != null)
            {
                double pointRadius = endPoint.Shape.Width / 2;
                if (pointRadius > 0 && length > pointRadius + headLength)
                {
                    tip -= along * pointRadius;
                }
            }

            return new Outline()
            {
                Tail = tail,
                HeadBase = tip - along * headLength,
                Tip = tip,
                Across = new Point(-along.Y, along.X),
                HalfShaft = System.Math.Max(width / 2, 0.5),
                HalfHead = HeadHalfWidth + HeadGrowth / 2 * width
            };
        }

        public override void UpdateVisual()
        {
            if (Drawing == null)
            {
                return;
            }

            var outline = Measure();
            var headBase = outline.HeadBase;
            var across = outline.Across;
            var points = new List<Point>()
            {
                headBase + across * outline.HalfHead,
                outline.Tip,
                headBase - across * outline.HalfHead
            };
            if (DrawsShaft)
            {
                points.Add(headBase - across * outline.HalfShaft);
                points.Add(outline.Tail - across * outline.HalfShaft);
                points.Add(outline.Tail + across * outline.HalfShaft);
                points.Add(headBase + across * outline.HalfShaft);
            }

            if (pointCache.Count != points.Count)
            {
                pointCache.Clear();
                pointCache.AddRange(points);
                vertexCoordinates = new Point[points.Count];
            }
            else
            {
                for (int i = 0; i < points.Count; i++)
                {
                    pointCache[i] = points[i];
                }
            }

            // the polygon's hit testing works on logical vertices
            for (int i = 0; i < points.Count; i++)
            {
                VertexCoordinates[i] = ToLogical(pointCache[i]);
            }

            Shape.PointsChanged();
        }

        /// <summary>Of the head of a hairline arrow, in pixels; both grow with the line width</summary>
        public const double HeadLength = 11;
        public const double HeadHalfWidth = 4;
        public const double HeadGrowth = 3;

        /// <summary>
        /// One solid color for the shaft and the head, and no outline: the color of the line
        /// style. A vector from an older drawing has a polygon style with a transparent
        /// outline; that one keeps its fill.
        /// </summary>
        public override void ApplyStyle()
        {
            base.ApplyStyle();
            Shape.Fill = GetBrush(Style);
            Shape.Stroke = null;
            Shape.StrokeDashArray = null;
        }

        /// <summary>The color an arrow of the style is drawn in, as the style looks under the theme on screen</summary>
        public static Avalonia.Media.IBrush GetBrush(IFigureStyle style)
        {
            var resolved = style?.Resolve();
            if (resolved is LineStyle lineStyle && lineStyle.Color.A > 0)
            {
                return new Avalonia.Media.SolidColorBrush(lineStyle.Color);
            }

            return (Avalonia.Media.IBrush)(resolved as ShapeStyle)?.Fill ?? Avalonia.Media.Brushes.Black;
        }
    }
}
