using System.Linq;
using Avalonia;

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

            // All in pixels: an arrow is a line with a head, as wide as its style says and with a
            // head that goes with that width - the same at any zoom.
            Point tail = ToPhysical(line.P1);
            Point tip = ToPhysical(line.P2);
            double length = tail.Distance(tip);

            var lineStyle = Style as LineStyle;
            double width = lineStyle != null ? lineStyle.StrokeWidth : 1;
            if (Selected && Settings.ChangeLineAppearanceWhenSelected)
            {
                width += 3;
            }

            double halfShaft = System.Math.Max(width / 2, 0.5);
            double headLength = System.Math.Min(HeadLength + HeadGrowth * width, length);
            double halfHead = HeadHalfWidth + HeadGrowth / 2 * width;

            var along = length > 1e-9 ? (tip - tail) / length : new Point(1, 0);
            var across = new Point(-along.Y, along.X);

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

            var headBase = tip - along * headLength;

            pointCache[0] = headBase + across * halfHead;
            pointCache[1] = tip;
            pointCache[2] = headBase - across * halfHead;
            pointCache[3] = headBase - across * halfShaft;
            pointCache[4] = tail - across * halfShaft;
            pointCache[5] = tail + across * halfShaft;
            pointCache[6] = headBase + across * halfShaft;

            // the polygon's hit testing works on logical vertices
            for (int i = 0; i < 7; i++)
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

            Avalonia.Media.IBrush brush = null;
            var lineStyle = Style as LineStyle;
            if (lineStyle != null && lineStyle.Color.A > 0)
            {
                brush = new Avalonia.Media.SolidColorBrush(lineStyle.Color);
            }
            else if (Style is ShapeStyle shapeStyle)
            {
                brush = shapeStyle.Fill;
            }

            Shape.Fill = brush ?? Avalonia.Media.Brushes.Black;
            Shape.Stroke = null;
            Shape.StrokeDashArray = null;
        }
    }
}
