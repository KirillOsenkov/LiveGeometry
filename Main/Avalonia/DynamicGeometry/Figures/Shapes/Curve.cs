using System.Collections.Generic;
using Avalonia;
using Avalonia.Media;
using Avalonia.Controls.Shapes;

namespace DynamicGeometry
{
    public abstract class Curve : ShapeBase<Path>, ILinearFigure
    {
        public Curve()
        {
            Shape = CreateShape();
        }

        /// <summary>
        /// Among the points of a curve (<see cref="GetPoints"/>): the curve is interrupted
        /// here, the point before and the point after are not joined. A graph has gaps
        /// where the function has no value or jumps, a locus where the traced point is
        /// not there. (A curve was one line through all its points: the graph of
        /// sqrt(x^2 - 4) had a straight piece across the stretch where there is no root,
        /// that of 1/x a vertical line at 0 joining its two branches.)
        /// </summary>
        public static readonly Point Gap = new Point(double.NaN, double.NaN);

        public override void Recalculate()
        {
        }

        public override void UpdateVisual()
        {
            try
            {
                logicalPoints.Clear();
                GetPoints(logicalPoints);

                // one figure for each stretch between gaps, each open and not filled: in
                // Avalonia a figure is closed unless told otherwise, which draws a line
                // from the end of a graph back to its start
                var geometry = new StreamGeometry();
                using (var context = geometry.Open())
                {
                    var coordinateSystem = Drawing.CoordinateSystem;
                    bool isDrawing = false;
                    for (int i = 0; i < logicalPoints.Count; i++)
                    {
                        if (!logicalPoints[i].Exists())
                        {
                            if (isDrawing)
                            {
                                context.EndFigure(isClosed: false);
                                isDrawing = false;
                            }

                            continue;
                        }

                        var physical = coordinateSystem.ToPhysical(logicalPoints[i]);
                        if (isDrawing)
                        {
                            context.LineTo(physical);
                        }
                        else if (i + 1 < logicalPoints.Count && logicalPoints[i + 1].Exists())
                        {
                            // (a point with a gap on either side is no line to draw)
                            context.BeginFigure(physical, isFilled: false);
                            isDrawing = true;
                        }
                    }

                    if (isDrawing)
                    {
                        context.EndFigure(isClosed: false);
                    }
                }

                Shape.Data = geometry;
            }
            catch (System.Exception)
            {
            }
        }

        // as GetPoints gave them, in the units of the drawing, gaps included
        List<Point> logicalPoints = new List<Point>();

        public override Point Center
        {
            get
            {
                // the middle one of the points, or the nearest to it that is not a gap
                for (int offset = 0; offset < logicalPoints.Count; offset++)
                {
                    int middle = logicalPoints.Count / 2;
                    foreach (int index in new[] { middle + offset, middle - offset })
                    {
                        if (index >= 0 && index < logicalPoints.Count && logicalPoints[index].Exists())
                        {
                            return logicalPoints[index];
                        }
                    }
                }

                return base.Center;
            }
        }

        //public static void PolylineRounding(List<Point> points, PathSegmentCollection segments)
        //{
        //    double radius = 16;
        //    if (points.Count < 2)
        //    {
        //        return;
        //    }
        //    else if (points.Count == 2)
        //    {
        //        segments.Add(new LineSegment() { Point = points[1] });
        //        return;
        //    }

        //    int previousSign = Math.VectorProduct(points[0], points[1], points[2]).Sign();
        //    var tangentPoints = Math.GetTangentPoints(points[0], points[1], radius);
        //    Point previousPoint;
        //    if (previousSign > 0)
        //    {
        //        previousPoint = tangentPoints.P1;
        //    }
        //    else
        //    {
        //        previousPoint = tangentPoints.P2;
        //    }
        //    if (previousPoint.Exists() && points[0].Distance(previousPoint) >= radius)
        //    {
        //        segments.Add(new LineSegment() { Point = previousPoint });
        //    }

        //    for (int i = 2; i < points.Count - 1; i++)
        //    {
        //        Point p1 = new Point();
        //        Point p2 = new Point();
        //        int sign = Math.VectorProduct(points[i - 1], points[i], points[i + 1]).Sign();
        //        if (previousSign == 0)
        //        {
        //            previousSign = sign;
        //        }
        //        if (sign == 0)
        //        {
        //            p2 = points[i];
        //        }
        //        else if (sign == 1 && previousSign == 1)
        //        {
        //            var vector = Math.RotatePoint(
        //                points[i - 1],
        //                radius,
        //                (points[i - 1].AngleTo(points[i]) - Math.PI / 2)).Minus(points[i - 1]);
        //            p1 = points[i - 1].Plus(vector);
        //            p2 = points[i].Plus(vector);
        //            segments.Add(new ArcSegment()
        //            {
        //                SweepDirection = SweepDirection.Clockwise,
        //                Size = new Size(radius, radius),
        //                Point = p1,
        //                IsLargeArc = false
        //            });
        //        }
        //        else if (sign == -1 && previousSign == -1)
        //        {
        //            var vector = Math.RotatePoint(
        //                points[i - 1],
        //                radius,
        //                (points[i - 1].AngleTo(points[i]) + Math.PI / 2)).Minus(points[i - 1]);
        //            p1 = points[i - 1].Plus(vector);
        //            p2 = points[i].Plus(vector);
        //            segments.Add(new ArcSegment()
        //            {
        //                SweepDirection = SweepDirection.CounterClockwise,
        //                Size = new Size(radius, radius),
        //                Point = p1,
        //                IsLargeArc = false
        //            });
        //        }
        //        else if (previousSign == -1 && sign == 1)
        //        {
        //            var midpoint = Math.Midpoint(points[i - 1], points[i]);
        //            tangentPoints = Math.GetTangentPoints(midpoint, points[i - 1], radius);
        //            p1 = tangentPoints.P1;
        //            p2 = p1.Reflect(midpoint);
        //            segments.Add(new ArcSegment()
        //            {
        //                SweepDirection = SweepDirection.CounterClockwise,
        //                Size = new Size(radius, radius),
        //                Point = p1,
        //                IsLargeArc = false
        //            });
        //        }
        //        else if (previousSign == 1 && sign == -1)
        //        {
        //            var midpoint = Math.Midpoint(points[i - 1], points[i]);
        //            tangentPoints = Math.GetTangentPoints(midpoint, points[i - 1], radius);
        //            p1 = tangentPoints.P2;
        //            p2 = p1.Reflect(midpoint);
        //            segments.Add(new ArcSegment()
        //            {
        //                SweepDirection = SweepDirection.Clockwise,
        //                Size = new Size(radius, radius),
        //                Point = p1,
        //                IsLargeArc = false
        //            });
        //        }
        //        segments.Add(new LineSegment() { Point = p2 });
        //        previousPoint = p2;
        //        previousSign = sign;
        //    }

        //    tangentPoints = Math.GetTangentPoints(points[points.Count - 1],
        //        points[points.Count - 2],
        //        radius);
        //    previousPoint = previousSign == 1 ? tangentPoints.P2 : tangentPoints.P1;
        //    if (points[points.Count - 1].Distance(previousPoint) >= radius
        //        && previousPoint != Math.InfinitePoint)
        //    {
        //        segments.Add(new ArcSegment()
        //        {
        //            SweepDirection = previousSign == 1 ? SweepDirection.Clockwise : SweepDirection.CounterClockwise,
        //            Size = new Size(radius, radius),
        //            Point = previousPoint,
        //            IsLargeArc = false
        //        });
        //        segments.Add(new LineSegment() { Point = points[points.Count - 1] });
        //    }
        //    else
        //    {
        //        segments.Add(new ArcSegment()
        //        {
        //            SweepDirection = previousSign == 1 ? SweepDirection.Clockwise : SweepDirection.CounterClockwise,
        //            Size = new Size(radius, radius),
        //            Point = points[points.Count - 1],
        //            IsLargeArc = false
        //        });
        //    }
        //}

        /// <summary>The points of the curve, in the units of the drawing, with a <see cref="Gap"/> wherever it is interrupted</summary>
        public abstract void GetPoints(List<Point> points);

        public override IFigure HitTest(Avalonia.Point point)
        {
            double epsilon = ToLogical(Shape.StrokeThickness / 2 + Math.CursorTolerance);
            for (int i = 1; i < logicalPoints.Count; i++)
            {
                var from = logicalPoints[i - 1];
                var to = logicalPoints[i];
                if (!from.Exists() || !to.Exists())
                {
                    continue;
                }

                // (at a corner, past the end of both pieces, the place is near the corner itself)
                if (Math.IsPointOnSegment(new PointPair(from, to), point, epsilon)
                    || from.Distance(point) < epsilon
                    || to.Distance(point) < epsilon)
                {
                    return this;
                }
            }

            return null;
        }

        protected override Path CreateShape()
        {
            var result = new Path();
            result.Stroke = new SolidColorBrush(Color.FromArgb(255, 255, 150, 150));
            result.StrokeThickness = 1;
            return result;
        }

        public virtual double GetNearestParameterFromPoint(Point point)
        {
            return Math.GetNearestParameterFromPointOnPolyline(logicalPoints, point);
        }

        public virtual Point GetPointFromParameter(double parameter)
        {
            return Math.GetPointOnPolylineFromParameter(logicalPoints, parameter);
        }

        public virtual Tuple<double, double> GetParameterDomain()
        {
            return Tuple.Create(0.0, 1.0);
        }
    }

    public class CustomCurve : Curve
    {
        public List<Point> Points = new List<Point>();

        public override void GetPoints(List<Point> points)
        {
            if (Points.IsEmpty())
            {
                return;
            }

            points.AddRange(Points);
        }
    }
}
