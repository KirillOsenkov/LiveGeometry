using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;

namespace DynamicGeometry
{
    public abstract class Curve : ShapeBase<Path>, ILinearFigure
    {
        public Curve()
        {
            Shape = CreateShape();
            pathSegments = new PathSegmentCollection();
            pathFigure = new PathFigure()
            {
                Segments = pathSegments
            };
            Shape.Data = new PathGeometry()
            {
                Figures = new PathFigureCollection() 
                { 
                    pathFigure 
                }
            };
        }

        public override void Recalculate()
        {
        }

        public override void UpdateVisual()
        {
            try
            {
                Points.Clear();
                GetPoints(Points);
                if (Points.Count == 0)
                {
                    // Clearing the pathSegments is not triggering an immediate redraw.  Creating a new pathSegments does. - D.H.
                    //pathSegments.Clear();
                    pathSegments = new PathSegmentCollection();
                    pathFigure.Segments = pathSegments;
                    return;
                }

                logicalPoints.Capacity = System.Math.Max(Points.Count, logicalPoints.Capacity);
                var coordinateSystem = Drawing.CoordinateSystem;
                for (int i = 0; i < Points.Count; i++)
                {
                    if (logicalPoints.Count <= i)
                    {
                        logicalPoints.Add(Points[i]);
                    }
                    else
                    {
                        logicalPoints[i] = Points[i];
                    }

                    Points[i] = coordinateSystem.ToPhysical(Points[i]);
                }

                int count = logicalPoints.Count;
                if (count > Points.Count)
                {
                    for (int i = 0; i < count - Points.Count; i++)
                    {
                        logicalPoints.RemoveAt(count - i - 1);
                    }
                }

                pathFigure.StartPoint = Points[0];
                ConstructPolyline(Points);
            }
            catch (System.Exception)
            {
            }
        }

        List<Point> Points = new List<Point>();
        List<Point> logicalPoints = new List<Point>();

        protected virtual void ConstructPolyline(List<Point> points)
        {
            PolylineDirect(points);
        }

        private void PolylineDirect(List<Point> points)
        {
            for (int i = 0; i < points.Count - 1; i++)
            {
                if (pathSegments.Count <= i)
                {
                    pathSegments.Add(new LineSegment() { Point = points[i + 1] });
                }
                else
                {
                    ((LineSegment)pathSegments[i]).Point = points[i + 1];
                }
            }

            int segmentCount = pathSegments.Count;
            if (segmentCount > points.Count - 1)
            {
                for (int i = 0; i < segmentCount - points.Count + 1; i++)
                {
                    pathSegments.RemoveAt(segmentCount - i - 1);
                }
            }
        }

        public override Point Center
        {
            get
            {
                if (logicalPoints.IsEmpty())
                {
                    return base.Center;
                }
                return logicalPoints[(int)System.Math.Floor((double)(logicalPoints.Count)/2)];
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
        //                SweepDirection = SweepDirection.Counterclockwise,
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
        //                SweepDirection = SweepDirection.Counterclockwise,
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
        //            SweepDirection = previousSign == 1 ? SweepDirection.Clockwise : SweepDirection.Counterclockwise,
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
        //            SweepDirection = previousSign == 1 ? SweepDirection.Clockwise : SweepDirection.Counterclockwise,
        //            Size = new Size(radius, radius),
        //            Point = points[points.Count - 1],
        //            IsLargeArc = false
        //        });
        //    }
        //}

        protected PathFigure pathFigure;
        protected PathSegmentCollection pathSegments;

        public abstract void GetPoints(List<Point> points);

        public override IFigure HitTest(System.Windows.Point point)
        {
            double epsilon = ToLogical(Shape.StrokeThickness / 2 + Math.CursorTolerance);
            if (Math.IsPointOnPolygonalChain(logicalPoints, point, epsilon, false))
            {
                return this;
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
