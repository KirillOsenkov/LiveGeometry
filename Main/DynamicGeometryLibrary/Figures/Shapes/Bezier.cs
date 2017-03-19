using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;

namespace DynamicGeometry
{
    public class Bezier : ShapeBase<Path>, ILinearFigure
    {
        public PathFigure Figure { get; set; }
        public BezierSegment BezierShape { get; set; }
        Math.BezierInfo Info;

        public override void Recalculate()
        {
            var p0 = Point(0);
            var p1 = Point(1);
            var p2 = Point(2);
            var p3 = Point(3);

            Info = new Math.BezierInfo(p0, p1, p2, p3);
        }

        public override void UpdateVisual()
        {
            var p0 = Point(0);
            var p1 = Point(1);
            var p2 = Point(2);
            var p3 = Point(3);

            Figure.StartPoint = ToPhysical(p0);
            BezierShape.Point1 = ToPhysical(p1);
            BezierShape.Point2 = ToPhysical(p2);
            BezierShape.Point3 = ToPhysical(p3);

            ShapeStyle shapeStyle = Style as ShapeStyle;
            if (shapeStyle != null)
            {
                Figure.IsFilled = shapeStyle.IsFilled;
            }
            else
            {
                Figure.IsFilled = false;
            }
        }

        public override Point Center
        {
            get
            {
                // For now use a very rough centroid approximation.
                return Math.Midpoint(Point(0), Point(1), Point(2), Point(3));
            }
        }

        protected override Path CreateShape()
        {
            BezierShape = new BezierSegment()
            {

            };
            Figure = new PathFigure()
            {
                IsClosed = false,
                IsFilled = false,
                Segments = new PathSegmentCollection()
                {
                    BezierShape
                }
            };
            return new Path()
            {
                Data = new PathGeometry()
                {
                    Figures = new PathFigureCollection()
                    {
                        Figure
                    }
                },
                Stroke = new SolidColorBrush(Colors.Black),
                StrokeThickness = 1
            };
        }

        public override IFigure HitTest(Point point)
        {
            // Curve HitTest
            var projection = Info.GetProjection(point);
            if (projection.DistanceToLine < ToLogical(Shape.StrokeThickness / 2 + Math.CursorTolerance))
            {
                return this;
            }

            // Fill HitTest
            ShapeStyle shapeStyle = Style as ShapeStyle;
            if (shapeStyle != null)
            {
                if (shapeStyle.IsFilled)
                {
                    return HitTestShape(point);
                }
            }
            return null;
        }

        public double GetNearestParameterFromPoint(Point point)
        {
            return Info.GetNearestParameterFromPoint(point);
        }

        public Point GetPointFromParameter(double parameter)
        {
            return Math.GetPointOnPolylineFromParameter(Info.Points, parameter);
        }

        public Tuple<double, double> GetParameterDomain()
        {
            return Tuple.Create(0.0, 1.0);
        }
    }
}
