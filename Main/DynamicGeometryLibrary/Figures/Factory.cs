using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace DynamicGeometry
{
    public partial class Factory
    {
        const int size = 8;

        public static Shape CreatePointShape()
        {
            System.Windows.Shapes.Ellipse ellipse = new System.Windows.Shapes.Ellipse()
            {
                Width = size,
                Height = size,
                Fill = new SolidColorBrush(Colors.Yellow),
                Stroke = new SolidColorBrush(Colors.Black),
                StrokeThickness = 0.5
            };

            return ellipse;
        }

        public static Shape CreateDependentPointShape()
        {
            System.Windows.Shapes.Ellipse ellipse = new System.Windows.Shapes.Ellipse()
            {
                Width = size,
                Height = size,
                //Fill = new SolidColorBrush(Color.FromArgb(255, 240, 240, 240)),
                Fill = CreateDefaultFillBrush(),
                Stroke = new SolidColorBrush(Colors.Black),
                StrokeThickness = 0.5
            };

            return ellipse;
        }

        public static LinearGradientBrush CreateLinearGradient(Color source, Color target, double angle)
        {
            return new LinearGradientBrush(
                new GradientStopCollection()
                {
                    new GradientStop
                    {
                        Offset = 0,
                        Color = source
                    },
                    new GradientStop
                    {
                        Offset = 1,
                        Color = target
                    }
                },
                angle);
        }

        public static Shape CreateCircleShape()
        {
            int size = 8;
            System.Windows.Shapes.Ellipse ellipse = new System.Windows.Shapes.Ellipse()
            {
                Width = size,
                Height = size,
                // Stroke and Fill Color are used by IconBuilder but not by shape creators.
                Stroke = new SolidColorBrush(Colors.Black),
                StrokeThickness = 1
            };

            return ellipse;
        }

        public static System.Windows.Shapes.Polygon CreatePolygonShape()
        {
            return new System.Windows.Shapes.Polygon()
            {
                Fill = CreateDefaultFillBrush()
            };
        }

        public static Brush CreateDefaultFillBrush()
        {
            return new SolidColorBrush(Color.FromArgb(255, 255, 255, 200));
        }

        public static Brush CreateGradientBrush(Color c1, Color c2, double angle)
        {
            LinearGradientBrush result = new LinearGradientBrush(
                new GradientStopCollection()
                {
                    new GradientStop() { Offset = 0.1, Color = c1 },
                    new GradientStop() { Offset = 1.0, Color = c2 }
                }, angle);
            return result;
        }

        public static Line CreateLineShape()
        {
            Line result = new Line()
            {
                Stroke = new SolidColorBrush(Colors.Black),
                StrokeThickness = 1
            };

            return result;
        }

        public static TextBlock CreateLabelShape()
        {
            return new TextBlock()
            {
            };
        }

        public static MidPoint CreateMidPoint(Drawing drawing, IList<IFigure> dependencies)
        {
            MidPoint result = new MidPoint() { Drawing = drawing, Dependencies = dependencies };
            return result;
        }

        public static LineTwoPoints CreateLineTwoPoints(Drawing drawing, IList<IFigure> dependencies)
        {
            return new LineTwoPoints() { Drawing = drawing, Dependencies = dependencies };
        }

        public static FreePoint CreateFreePoint(Drawing drawing, Point coordinates)
        {
            FreePoint result = new FreePoint() { Drawing = drawing };
            result.MoveTo(coordinates);
            return result;
        }


        /// <param name="angle">in degrees.</param>
        public static RotatedPoint CreateRotatedPoint(Drawing drawing, IList<IFigure> dependencies, double angle)
        {
            return new RotatedPoint() { Drawing = drawing, Dependencies = dependencies, Angle = angle };
        }

        public static Segment CreateSegment(Drawing drawing, IList<IFigure> dependencies)
        {
            return new Segment() { Drawing = drawing, Dependencies = dependencies };
        }

        public static Vector CreateVector(Drawing drawing, IList<IFigure> dependencies)
        {
            return new Vector() { Drawing = drawing, Dependencies = dependencies };
        }

        public static Segment CreateSegment(Drawing drawing, IPoint point1, IPoint point2)
        {
            return CreateSegment(drawing, new [] { point1, point2 });
        }

        public static SegmentBisector CreateSegmentBisector(Drawing drawing, IList<IFigure> dependencies)
        {
            return new SegmentBisector() { Drawing = drawing, Dependencies = dependencies };
        }

        public static TranslatedPoint CreateTranslatedPoint(Drawing drawing, IList<IFigure> dependencies, double magnitude, double direction)
        {
            return new TranslatedPoint() { Drawing = drawing, Dependencies = dependencies, Magnitude = magnitude, Direction = direction };
        }

        public static Ray CreateRay(Drawing drawing, IList<IFigure> dependencies)
        {
            return new Ray() { Drawing = drawing, Dependencies = dependencies };
        }

        public static Circle CreateCircle(Drawing drawing, IList<IFigure> dependencies)
        {
            return new Circle() { Drawing = drawing, Dependencies = dependencies };
        }

        public static CircleByRadius CreateCircleByRadius(Drawing drawing, IList<IFigure> dependencies)
        {
            return new CircleByRadius() { Drawing = drawing, Dependencies = dependencies };
        }

        public static Ellipse CreateEllipse(Drawing drawing, IList<IFigure> dependencies)
        {
            return new Ellipse() { Drawing = drawing, Dependencies = dependencies };
        }

        public static ParallelLine CreateParallelLine(Drawing drawing, IList<IFigure> dependencies)
        {
            return new ParallelLine() { Drawing = drawing, Dependencies = dependencies };
        }

        public static PerpendicularLine CreatePerpendicularLine(Drawing drawing, IList<IFigure> dependencies)
        {
            return new PerpendicularLine() { Drawing = drawing, Dependencies = dependencies };
        }

        public static IntersectionPoint CreateIntersectionPoint(
            Drawing drawing, IFigure figure1, IFigure figure2, Point hintPoint)
        {
            return new IntersectionPoint(hintPoint, new List<IFigure> { figure1, figure2 }) { Drawing = drawing };
        }

        /// <summary>
        /// Creates a point on a linear figure (ILinearFigure)
        /// </summary>
        /// <param name="drawing">Drawing to add the figure to</param>
        /// <param name="iFigure">A linear figure such as a line or a circle</param>
        /// <param name="point">Hint point coordinates - this point will be
        /// projected on to the figure if it's not already on it</param>
        /// <returns>The newly created point on the figure</returns>
        public static PointOnFigure CreatePointOnFigure(Drawing drawing, IFigure iFigure, Point point)
        {
            var result = new PointOnFigure()
            {
                Drawing = drawing,
                Dependencies = new [] { iFigure },
            };
            result.Parameter = result.LinearFigure.GetNearestParameterFromPoint(point);
            return result;
        }

        /// <summary>
        /// Creates a point on a linear figure (ILinearFigure)
        /// </summary>
        /// <param name="drawing">Drawing to add the figure to</param>
        /// <param name="figure">A linear figure such as a line or a circle</param>
        /// <param name="parameter">A double parameter that defines the
        /// position of the point on the figure
        /// [0, 2 * PI) for circles, [0, 1] for segments, etc. </param>
        /// <returns>The newly created point on the figure</returns>
        public static PointOnFigure CreatePointOnFigure(Drawing drawing, IFigure figure, double parameter)
        {
            var result = new PointOnFigure()
            {
                Drawing = drawing,
                Dependencies = new [] { figure },
            };
            result.Parameter = parameter;
            return result;
        }

        public static ReflectedPoint CreateReflectedPoint(Drawing drawing, IList<IFigure> dependencies)
        {
            return new ReflectedPoint() { Drawing = drawing, Dependencies = dependencies };
        }

        public static DilatedPoint CreateDilatedPoint(Drawing drawing, IList<IFigure> dependencies, double factor)
        {
            return new DilatedPoint() { Drawing = drawing, Dependencies = dependencies, Factor = factor };
        }

        public static DistanceMeasurement CreateDistanceMeasurement(Drawing drawing, IList<IFigure> dependencies)
        {
            return new DistanceMeasurement() { Drawing = drawing, Dependencies = dependencies };
        }

        public static PointLabel CreatePointLabel(Drawing drawing, IList<IFigure> dependencies)
        {
            return new PointLabel() { Drawing = drawing, Dependencies = dependencies };
        }

        public static Label CreateLabel(Drawing drawing, IList<IFigure> dependencies)
        {
            return new Label() { Drawing = drawing, Dependencies = dependencies };
        }

        public static Label CreateLabel(Drawing drawing)
        {
            return new Label() { Drawing = drawing };
        }

        public static Polygon CreatePolygon(Drawing drawing, IList<IFigure> dependencies)
        {
            var result = new Polygon() { Drawing = drawing, Dependencies = dependencies };
            return result;
        }

        /// <summary>
        /// Creates an angle measurement figure (doesn't include AngleArc)
        /// </summary>
        /// <param name="drawing">Drawing to add the figure to</param>
        /// <param name="dependencies">3 points: vertex, point1 and point2
        /// such that the turn from point1 to point2 is counterclockwise
        /// </param>
        /// <returns>A newly created angle measurement (doesn't include AngleArc)</returns>
        public static AngleMeasurement CreateAngleMeasurement(Drawing drawing, IList<IFigure> dependencies)
        {
            return new AngleMeasurement() { Drawing = drawing, Dependencies = dependencies };
        }

        public static AngleBisector CreateAngleBisector(Drawing drawing, IList<IFigure> dependencies)
        {
            return new AngleBisector() { Drawing = drawing, Dependencies = dependencies };
        }

        public static AngleArc CreateAngleArc(Drawing drawing, IList<IFigure> dependencies)
        {
            return new AngleArc() { Drawing = drawing, Dependencies = dependencies };
        }

        public static AreaMeasurement CreateAreaMeasurement(Drawing drawing, IList<IFigure> dependencies)
        {
            return new AreaMeasurement() { Drawing = drawing, Dependencies = dependencies };
        }

        public static PointByCoordinates CreatePointByCoordinates(Drawing drawing, IList<IFigure> dependencies)
        {
            return new PointByCoordinates() { Drawing = drawing, Dependencies = dependencies };
        }

        public static PointByCoordinates CreatePointByCoordinates(Drawing drawing, string x, string y)
        {
            var result = new PointByCoordinates() { Drawing = drawing };
            result.XExpression.Text = x;
            result.YExpression.Text = y;
            result.XExpression.Recalculate();
            result.YExpression.Recalculate();
            return result;
        }

        public static PointByCoordinates CreatePointByCoordinates(Drawing drawing, Func<double> xValueProvider, Func<double> yValueProvider)
        {
            var result = new PointByCoordinates()
            {
                Drawing = drawing
            };
            result.XExpression.Value = xValueProvider;
            result.YExpression.Value = yValueProvider;
            return result;
        }

        public static Axis CreateAxis(Drawing drawing, IList<IFigure> dependencies)
        {
            var result = new Axis() { Drawing = drawing };
            result.Line.Dependencies = dependencies;
            return result;
        }

        public static Locus CreateLocus(Drawing Drawing, IList<IFigure> dependencies)
        {
            var result = new Locus() { Drawing = Drawing, Dependencies = dependencies };
            return result;
        }

        public static CircleArc CreateArc(Drawing drawing, IList<IFigure> foundDependencies)
        {
            var result = new CircleArc() { Drawing = drawing, Dependencies = foundDependencies };
            return result;
        }

        public static EllipseArc CreateEllipseArc(Drawing drawing, IList<IFigure> foundDependencies)
        {
            var result = new EllipseArc() { Drawing = drawing, Dependencies = foundDependencies };
            return result;
        }

        public static CircleSegment CreateCircleSegment(Drawing drawing, IList<IFigure> foundDependencies)
        {
            var result = new CircleSegment() { Drawing = drawing, Dependencies = foundDependencies };
            return result;
        }

        public static EllipseSegment CreateEllipseSegment(Drawing drawing, IList<IFigure> foundDependencies)
        {
            var result = new EllipseSegment() { Drawing = drawing, Dependencies = foundDependencies };
            return result;
        }

        public static EllipseSector CreateEllipseSector(Drawing drawing, IList<IFigure> foundDependencies)
        {
            var result = new EllipseSector() { Drawing = drawing, Dependencies = foundDependencies };
            return result;
        }

        public static CircleSector CreateCircleSector(Drawing drawing, IList<IFigure> foundDependencies)
        {
            var result = new CircleSector() { Drawing = drawing, Dependencies = foundDependencies };
            return result;
        }

        public static Tuple<Path, PathFigure, ArcSegment> CreateArcShape()
        {
            var arcSegment = new ArcSegment()
            {
                SweepDirection = SweepDirection.Counterclockwise,
                RotationAngle = 0
            };
            var figure = new PathFigure()
            {
                IsClosed = false,
                IsFilled = false,
                Segments = new PathSegmentCollection()
                {
                    arcSegment
                }
            };
            var path = new Path()
            {
                Data = new PathGeometry()
                {
                    Figures = new PathFigureCollection()
                    {
                        figure
                    }
                },
                Stroke = new SolidColorBrush(Colors.Black),
                StrokeThickness = 1
            };
            return Tuple.Create(path, figure, arcSegment);
        }

        public static MidPoint CreateMidPoint(Drawing Drawing, Segment segment)
        {
            return Factory.CreateMidPoint(Drawing, segment.Dependencies);
        }

        public static IFigure CreateBezier(Drawing drawing, IList<IFigure> foundDependencies)
        {
            return new Bezier() { Drawing = drawing, Dependencies = foundDependencies };
        }

        public static LineByEquation CreateLineByEquation(Drawing drawing, IList<IFigure> figureList)
        {
            return new LineByEquation() { Drawing = drawing, Dependencies = figureList };
        }

        public static LineByEquation CreateLineByEquation(Drawing drawing, string m, string b)
        {
            var result = new LineByEquation() { Drawing = drawing };
            var equation = new SlopeInterseptLineEquation(result, m, b);
            result.Equation = equation;
            equation.Recalculate();
            return result;
        }

        public static LineByEquation CreateLineByEquation(Drawing drawing, string a, string b, string c)
        {
            var result = new LineByEquation() { Drawing = drawing };
            var equation = new GeneralFormLineEquation(result, a, b, c);
            result.Equation = equation;
            equation.Recalculate();
            return result;
        }

        public static CircleByEquation CreateCircleByEquation(Drawing drawing, IList<IFigure> figureList)
        {
            return new CircleByEquation() { Drawing = drawing, Dependencies = figureList };
        }

        public static CircleByEquation CreateCircleByEquation(Drawing drawing, string centerX, string centerY, string radius)
        {
            var circle = new CircleByEquation { Drawing = drawing };
            circle.X = new DrawingExpression(circle, "Center X =", centerX);
            circle.Y = new DrawingExpression(circle, "Center Y =", centerY);
            circle.R = new DrawingExpression(circle, "Radius =", radius);
            circle.X.Recalculate();
            circle.Y.Recalculate();
            circle.R.Recalculate();
            return circle;
        }

        // added by banchmax ------------------------------------------------

        /// <summary>
        /// Creates an angle measurement figure with horizontal axis
        /// </summary>
        public static HorizontalAngleMeasurement CreateHorizontalAngleMeasurement(Drawing drawing, IList<IFigure> figureList)
        {
            return new HorizontalAngleMeasurement() { Drawing = drawing, Dependencies = figureList };
        }

        public static Polyline CreatePolyline(Drawing drawing, IList<IFigure> figureList)
        {
            var result = new Polyline() { Drawing = drawing, Dependencies = figureList };
            return result;
        }

        public static Brush CreateBlackFillBrush()
        {
            return new SolidColorBrush(Color.FromArgb(255, 0, 0, 0));
        }

        public static System.Windows.Shapes.Polyline CreatePolylineShape()
        {
            return new System.Windows.Shapes.Polyline()
            {
                Stroke = new SolidColorBrush(Colors.Black),
                StrokeThickness = 1,
            };
        }

        public static RegularPolygon CreateRegularPolygon(Drawing drawing, IList<IFigure> dependencies)
        {
            var result = new RegularPolygon() { Drawing = drawing, Dependencies = dependencies };
            return result;
        }

        public static PolygonIntersection CreatePolygonIntersection(Drawing drawing, IList<IFigure> dependencies) 
        {
            var result = new PolygonIntersection { Drawing = drawing, Dependencies = dependencies };
            return result;
        }
    }
}
