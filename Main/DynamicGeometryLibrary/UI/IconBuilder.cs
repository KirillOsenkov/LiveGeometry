using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace DynamicGeometry
{
    public partial class IconBuilder
    {
        public IconBuilder()
            : this(IconSize)
        {
        }

        public IconBuilder(double size)
        {
            Canvas = new Canvas();
            Canvas.Width = size;
            Canvas.Height = Canvas.Width;
        }

        public static double IconSize
        {
            get
            {
                return 32;
            }
        }

        public Canvas Canvas { get; set; }

        public static IconBuilder BuildIcon()
        {
            return new IconBuilder();
        }

        public static IconBuilder BuildIcon(double size)
        {
            return new IconBuilder(size);
        }

        public IconBuilder Point(double x, double y)
        {
            Shape point = Factory.CreatePointShape();
            Canvas.Children.Add(point);
            Canvas.SetLeft(point, Canvas.Width * x - point.Width / 2);
            Canvas.SetTop(point, Canvas.Height * y - point.Height / 2);
            return this;
        }

        public IconBuilder Point(double x, double y, Brush fill)
        {
            Shape point = Factory.CreatePointShape();
            Canvas.Children.Add(point);
            Canvas.SetLeft(point, Canvas.Width * x - point.Width / 2);
            Canvas.SetTop(point, Canvas.Height * y - point.Height / 2);
            point.Fill = fill;
            return this;
        }

        public IconBuilder TransparentPoint(double x, double y, double transparency)
        {
            Shape point = Factory.CreatePointShape();
            point.Opacity = transparency;
            Canvas.Children.Add(point);
            Canvas.SetLeft(point, Canvas.Width * x - point.Width / 2);
            Canvas.SetTop(point, Canvas.Height * y - point.Height / 2);
            return this;
        }

        public IconBuilder TransparentLine(double x1, double y1, double x2, double y2, double transparency)
        {
            Line line = Factory.CreateLineShape();
            Canvas.Children.Add(line);
            line.X1 = Canvas.Width * x1;
            line.Y1 = Canvas.Height * y1;
            line.X2 = Canvas.Width * x2;
            line.Y2 = Canvas.Height * y2;
            line.Opacity = transparency;
            return this;
        }

        public IconBuilder DependentPoint(double x, double y)
        {
            Shape point = Factory.CreateDependentPointShape();
            Canvas.Children.Add(point);
            Canvas.SetLeft(point, Canvas.Width * x - point.Width / 2);
            Canvas.SetTop(point, Canvas.Height * y - point.Height / 2);
            return this;
        }

        public IconBuilder Line(double x1, double y1, double x2, double y2)
        {
            Line line = Factory.CreateLineShape();
            Canvas.Children.Add(line);
            line.X1 = Canvas.Width * x1;
            line.Y1 = Canvas.Height * y1;
            line.X2 = Canvas.Width * x2;
            line.Y2 = Canvas.Height * y2;
            return this;
        }

        public IconBuilder Bezier(double x1, double y1, double x2, double y2, double x3, double y3, double x4, double y4)
        {
            var segment = new BezierSegment()
            {
                Point1 = new Point(Canvas.Width * x2, Canvas.Height * y2),
                Point2 = new Point(Canvas.Width * x3, Canvas.Height * y3),
                Point3 = new Point(Canvas.Width * x4, Canvas.Height * y4)
            };
            var figure = new PathFigure()
            {
                IsClosed = false,
                IsFilled = false,
                Segments = new PathSegmentCollection()
                {
                    segment
                },
                StartPoint = new Point(Canvas.Width * x1, Canvas.Height * y1)
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
            Canvas.Children.Add(path);
            return this;
        }

        public IconBuilder Line(Color color, double x1, double y1, double x2, double y2)
        {
            Line line = Factory.CreateLineShape();
            Canvas.Children.Add(line);
            line.X1 = Canvas.Width * x1;
            line.Y1 = Canvas.Height * y1;
            line.X2 = Canvas.Width * x2;
            line.Y2 = Canvas.Height * y2;
            line.Stroke = new SolidColorBrush(color);
            return this;
        }

        public IconBuilder Line(double strokeThickness, Color color, double x1, double y1, double x2, double y2)
        {
            Line line = Factory.CreateLineShape();
            Canvas.Children.Add(line);
            line.X1 = Canvas.Width * x1;
            line.Y1 = Canvas.Height * y1;
            line.X2 = Canvas.Width * x2;
            line.Y2 = Canvas.Height * y2;
            line.Stroke = new SolidColorBrush(color);
            line.StrokeThickness = strokeThickness;
            return this;
        }

        public IconBuilder Circle(double x, double y, double radius)
        {
            Shape circle = Factory.CreateCircleShape();
            Canvas.Children.Add(circle);
            circle.Width = Canvas.Width * radius * 2;
            circle.Height = Canvas.Height * radius * 2;
            Canvas.SetLeft(circle, Canvas.Width * x - circle.Width / 2);
            Canvas.SetTop(circle, Canvas.Height * y - circle.Height / 2);
            return this;
        }

        public IconBuilder Ellipse(double x, double y, double semiMajor, double semiMinor)
        {
            Shape ellipse = Factory.CreateCircleShape();
            Canvas.Children.Add(ellipse);
            ellipse.Width = Canvas.Width * semiMajor * 2;
            ellipse.Height = Canvas.Height * semiMinor * 2;
            Canvas.SetLeft(ellipse, Canvas.Width * x - ellipse.Width / 2);
            Canvas.SetTop(ellipse, Canvas.Height * y - ellipse.Height / 2);
            return this;
        }

        public IconBuilder Arc(double xc, double yc, double x1, double y1, double x2, double y2)
        {
            var arcInfo = Factory.CreateArcShape();
            arcInfo.Item2.StartPoint = new Point(Canvas.Width * x1, Canvas.Height * y1);
            arcInfo.Item3.Point = new Point(Canvas.Width * x2, Canvas.Height * y2);
            var radius = Math.Distance(xc, yc, x1, y1);
            arcInfo.Item3.Size = new Size(Canvas.Width * radius, Canvas.Height * radius);
            arcInfo.Item3.SweepDirection = System.Windows.Media.SweepDirection.Counterclockwise;
            arcInfo.Item3.IsLargeArc = false;
            Canvas.Children.Add(arcInfo.Item1);
            return this;
        }

        public IconBuilder EllipseArc(double xc, double yc, double x1, double y1, double x2, double y2, Size s, double a)
        {
            var arcInfo = Factory.CreateArcShape();
            arcInfo.Item2.StartPoint = new Point(Canvas.Width * x1, Canvas.Height * y1);
            arcInfo.Item3.Point = new Point(Canvas.Width * x2, Canvas.Height * y2);
            var radius = Math.Distance(xc, yc, x1, y1);
            arcInfo.Item3.Size = new Size(Canvas.Width * s.Width, Canvas.Height * s.Height);
            arcInfo.Item3.SweepDirection = System.Windows.Media.SweepDirection.Counterclockwise;
            arcInfo.Item3.IsLargeArc = false;
            arcInfo.Item3.RotationAngle = a;
            Canvas.Children.Add(arcInfo.Item1);
            return this;
        }

        public System.Windows.Shapes.Polygon AddPolygon(IEnumerable<Point> points)
        {
            var polygon = Factory.CreatePolygonShape();
            Canvas.Children.Add(polygon);
            foreach (var p in points)
            {
                polygon.Points.Add(new Point(
                        Canvas.Width * p.X,
                        Canvas.Height * p.Y));
            }
            return polygon;
        }

        public IconBuilder Polygon(
            Brush fill, Brush stroke, params System.Windows.Point[] points)
        {
            var result = AddPolygon((IEnumerable<Point>)points);
            result.Fill = fill;
            result.Stroke = stroke;
            return this;
        }

        public IconBuilder Polygon(IEnumerable<Point> points)
        {
            AddPolygon(points);
            return this;
        }

        public IconBuilder Polygon(params System.Windows.Point[] points)
        {
            return Polygon((IEnumerable<Point>)points);
        }

        public IconBuilder Polyline(Brush stroke, params System.Windows.Point[] points)
        {
            var result = AddPolyline((IEnumerable<Point>)points);
            result.Stroke = stroke;
            return this;
        }

        public System.Windows.Shapes.Polyline AddPolyline(IEnumerable<Point> points)
        {
            var polyline = Factory.CreatePolylineShape();
            Canvas.Children.Add(polyline);
            foreach (var p in points)
            {
                polyline.Points.Add(new Point(
                        Canvas.Width * p.X,
                        Canvas.Height * p.Y));
            }
            return polyline;
        }

        public IconBuilder Text(Color color, double x1, double y1, string text)
        {
            TextBlock textblock = new TextBlock();
            textblock.Text = text;
            Canvas.Children.Add(textblock);
            textblock.SetValue(Canvas.LeftProperty, x1 * Canvas.Width);
            textblock.SetValue(Canvas.TopProperty, y1 * Canvas.Height);
            textblock.Foreground = new SolidColorBrush(color);

            return this;
        }
    }
}
