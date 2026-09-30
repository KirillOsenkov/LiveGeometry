using System.Collections.Generic;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Controls.Shapes;

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

        /// <summary>The points from here on are this many pixels across (a crowded icon wants them smaller than the usual 8)</summary>
        public IconBuilder PointSize(double size)
        {
            pointSize = size;
            return this;
        }

        double? pointSize;

        public static IconBuilder BuildIcon()
        {
            return new IconBuilder();
        }

        public static IconBuilder BuildIcon(double size)
        {
            return new IconBuilder(size);
        }

        /// <summary>A free point as the theme draws it: its fill, rimmed in ink</summary>
        public IconBuilder Point(double x, double y)
        {
            AddPoint(x, y, nameof(AppTheme.FreePointFill));
            return this;
        }

        /// <summary>A point filled in a color of the theme (<c>nameof(AppTheme.PointOnFigureFill)</c>), rimmed in ink</summary>
        public IconBuilder Point(double x, double y, string fillThemeColor)
        {
            AddPoint(x, y, fillThemeColor);
            return this;
        }

        public IconBuilder TransparentPoint(double x, double y, double transparency)
        {
            AddPoint(x, y, nameof(AppTheme.FreePointFill)).Opacity = transparency;
            return this;
        }

        Shape AddPoint(double x, double y, string fillThemeColor)
        {
            Shape point = Factory.CreatePointShape();
            if (pointSize != null)
            {
                point.Width = pointSize.Value;
                point.Height = pointSize.Value;
            }

            point.BindTheme(Shape.StrokeProperty, nameof(AppTheme.Ink));
            if (fillThemeColor != null)
            {
                point.BindTheme(Shape.FillProperty, fillThemeColor);
            }

            Canvas.Children.Add(point);
            Canvas.SetLeft(point, Canvas.Width * x - point.Width / 2);
            Canvas.SetTop(point, Canvas.Height * y - point.Height / 2);
            return point;
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
            AddPoint(x, y, nameof(AppTheme.DependentPointFill));
            return this;
        }

        /// <summary>A line in the theme's ink, as a figure's line is drawn on the paper</summary>
        public IconBuilder Line(double x1, double y1, double x2, double y2)
        {
            return Line(nameof(AppTheme.Ink), x1, y1, x2, y2);
        }

        /// <summary>A line in a color of the theme (<c>nameof(AppTheme.Accent)</c>)</summary>
        public IconBuilder Line(string themeColor, double x1, double y1, double x2, double y2)
        {
            Line line = Factory.CreateLineShape();
            Canvas.Children.Add(line);
            line.X1 = Canvas.Width * x1;
            line.Y1 = Canvas.Height * y1;
            line.X2 = Canvas.Width * x2;
            line.Y2 = Canvas.Height * y2;
            line.BindTheme(Shape.StrokeProperty, themeColor);
            return this;
        }

        public IconBuilder Line(double strokeThickness, string themeColor, double x1, double y1, double x2, double y2)
        {
            Line(themeColor, x1, y1, x2, y2);
            ((Line)Canvas.Children[Canvas.Children.Count - 1]).StrokeThickness = strokeThickness;
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
                StrokeThickness = 1
            };
            path.BindTheme(Shape.StrokeProperty, nameof(AppTheme.Ink));
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
            circle.BindTheme(Shape.StrokeProperty, nameof(AppTheme.Ink));
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
            ellipse.BindTheme(Shape.StrokeProperty, nameof(AppTheme.Ink));
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
            arcInfo.Item3.SweepDirection = Avalonia.Media.SweepDirection.CounterClockwise;
            arcInfo.Item3.IsLargeArc = false;
            arcInfo.Item1.BindTheme(Shape.StrokeProperty, nameof(AppTheme.Ink));
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
            arcInfo.Item3.SweepDirection = Avalonia.Media.SweepDirection.CounterClockwise;
            arcInfo.Item3.IsLargeArc = false;
            arcInfo.Item3.RotationAngle = a;
            arcInfo.Item1.BindTheme(Shape.StrokeProperty, nameof(AppTheme.Ink));
            Canvas.Children.Add(arcInfo.Item1);
            return this;
        }

        public Avalonia.Controls.Shapes.Polygon AddPolygon(IEnumerable<Point> points)
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
            Brush fill, Brush stroke, params Avalonia.Point[] points)
        {
            var result = AddPolygon((IEnumerable<Point>)points);
            result.Fill = fill;
            result.Stroke = stroke;
            result.StrokeThickness = 1;
            result.StrokeJoin = PenLineJoin.Round;
            return this;
        }

        /// <summary>A filled polygon outlined in a color of the theme (the ink, as on the paper)</summary>
        public IconBuilder Polygon(
            Brush fill, string strokeThemeColor, params Avalonia.Point[] points)
        {
            var result = AddPolygon((IEnumerable<Point>)points);
            result.Fill = fill;
            result.BindTheme(Shape.StrokeProperty, strokeThemeColor);
            result.StrokeThickness = 1;
            result.StrokeJoin = PenLineJoin.Round;
            return this;
        }

        /// <summary>A polygon filled and outlined in a color of the theme (an arrowhead in ink)</summary>
        public IconBuilder Polygon(
            string fillThemeColor, string strokeThemeColor, params Avalonia.Point[] points)
        {
            var result = AddPolygon((IEnumerable<Point>)points);
            result.BindTheme(Shape.FillProperty, fillThemeColor);
            result.BindTheme(Shape.StrokeProperty, strokeThemeColor);
            result.StrokeThickness = 1;
            result.StrokeJoin = PenLineJoin.Round;
            return this;
        }

        /// <summary>A polygon in the theme's default fill, as a new polygon is drawn</summary>
        public IconBuilder Polygon(IEnumerable<Point> points)
        {
            AddPolygon(points).BindTheme(Shape.FillProperty, nameof(AppTheme.ShapeFill));
            return this;
        }

        public IconBuilder Polygon(params Avalonia.Point[] points)
        {
            return Polygon((IEnumerable<Point>)points);
        }

        public IconBuilder Polyline(Brush stroke, params Avalonia.Point[] points)
        {
            var result = AddPolyline((IEnumerable<Point>)points);
            result.Stroke = stroke;
            return this;
        }

        public Avalonia.Controls.Shapes.Polyline AddPolyline(IEnumerable<Point> points)
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

        /// <summary>Text in a color of the theme (the ink)</summary>
        public IconBuilder Text(string themeColor, double x1, double y1, string text, double fontSize = 0)
        {
            TextBlock textblock = new TextBlock();
            textblock.Text = text;
            if (fontSize > 0)
            {
                textblock.FontSize = fontSize;
            }

            Canvas.Children.Add(textblock);
            textblock.SetValue(Canvas.LeftProperty, x1 * Canvas.Width);
            textblock.SetValue(Canvas.TopProperty, y1 * Canvas.Height);
            textblock.BindTheme(TextBlock.ForegroundProperty, themeColor);

            return this;
        }
    }
}
