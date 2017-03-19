using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using DynamicGeometry;
using PolylineRouting;

namespace PolylineRoutingDemo1
{
    public partial class Page : UserControl
    {
        public Page()
        {
            InitializeComponent();
        }

        public Drawing CurrentDrawing { get; set; }
        public DynamicGeometry.Polygon Polygon { get; set; }
        public Point StartPoint { get; set; }
        public Point EndPoint { get; set; }
        public List<Point> Vertices { get; set; }
        public Segment Segment { get; set; }
        public DynamicGeometry.PointBase Start { get; set; }
        public DynamicGeometry.PointBase End { get; set; }
        public Route Route { get; set; }

        void canvas1_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (CurrentDrawing == null)
            {
                UpdateConfigurationFromText(@"12,10,1,1
10
2,7,2,6,4,2,6,1,8,2,9,5,10,5,11,6,9,9,4,10");
            }
        }

        private void UpdateConfigurationFromText(string text)
        {
            var lines = text.Split('\n', '\r')
                .Where(s => !string.IsNullOrEmpty(s)).ToArray();
            if (lines == null || lines.Length < 3)
            {
                return;
            }
            RoutingAlgorithm algorithm = new RoutingAlgorithm();
            algorithm.ParseInput(lines[0], lines[1], lines[2]);
            CreateConfiguration(algorithm.Start, algorithm.End, algorithm.Polygon);
        }

        void CreateConfiguration(Point start, Point end, List<Point> vertices)
        {
            StartPoint = start;
            EndPoint = end;
            Vertices = vertices;

            if (CurrentDrawing != null)
            {
                CurrentDrawing.Canvas = null;
            }
            CurrentDrawing = new Drawing(canvas1);
            CurrentDrawing.Behavior = new Dragger();
            CurrentDrawing.CoordinateGrid.Visible = false;
            CurrentDrawing.CoordinateSystem.UnitLength = 20;
            CurrentDrawing.CoordinateSystem.MoveTo(vertices.Midpoint().Minus());

            List<IFigure> points = new List<IFigure>();
            foreach (var vertex in vertices)
            {
                var point = Factory.CreateFreePoint(CurrentDrawing, vertex);
                points.Add(point);
                Actions.Add(CurrentDrawing, point);
            }

            Polygon = Factory.CreatePolygon(CurrentDrawing, points);
            Actions.Add(CurrentDrawing, Polygon);
            var polygonStyle = Polygon.Style as ShapeStyle;
            polygonStyle.Fill = new SolidColorBrush(Colors.Cyan);

            Start = Factory.CreateFreePoint(CurrentDrawing, start);
            End = Factory.CreateFreePoint(CurrentDrawing, end);

            Actions.Add(CurrentDrawing, Start);
            Actions.Add(CurrentDrawing, End);
            var pointStyle = Start.Style as PointStyle;
            pointStyle.Fill = new SolidColorBrush(Color.FromArgb(255, 0, 255, 0));

            Route = new Route()
            {
                Dependencies = new List<IFigure>()
                {
                    Start,
                    End,
                    Polygon
                },
                Drawing = CurrentDrawing
            };

            Actions.Add(CurrentDrawing, Route);
        }
    }
}
