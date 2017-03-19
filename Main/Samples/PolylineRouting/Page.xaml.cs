using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using DynamicGeometry;

namespace PolylineRouting
{
    public partial class Page : UserControl
    {
        public Page()
        {
            InitializeComponent();
            text.Text = @"4,7,1,6
11
2,1,6,1,3,2,6,3,3,5,5,5,7,3,5,2,8,1,8,6,2,6";
            //            text.Text = @"4,7,1,6
            //4
            //2,4,5,4,5,6,2,6";
        }

        public Drawing CurrentDrawing { get; set; }
        public DynamicGeometry.Polygon Polygon { get; set; }
        public Point StartPoint { get; set; }
        public Point EndPoint { get; set; }
        public List<Point> Vertices { get; set; }
        public Segment Segment { get; set; }
        public DynamicGeometry.FreePoint Start { get; set; }
        public DynamicGeometry.FreePoint End { get; set; }
        public Route Route { get; set; }

        void canvas1_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (CurrentDrawing == null)
            {
                UpdateConfigurationFromText();
            }
        }

        void ActionManager_CollectionChanged(object sender, System.EventArgs e)
        {
            if (Start == null || End == null || Polygon == null || Route == null)
            {
                return;
            }
            text.Text = UpdateText();
        }

        string Display(double number)
        {
            return Math.Round(number, 2).ToString();
        }

        string UpdateText()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(Display(Start.Coordinates.X)).Append(",");
            sb.Append(Display(Start.Coordinates.Y)).Append(",");
            sb.Append(Display(End.Coordinates.X)).Append(",");
            sb.AppendLine(Display(End.Coordinates.Y));

            var vertices = Polygon.Dependencies.ToPoints().ToList();
            sb.AppendLine(vertices.Count.ToString());

            foreach (var vertex in vertices)
            {
                sb.Append(Display(vertex.X)).Append(",").Append(Display(vertex.Y)).Append(",");
            }
            sb.Remove(sb.Length - 1, 1);
            sb.AppendLine();
            try
            {
                var points = new List<Point>();
                Route.GetPoints(points);
                sb.AppendLine("Length: " + Display(points.PolylineLength()));
            }
            catch (System.Exception e)
            {
                sb.AppendLine(e.ToString());
            }

            return sb.ToString();
        }

        void Button_Click(object sender, RoutedEventArgs e)
        {
            UpdateConfigurationFromText();
        }

        private void UpdateConfigurationFromText()
        {
            var lines = text.Text.Split('\n', '\r')
                .Where(s => !string.IsNullOrEmpty(s)).ToArray();
            if (lines == null || lines.Length < 3)
            {
                return;
            }
            RoutingAlgorithm algorithm = new RoutingAlgorithm();
            algorithm.ParseInput(lines[0], lines[1], lines[2]);
            CreateConfiguration(algorithm.Start, algorithm.End, algorithm.Polygon);
            CurrentDrawing.Recalculate();
        }

        class IntegralDragger : Dragger
        {
            //protected override Point Coordinates(System.Windows.Input.MouseEventArgs e)
            //{
            //    var original = base.Coordinates(e);
            //    return new Point(Math.Round(original.X, 0), Math.Round(original.Y, 0));
            //}
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
            CurrentDrawing.CoordinateSystem.UnitLength = 24;
            CurrentDrawing.Behavior = new IntegralDragger();
            CurrentDrawing.ActionManager.CollectionChanged += ActionManager_CollectionChanged;
            CurrentDrawing.SelectionChanged += new System.EventHandler<Drawing.SelectionChangedEventArgs>(CurrentDrawing_SelectionChanged);

            var points = new List<FreePoint>();
            points.AddRange(
                from i in vertices
                select Factory.CreateFreePoint(CurrentDrawing, i));
            Actions.AddMany(CurrentDrawing, points.Cast<IFigure>());
            
            int j = 0;
            foreach (var p in points)
            {
                p.Name = (j++).ToString();
                p.ShowName = true;
            }
            
            Polygon = Factory.CreatePolygon(CurrentDrawing, points.Cast<IFigure>().ToArray());
            Polygon.Shape.Fill = new SolidColorBrush(Colors.Yellow);
            Actions.Add(CurrentDrawing, Polygon);

            Start = Factory.CreateFreePoint(CurrentDrawing, start);
            End = Factory.CreateFreePoint(CurrentDrawing, end);

            Start.Name = "s";
            Start.ShowName = true;
            End.Name = "e";
            End.ShowName = true;

            points = new List<FreePoint>()
                {
                    Start,
                    End
                };

            Actions.AddMany(CurrentDrawing, points.Cast<IFigure>());
            Segment = Factory.CreateSegment(CurrentDrawing, points.Cast<IFigure>().ToArray());
            Segment.Shape.Stroke = new SolidColorBrush(Colors.LightGray);
            Actions.Add(CurrentDrawing, Segment);

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

        void CurrentDrawing_SelectionChanged(object sender, Drawing.SelectionChangedEventArgs e)
        {
            var selection = e.SelectedFigures.Where(f => f != null && f.Visible).FirstOrDefault();
        }

        private void ZoomIn_Click(object sender, RoutedEventArgs e)
        {
            CurrentDrawing.CoordinateSystem.ZoomIn();
        }

        private void ZoomOut_Click(object sender, RoutedEventArgs e)
        {
            CurrentDrawing.CoordinateSystem.ZoomOut();
        }
    }
}
