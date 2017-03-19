using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Media;

namespace DynamicGeometry
{
    public partial class CartesianGrid : CompositeFigure, ICustomPropertyProvider, ICustomMethodProvider
    {
        public PointByCoordinates OriginPoint { get; set; }
        public PointByCoordinates XUnitPoint { get; set; }
        public PointByCoordinates YUnitPoint { get; set; }
        public Axis XAxisLine { get; set; }
        public Axis YAxisLine { get; set; }
        AxisLabelsCollection AxisLabels { get; set; }
        GridLinesCollection GridLines { get; set; }

        public CartesianGrid()
        {
            //ShapeStyle arrowStyle = new ShapeStyle()
            //{
            //    Color = Color.FromArgb(255, 128, 128, 255),
            //    Fill = new SolidColorBrush(Color.FromArgb(255, 128, 128, 255)),
            //    StrokeWidth = 1,
            //    Name = "ArrowStyle"
            //};
            LineStyle axisStyle = new LineStyle()
            {
                Color = Color.FromArgb(255, 128, 128, 255),
                Name = "AxisStyle",
                StrokeWidth = 1
            };
            LineStyle gridStyle = new LineStyle()
            {
                Color = Colors.LightGray,
                Name = "GridStyle",
                StrokeWidth = 0.5
            };
            TextStyle labelsStyle = new TextStyle()
            {
                Color = Color.FromArgb(255, 128, 128, 255),
                FontSize = 12.0,
                Name = "LabelsStyle"
            };

            OriginPoint = Factory.CreatePointByCoordinates(Drawing, () => 0, () => 0);
            XUnitPoint = Factory.CreatePointByCoordinates(Drawing, () => 1, () => 0);
            YUnitPoint = Factory.CreatePointByCoordinates(Drawing, () => 0, () => 1);
            OriginPoint.Name = "Origin";
            XUnitPoint.Name = "XUnitPoint";
            YUnitPoint.Name = "YUnitPoint";
            OriginPoint.Visible = false;
            XUnitPoint.Visible = false;
            YUnitPoint.Visible = false;

            XAxisLine = Factory.CreateAxis(Drawing, new[] { OriginPoint, XUnitPoint });
            YAxisLine = Factory.CreateAxis(Drawing, new[] { OriginPoint, YUnitPoint });
            XAxisLine.Name = "XAxisLine";
            YAxisLine.Name = "YAxisLine";
            AxisLabels = new AxisLabelsCollection() { Drawing = Drawing };
            GridLines = new RectangularGridLinesCollection() { Drawing = Drawing };

            //XAxisLine.Arrow.Style = arrowStyle;
            XAxisLine.Line.Style = axisStyle;
            //YAxisLine.Arrow.Style = arrowStyle;
            YAxisLine.Line.Style = axisStyle;
            GridLines.Style = gridStyle;
            AxisLabels.Style = labelsStyle;

            Children.Add(
                OriginPoint,
                XUnitPoint,
                YUnitPoint,
                XAxisLine,
                YAxisLine,
                AxisLabels,
                GridLines
                );
        }

        private bool visible = false;
        public override bool Visible
        {
            get
            {
                return visible;
            }
            set
            {
                visible = value;
                if (ShowAxes)
                {
                    AxisLabels.Visible = value;
                    XAxisLine.Visible = value;
                    YAxisLine.Visible = value;
                }
                GridLines.Visible = value;
                Settings.Instance.ShowGrid = value;
                if (value && this.Drawing != null)
                {
                    UpdateVisual();
                }
            }
        }

        private bool showAxes = true;
        [PropertyGridVisible]
        [PropertyGridName("Show Axes")]
        public bool ShowAxes
        {
            get
            {
                return showAxes;
            }
            set
            {
                AxisLabels.Visible = value;
                XAxisLine.Visible = value;
                YAxisLine.Visible = value;
                showAxes = value;
                if (value && this.Drawing != null)
                {
                    UpdateVisual();
                }
            }
        }

        public override IFigure HitTest(Point point, System.Predicate<IFigure> filter)
        {
            return null;
        }

        public static FrameworkElement GetIcon()
        {
#if TABULAPLAYER
            return null;    // never used and my player excludes IconBuilder
#else
            var builder = IconBuilder
                .BuildIcon()
                .Polygon(
                    new SolidColorBrush(Colors.Blue),
                    new SolidColorBrush(Colors.Blue),
                    new Point(0.5, 0),
                    new Point(0.4, 0.2),
                    new Point(0.6, 0.2))
                .Polygon(
                    new SolidColorBrush(Colors.Blue),
                    new SolidColorBrush(Colors.Blue),
                    new Point(1, 0.5),
                    new Point(0.8, 0.4),
                    new Point(0.8, 0.6))
                .Line(Color.FromArgb(255, 0, 0, 255), 0.5, 0, 0.5, 1)
                .Line(Color.FromArgb(255, 0, 0, 255), 0, 0.5, 1, 0.5);
            for (double i = 0.1; i < 1; i += 0.2)
            {
                builder.Line(Color.FromArgb(100, 0, 0, 255), i, 0, i, 1);
                builder.Line(Color.FromArgb(100, 0, 0, 255), 0, i, 1, i);
            }
            return builder.Canvas;
#endif
        }

        public override string ToString()
        {
            return "Coordinate grid";
        }

        public IEnumerable<IValueProvider> GetProperties()
        {
            return PropertyDiscoveryStrategy.GetValuesFromProperties(this, "Visible", "Locked", "ShowAxes");
        }

        public IEnumerable<IOperationDescription> GetMethods()
        {
            return Enumerable.Empty<IOperationDescription>();
        }

        public override bool Serializable
        {
            get
            {
                return false;
            }
        }
    }
}
