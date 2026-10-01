using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Media;

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

        // the grid's own styles, in the theme's colors (not in the drawing's style list)
        readonly LineStyle axisStyle = new LineStyle()
        {
            Name = "AxisStyle",
            StrokeWidth = 1
        };

        readonly LineStyle gridStyle = new LineStyle()
        {
            Name = "GridStyle",
            StrokeWidth = 0.5
        };

        readonly LineStyle minorGridStyle = new LineStyle()
        {
            Name = "MinorGridStyle",
            StrokeWidth = 0.5
        };

        readonly TextStyle labelsStyle = new TextStyle()
        {
            FontSize = 12.0,
            Name = "LabelsStyle"
        };

        /// <summary>
        /// A color of the grid under the theme. The theme's colors are made for the theme's
        /// paper; on a solid paper of the drawing's own the grid takes those of the theme
        /// whose paper is nearest to it in lightness, shifted by as much as the paper differs
        /// from that theme's. So a GeoGebra worksheet, light gray under the dark theme, has
        /// the light theme's grid there, as much darker as its paper is.
        /// </summary>
        Color GetColor(AppTheme theme, System.Func<AppTheme, Color> color)
        {
            if (!(Drawing?.GetOwnBackground(theme) is SolidColorBrush paper))
            {
                return color(theme);
            }

            var nearest = AppTheme.All
                .OrderBy(candidate => System.Math.Abs(Lightness(candidate.Paper) - Lightness(paper.Color)))
                .First();
            var made = color(nearest);
            return Color.FromArgb(
                made.A,
                Shift(made.R, nearest.Paper.R, paper.Color.R),
                Shift(made.G, nearest.Paper.G, paper.Color.G),
                Shift(made.B, nearest.Paper.B, paper.Color.B));
        }

        static int Lightness(Color color)
        {
            return color.R + color.G + color.B;
        }

        static byte Shift(byte value, byte from, byte to)
        {
            return (byte)System.Math.Clamp(value + to - from, 0, 255);
        }

        public CartesianGrid()
        {
            axisStyle.BindToTheme(nameof(LineStyle.Color), theme => GetColor(theme, t => t.Axis));
            gridStyle.BindToTheme(nameof(LineStyle.Color), theme => GetColor(theme, t => t.GridMajor));
            minorGridStyle.BindToTheme(nameof(LineStyle.Color), theme => GetColor(theme, t => t.GridMinor));
            labelsStyle.BindToTheme(nameof(TextStyle.Color), theme => GetColor(theme, t => t.Axis));

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
            GridLines = new RectangularGridLinesCollection() { Drawing = Drawing, MinorStyle = minorGridStyle };

            XAxisLine.Arrow.Style = axisStyle;
            YAxisLine.Arrow.Style = axisStyle;
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

        /// <summary>A theme color was tweaked, or the paper is another: the grid's styles read their colors again</summary>
        public void RefreshTheme()
        {
            axisStyle.RefreshFromTheme();
            gridStyle.RefreshFromTheme();
            minorGridStyle.RefreshFromTheme();
            labelsStyle.RefreshFromTheme();
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
                // the axes of a grid that shows: ticked while the grid is hidden, they used
                // to come up on their own, and were gone when the file was opened again
                bool shown = value && visible;
                AxisLabels.Visible = shown;
                XAxisLine.Visible = shown;
                YAxisLine.Visible = shown;
                showAxes = value;
                if (shown && this.Drawing != null)
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
            var builder = IconBuilder.BuildIcon();
            for (double i = 0.1; i < 1; i += 0.2)
            {
                builder.TransparentLine(nameof(AppTheme.Guide), i, 0, i, 1, transparency: 0.4);
                builder.TransparentLine(nameof(AppTheme.Guide), 0, i, 1, i, transparency: 0.4);
            }

            const string axis = nameof(AppTheme.LineAccent);
            builder
                .Polygon(
                    axis,
                    axis,
                    new Point(0.5, 0),
                    new Point(0.4, 0.2),
                    new Point(0.6, 0.2))
                .Polygon(
                    axis,
                    axis,
                    new Point(1, 0.5),
                    new Point(0.8, 0.4),
                    new Point(0.8, 0.6))
                .Line(axis, 0.5, 0, 0.5, 1)
                .Line(axis, 0, 0.5, 1, 0.5);
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
