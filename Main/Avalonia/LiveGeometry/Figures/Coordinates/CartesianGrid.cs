using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Media;
using GuiLabs.Undo;

namespace LiveGeometry;

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
        get => visible;
        set
        {
            visible = value;
            if (ShowAxes)
            {
                AxisLabels.Visible = value;
                XAxisLine.Visible = value;
                YAxisLine.Visible = value;
            }
            GridLines.Visible = value && showGridLines;
            if (value && this.Drawing != null)
            {
                UpdateVisual();
            }
        }
    }

    private bool showGridLines = true;
    /// <summary>
    /// The lines of a grid that shows, as <see cref="ShowAxes"/> is its axes. While the
    /// grid shows at least one of the two does (<see cref="SetShown"/>), and hidden, the
    /// two say what showing it brings back.
    /// </summary>
    [PropertyGridVisible]
    [PropertyGridName("Show Grid Lines")]
    public bool ShowGridLines
    {
        get => showGridLines;
        set
        {
            bool shown = value && visible;
            GridLines.Visible = shown;
            showGridLines = value;
            if (shown && this.Drawing != null)
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
        get => showAxes;
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

    /// <summary>Whether the axes are on screen, where a click can take them (<see cref="AxisLine"/>)</summary>
    public bool ShowsAxes => visible && showAxes;

    /// <summary>Whether the grid's lines are on screen</summary>
    public bool ShowsGridLines => visible && showGridLines;

    /// <summary>
    /// The axes and the lines shown or hidden one by one, as the Grid button's menu and
    /// the paper's context menu do, one undo step (outside a construction, whose step is
    /// its figure's alone). Hiding the last of the two hides the grid and leaves both
    /// on for the next time it is shown: a grid that shows nothing would be a Grid
    /// button down over an empty paper. Showing one while the grid is hidden shows the
    /// grid with that one alone.
    /// </summary>
    public void SetShown(bool axes, bool lines)
    {
        bool show = axes || lines;
        SetState(show, axes || !show, lines || !show);
    }

    /// <summary>
    /// The Grid button and G: the grid hidden, or shown with what it showed before (the
    /// axes alone stay the axes alone)
    /// </summary>
    public void Toggle()
    {
        if (visible)
        {
            SetState(visible: false, showAxes, showGridLines);
        }
        else if (showAxes || showGridLines)
        {
            SetState(visible: true, showAxes, showGridLines);
        }
        else
        {
            SetState(visible: true, axes: true, lines: true);
        }
    }

    void SetState(bool visible, bool axes, bool lines)
    {
        // the grid is the drawing's (a file says whether it shows), so this is an undo
        // step - except in the middle of a construction, whose step is its figure's alone
        if (Drawing.IsRecordingTransaction)
        {
            ShowAxes = axes;
            ShowGridLines = lines;
            Visible = visible;
            return;
        }

        var manager = Drawing.ActionManager;
        using (Transaction.Create(manager, false))
        {
            if (ShowAxes != axes)
            {
                manager.SetProperty(this, nameof(ShowAxes), axes);
            }

            if (ShowGridLines != lines)
            {
                manager.SetProperty(this, nameof(ShowGridLines), lines);
            }

            if (Visible != visible)
            {
                manager.SetProperty(this, nameof(Visible), visible);
            }
        }
    }

    public override IFigure HitTest(Point point, System.Predicate<IFigure> filter)
    {
        return null;
    }

    public static FrameworkElement GetIcon()
    {
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
    }

    public override string ToString()
    {
        return "Coordinate grid";
    }

    public IEnumerable<IValueProvider> GetProperties()
    {
        return PropertyDiscoveryStrategy.GetValuesFromProperties(this, "Visible", "Locked", "ShowAxes", "ShowGridLines");
    }

    public IEnumerable<IOperationDescription> GetMethods()
    {
        return [];
    }

    public override bool Serializable => false;
}
