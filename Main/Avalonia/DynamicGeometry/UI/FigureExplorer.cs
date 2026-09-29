using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;

namespace DynamicGeometry;

/// <summary>
/// The Figure List at the left of the canvas: every figure of the drawing, by the title the
/// property grid gives it, with the icon of the tool that makes it. Its selection is the
/// drawing's (click, Ctrl+click, Shift+click, the arrow keys); for the figure the keyboard is
/// on, arrows in the left margin point up to the figures it is built on. Hidden figures (a
/// square's helpers, an auxiliary number) are listed faded.
/// </summary>
/// <remarks>
/// The list follows the undo history rather than the figure collection: a construction in
/// progress adds its preview and temporary points without recording them, and even its
/// recorded steps stay inside its transaction until the figure is made or given up. So the
/// list is rebuilt after each change of the history (coalesced: dragging the plane records a
/// step per mouse move), never while a construction is under way.
/// </remarks>
public class FigureExplorer : Border
{
    const double RowHeight = 22;
    const double IconSize = 16;

    /// <summary>Room at the left of the rows for the arrows to the dependencies</summary>
    const double ArrowMargin = 16;
    const double TrunkX = 5;
    const double ArrowHeadLength = 4;
    const double HiddenOpacity = 0.6;

    /// <summary>Between the edge of the window and the arrows</summary>
    const double LeftPadding = 8;

    readonly StackPanel rowPanel = new StackPanel();
    readonly Canvas arrows = new Canvas() { IsHitTestVisible = false };
    readonly ScrollViewer scrollViewer;

    List<Row> rows = new List<Row>();
    Dictionary<IFigure, Row> rowsByFigure = new Dictionary<IFigure, Row>();

    /// <summary>The selection as it was last seen, to tell which figures a click on the canvas added</summary>
    HashSet<IFigure> lastSelection = new HashSet<IFigure>();

    Drawing drawing;

    /// <summary>The row the keyboard is on, whose dependencies the arrows show</summary>
    IFigure focused;

    /// <summary>Where a Shift range starts</summary>
    IFigure anchor;

    bool refreshPosted;
    bool constructing;

    /// <summary>The list itself is changing the selection</summary>
    bool selecting;

    public FigureExplorer()
    {
        Focusable = true;
        this.BindTheme(BackgroundProperty, nameof(AppTheme.Background));

        var content = new Panel() { Margin = new Thickness(LeftPadding, 0, 0, 0) };
        content.Children.Add(rowPanel);
        content.Children.Add(arrows);
        scrollViewer = new ScrollViewer()
        {
            HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
            VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            Content = content
        };

        var header = new TextBlock()
        {
            Text = "Figures",
            FontSize = 13,
            FontWeight = FontWeight.SemiBold,
            Margin = new Thickness(LeftPadding + ArrowMargin + 4, 8, 8, 6)
        };
        header.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Text));
        var dock = new DockPanel();
        DockPanel.SetDock(header, Dock.Top);
        dock.Children.Add(header);
        dock.Children.Add(scrollViewer);
        Child = dock;

        // the keyboard's row is outlined only while the keys go to the list
        GotFocus += (s, e) => UpdateLook();
        LostFocus += (s, e) => UpdateLook();
    }

    public Drawing Drawing
    {
        get
        {
            return drawing;
        }
        set
        {
            if (drawing == value)
            {
                return;
            }

            if (drawing != null)
            {
                drawing.ActionManager.CollectionChanged -= ActionManager_CollectionChanged;
                drawing.SelectionChanged -= Drawing_SelectionChanged;
                drawing.ConstructionStepStarted -= Drawing_ConstructionStepStarted;
                drawing.ConstructionStepComplete -= Drawing_ConstructionStepComplete;
            }

            drawing = value;
            focused = null;
            anchor = null;
            constructing = false;
            lastSelection.Clear();
            if (drawing != null)
            {
                drawing.ActionManager.CollectionChanged += ActionManager_CollectionChanged;
                drawing.SelectionChanged += Drawing_SelectionChanged;
                drawing.ConstructionStepStarted += Drawing_ConstructionStepStarted;
                drawing.ConstructionStepComplete += Drawing_ConstructionStepComplete;
            }

            // a drawing is attached empty and loaded right after
            Refresh();
            RequestRefresh();
        }
    }

    /// <summary>Whether a key is one the list moves by (and not one that pans the canvas)</summary>
    public static bool IsNavigationKey(Key key)
    {
        return key == Key.Up || key == Key.Down || key == Key.Home || key == Key.End;
    }

    #region Following the drawing

    void ActionManager_CollectionChanged(object sender, System.EventArgs e)
    {
        RequestRefresh();
    }

    void Drawing_ConstructionStepStarted(object sender, Drawing.ConstructionStepStartedEventArgs e)
    {
        constructing = true;
    }

    void Drawing_ConstructionStepComplete(object sender, Drawing.ConstructionStepCompleteEventArgs e)
    {
        if (e.ConstructionComplete)
        {
            constructing = false;
            RequestRefresh();
        }
    }

    void Drawing_SelectionChanged(object sender, Drawing.SelectionChangedEventArgs e)
    {
        if (selecting || !IsVisible)
        {
            return;
        }

        // the figure clicked last on the canvas is the one to follow
        var listed = rows.Select(row => row.Figure);
        var added = listed.LastOrDefault(figure => figure.Selected && !lastSelection.Contains(figure));
        if (added != null)
        {
            focused = added;
        }
        else if (focused == null || !focused.Selected)
        {
            focused = listed.FirstOrDefault(figure => figure.Selected) ?? focused;
        }

        anchor = focused;
        lastSelection = SelectedFigures();
        UpdateLook();
        if (added != null)
        {
            ScrollIntoView(focused);
        }
    }

    void RequestRefresh()
    {
        if (refreshPosted)
        {
            return;
        }

        refreshPosted = true;
        Dispatcher.UIThread.Post(() =>
        {
            refreshPosted = false;
            Refresh();
        }, DispatcherPriority.Background);
    }

    /// <summary>The rows read the drawing again: which figures there are, their titles, icons, visibility</summary>
    public void Refresh()
    {
        if (drawing == null)
        {
            rows.Clear();
            rowsByFigure.Clear();
            rowPanel.Children.Clear();
            arrows.Children.Clear();
            return;
        }

        // hidden, it is refreshed when shown; mid-construction, when the construction ends
        if (!IsVisible || constructing)
        {
            return;
        }

        var figures = drawing.Figures.Where(IsListed).ToList();
        if (!figures.SequenceEqual(rows.Select(row => row.Figure)))
        {
            int focusedIndex = focused != null && rowsByFigure.TryGetValue(focused, out var focusedRow) ? rows.IndexOf(focusedRow) : -1;
            var newRows = new List<Row>(figures.Count);
            var newRowsByFigure = new Dictionary<IFigure, Row>(figures.Count);
            foreach (var figure in figures)
            {
                if (!rowsByFigure.TryGetValue(figure, out var row))
                {
                    row = new Row(this, figure);
                }

                newRows.Add(row);
                newRowsByFigure.Add(figure, row);
            }

            rows = newRows;
            rowsByFigure = newRowsByFigure;
            rowPanel.Children.Clear();
            rowPanel.Children.AddRange(rows);

            // a deleted figure hands the keyboard to the one that took its place
            if (focused != null && !rowsByFigure.ContainsKey(focused))
            {
                focused = rows.Count == 0 ? null : rows[System.Math.Clamp(focusedIndex, 0, rows.Count - 1)].Figure;
            }

            if (anchor != null && !rowsByFigure.ContainsKey(anchor))
            {
                anchor = focused;
            }
        }

        foreach (var row in rows)
        {
            row.UpdateContent();
        }

        lastSelection = SelectedFigures();
        UpdateLook();
    }

    static bool IsListed(IFigure figure)
    {
        // a figure's name is part of the figure, the grid is part of the paper
        return !(figure is CartesianGrid) && !(figure is PointLabel) && !(figure is FigureLabel);
    }

    HashSet<IFigure> SelectedFigures()
    {
        return new HashSet<IFigure>(drawing != null ? drawing.GetSelectedFigures() : Enumerable.Empty<IFigure>());
    }

    #endregion

    #region Selecting

    /// <param name="range">From the anchor to the figure (Shift)</param>
    /// <param name="toggle">Add to or take from the selection (Ctrl)</param>
    void Select(IFigure figure, bool range, bool toggle)
    {
        selecting = true;
        try
        {
            if (range && anchor != null && rowsByFigure.ContainsKey(anchor))
            {
                if (!toggle)
                {
                    ClearSelection();
                }

                int from = IndexOf(anchor);
                int to = IndexOf(figure);
                for (int i = System.Math.Min(from, to); i <= System.Math.Max(from, to); i++)
                {
                    rows[i].Figure.Selected = true;
                }
            }
            else if (toggle)
            {
                figure.Selected = !figure.Selected;
                anchor = figure;
            }
            else
            {
                ClearSelection();
                figure.Selected = true;
                anchor = figure;
            }

            focused = figure;
            drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());
        }
        finally
        {
            selecting = false;
        }

        lastSelection = SelectedFigures();
        UpdateLook();
        ScrollIntoView(figure);
    }

    /// <summary>Without raising SelectionChanged: the new selection is raised once, whole</summary>
    void ClearSelection()
    {
        foreach (var figure in drawing.Figures)
        {
            if (figure.Selected)
            {
                figure.Selected = false;
            }
        }
    }

    int IndexOf(IFigure figure)
    {
        return rows.IndexOf(rowsByFigure[figure]);
    }

    void ScrollIntoView(IFigure figure)
    {
        if (figure != null && rowsByFigure.TryGetValue(figure, out var row))
        {
            row.BringIntoView();
        }
    }

    void Row_PointerPressed(Row row, PointerPressedEventArgs e)
    {
        if (!e.GetCurrentPoint(row).Properties.IsLeftButtonPressed)
        {
            return;
        }

        Focus();
        Select(
            row.Figure,
            range: e.KeyModifiers.HasFlag(KeyModifiers.Shift),
            toggle: e.KeyModifiers.HasFlag(KeyModifiers.Control));
        e.Handled = true;
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        if (!IsNavigationKey(e.Key) || rows.Count == 0)
        {
            base.OnKeyDown(e);
            return;
        }

        int index = focused != null && rowsByFigure.ContainsKey(focused) ? IndexOf(focused) : -1;
        int target = e.Key switch
        {
            Key.Up => index < 0 ? rows.Count - 1 : index - 1,
            Key.Down => index + 1,
            Key.Home => 0,
            _ => rows.Count - 1
        };
        target = System.Math.Clamp(target, 0, rows.Count - 1);
        Select(rows[target].Figure, range: e.KeyModifiers.HasFlag(KeyModifiers.Shift), toggle: false);
        e.Handled = true;
    }

    #endregion

    #region Look

    /// <summary>Selection, the keyboard's row and the arrows</summary>
    void UpdateLook()
    {
        foreach (var row in rows)
        {
            row.UpdateLook(isFocused: row.Figure == focused && IsFocused);
        }

        UpdateArrows();
    }

    /// <summary>
    /// From the figure the keyboard is on, one arrow up to each figure it is built on: a trunk
    /// down the margin, a branch into each row. Only for that one figure, or the margin would
    /// be a thicket.
    /// </summary>
    void UpdateArrows()
    {
        arrows.Children.Clear();
        if (focused == null || !focused.Selected || !rowsByFigure.ContainsKey(focused))
        {
            return;
        }

        var targets = focused.Dependencies
            .Distinct()
            .Where(rowsByFigure.ContainsKey)
            .Select(dependency => RowCenter(IndexOf(dependency)))
            .ToList();
        if (targets.Count == 0)
        {
            return;
        }

        double origin = RowCenter(IndexOf(focused));
        double top = System.Math.Min(origin, targets.Min());
        double bottom = System.Math.Max(origin, targets.Max());
        double tip = ArrowMargin - 1;

        AddLine(new Point(TrunkX, top), new Point(TrunkX, bottom));
        AddLine(new Point(TrunkX, origin), new Point(tip - 2, origin));
        var dot = new Avalonia.Controls.Shapes.Ellipse()
        {
            Width = 5,
            Height = 5,
            [Canvas.LeftProperty] = tip - 4.5,
            [Canvas.TopProperty] = origin - 2.5
        };
        dot.BindTheme(Avalonia.Controls.Shapes.Shape.FillProperty, nameof(AppTheme.Accent));
        arrows.Children.Add(dot);

        foreach (var y in targets)
        {
            AddLine(new Point(TrunkX, y), new Point(tip - ArrowHeadLength, y));
            var head = new Avalonia.Controls.Shapes.Polygon()
            {
                Points = new List<Point>()
                {
                    new Point(tip, y),
                    new Point(tip - ArrowHeadLength, y - 3),
                    new Point(tip - ArrowHeadLength, y + 3)
                }
            };
            head.BindTheme(Avalonia.Controls.Shapes.Shape.FillProperty, nameof(AppTheme.Accent));
            arrows.Children.Add(head);
        }

        void AddLine(Point start, Point end)
        {
            var line = new Avalonia.Controls.Shapes.Line()
            {
                StartPoint = start,
                EndPoint = end,
                StrokeThickness = 1.5
            };
            line.BindTheme(Avalonia.Controls.Shapes.Shape.StrokeProperty, nameof(AppTheme.Accent));
            arrows.Children.Add(line);
        }
    }

    static double RowCenter(int index)
    {
        return index * RowHeight + RowHeight / 2;
    }

    #endregion

    class Row : Border
    {
        readonly Decorator iconHost = new Decorator()
        {
            Width = IconSize,
            Height = IconSize,
            Margin = new Thickness(0, 0, 6, 0),
            VerticalAlignment = VerticalAlignment.Center
        };

        readonly TextBlock text = new TextBlock()
        {
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis
        };

        string iconKey;
        bool hover;
        bool isFocused;

        public Row(FigureExplorer explorer, IFigure figure)
        {
            Figure = figure;
            text.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Text));
            Height = RowHeight;
            Margin = new Thickness(ArrowMargin, 0, 4, 0);
            Padding = new Thickness(4, 0, 4, 0);
            BorderThickness = new Thickness(1);
            CornerRadius = new CornerRadius(3);

            var content = new DockPanel();
            DockPanel.SetDock(iconHost, Dock.Left);
            content.Children.Add(iconHost);
            content.Children.Add(text);
            Child = content;

            PointerPressed += (s, e) => explorer.Row_PointerPressed(this, e);
            PointerEntered += (s, e) =>
            {
                hover = true;
                Paint();
            };
            PointerExited += (s, e) =>
            {
                hover = false;
                Paint();
            };
            UpdateContent();
            Paint();
        }

        public IFigure Figure { get; }

        public void UpdateContent()
        {
            text.Text = Figure.Title;
            // a helper made for another figure (the Number of a fixed length) counts as hidden
            Opacity = Figure.Visible && !Figure.Auxiliary ? 1 : HiddenOpacity;
            var key = FigureIcons.GetKey(Figure);
            if (key != iconKey || iconHost.Child == null && key != null)
            {
                iconKey = key;
                iconHost.Child = FigureIcons.Create(key, IconSize);
            }
        }

        public void UpdateLook(bool isFocused)
        {
            this.isFocused = isFocused;
            Paint();
        }

        void Paint()
        {
            string background = Figure.Selected
                ? nameof(AppTheme.ButtonChecked)
                : hover ? nameof(AppTheme.ButtonHover) : null;
            this.BindTheme(BackgroundProperty, background, whenNone: Brushes.Transparent);
            this.BindTheme(BorderBrushProperty, isFocused ? nameof(AppTheme.ButtonCheckedBorder) : null, whenNone: Brushes.Transparent);
        }
    }
}
