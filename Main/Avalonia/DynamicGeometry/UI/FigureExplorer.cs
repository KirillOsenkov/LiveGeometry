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
/// on, arrows in the left margin point up to the figures it is built on, and with
/// `RecursiveArrows` gray ones from those on up to what they are built on, all the way.
/// Hidden figures (a square's helpers, an auxiliary number) are listed faded.
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

    /// <summary>From a trunk to the arrowheads of its branches</summary>
    const double BranchLength = 4;
    const double ArrowHeadLength = 4;

    /// <summary>Between an arrowhead and the trunk to its right, or the row</summary>
    const double HeadGap = 2;
    const double TrunkSpacing = BranchLength + ArrowHeadLength + HeadGap;
    const double HiddenOpacity = 0.6;

    /// <summary>Between the edge of the window and the arrows</summary>
    const double LeftPadding = 8;

    readonly StackPanel rowPanel = new StackPanel();
    readonly Canvas arrows = new Canvas() { IsHitTestVisible = false };
    readonly ScrollViewer scrollViewer;
    readonly TextBlock header;

    List<Row> rows = new List<Row>();
    Dictionary<IFigure, Row> rowsByFigure = new Dictionary<IFigure, Row>();

    /// <summary>For each listed figure, the listed figures its arrows go to</summary>
    Dictionary<IFigure, List<IFigure>> arrowTargets = new Dictionary<IFigure, List<IFigure>>();

    /// <summary>The trunks side by side the margin has room for: as many as the figure with the most needs</summary>
    int trunkCount = 1;

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

        header = new TextBlock()
        {
            Text = "Figures",
            FontSize = 13,
            FontWeight = FontWeight.SemiBold,
            Margin = HeaderMargin
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

    bool recursiveArrows;

    /// <summary>
    /// Whether the arrows go on up from the figures the keyboard's figure is built on, in
    /// gray, to what those are built on, all the way; otherwise only its own arrows show
    /// </summary>
    public bool RecursiveArrows
    {
        get
        {
            return recursiveArrows;
        }
        set
        {
            if (recursiveArrows == value)
            {
                return;
            }

            recursiveArrows = value;
            UpdateTrunkCount();
            UpdateArrows();
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
            arrowTargets.Clear();
            SetTrunkCount(1);
            return;
        }

        // hidden, it is refreshed when shown; mid-construction, when the construction ends
        if (!IsVisible || constructing)
        {
            return;
        }

        var figures = drawing.Figures.Where(IsListed).ToList();
        bool listChanged = !figures.SequenceEqual(rows.Select(row => row.Figure));
        if (listChanged)
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

                row.Index = newRows.Count;
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

        // read every time: a figure can change what it is built on and stay where it is in the list
        arrowTargets = rows.ToDictionary(row => row.Figure, row => FindArrowTargets(row.Figure));
        if (listChanged)
        {
            UpdateTrunkCount();
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
        // turning to the list gives up a construction under way, as Escape does: its
        // transaction is open, and an edit of the selected figure would join the undo step
        // of the figure being made
        if (constructing)
        {
            drawing.Behavior?.Restart();
        }

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
        return rowsByFigure[figure].Index;
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
            toggle: (e.KeyModifiers & (KeyModifiers.Control | KeyModifiers.Meta)) != 0);
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
    /// From the figure the keyboard is on, one arrow up to each figure it is built on: out of
    /// its row into the margin, a trunk up, a short branch at the height of each row. With
    /// `RecursiveArrows`, from each of those that is built on something in turn the same in
    /// gray, and so on up. Every
    /// figure has a trunk of its own, a step left of the trunks above it that run beside it
    /// (`LayOutTrunks`), and its branches stop short of the next trunk to the right, so no two
    /// arrows cross. Nothing from below the figure: what is built on it would be a thicket.
    /// </summary>
    void UpdateArrows()
    {
        arrows.Children.Clear();
        if (focused == null || !focused.Selected || !rowsByFigure.ContainsKey(focused))
        {
            return;
        }

        var sources = FindArrowSources(focused);
        if (sources.Count == 0)
        {
            return;
        }

        var levels = LayOutTrunks(sources);

        // a figure that changed what it is built on may need more room than the list had
        if (CountLevels(levels) > trunkCount)
        {
            SetTrunkCount(CountLevels(levels));
        }

        double tip = ArrowMargin - 1;
        var trunks = levels.ToDictionary(pair => pair.Key, pair => TrunkX(pair.Value));

        // the figure's own arrows last, over the gray ones where they meet
        for (int i = sources.Count - 1; i >= 0; i--)
        {
            var source = sources[i];
            string color = i == 0 ? nameof(AppTheme.Accent) : nameof(AppTheme.TextFaint);
            double x = trunks[source];
            double origin = RowCenter(IndexOf(source));
            var targets = arrowTargets[source];
            double top = System.Math.Min(origin, targets.Min(target => RowCenter(IndexOf(target))));
            double bottom = System.Math.Max(origin, targets.Max(target => RowCenter(IndexOf(target))));
            AddLine(new Point(x, top), new Point(x, bottom), color);
            AddLine(new Point(x, origin), new Point(i == 0 ? tip - 2 : tip, origin), color);
            if (i == 0)
            {
                var dot = new Avalonia.Controls.Shapes.Ellipse()
                {
                    Width = 5,
                    Height = 5,
                    [Canvas.LeftProperty] = tip - 4.5,
                    [Canvas.TopProperty] = origin - 2.5
                };
                dot.BindTheme(Avalonia.Controls.Shapes.Shape.FillProperty, color);
                arrows.Children.Add(dot);
            }

            double headTip = x + BranchLength + ArrowHeadLength;
            foreach (var target in targets)
            {
                double y = RowCenter(IndexOf(target));
                AddLine(new Point(x, y), new Point(x + BranchLength, y), color);
                var head = new Avalonia.Controls.Shapes.Polygon()
                {
                    Points = new List<Point>()
                    {
                        new Point(headTip, y),
                        new Point(headTip - ArrowHeadLength, y - 3),
                        new Point(headTip - ArrowHeadLength, y + 3)
                    }
                };
                head.BindTheme(Avalonia.Controls.Shapes.Shape.FillProperty, color);
                arrows.Children.Add(head);
            }
        }

        void AddLine(Point start, Point end, string color)
        {
            var line = new Avalonia.Controls.Shapes.Line()
            {
                StartPoint = start,
                EndPoint = end,
                StrokeThickness = 1.5
            };
            line.BindTheme(Avalonia.Controls.Shapes.Shape.StrokeProperty, color);
            arrows.Children.Add(line);
        }
    }

    /// <summary>
    /// The figures that get arrows when the keyboard is on this one: the figure itself and,
    /// above it, everything it is built on all the way up - those of them built on something
    /// listed. Lowest first.
    /// </summary>
    List<IFigure> FindArrowSources(IFigure figure)
    {
        if (!RecursiveArrows)
        {
            return arrowTargets[figure].Count == 0 ? new List<IFigure>() : new List<IFigure>() { figure };
        }

        int bottom = IndexOf(figure);
        var sources = new List<IFigure>();
        var found = new HashSet<IFigure>() { figure };
        var pending = new Stack<IFigure>();
        pending.Push(figure);
        while (pending.Count > 0)
        {
            var source = pending.Pop();
            var targets = arrowTargets[source];
            if (targets.Count == 0)
            {
                continue;
            }

            sources.Add(source);
            foreach (var target in targets)
            {
                if (IndexOf(target) < bottom && found.Add(target))
                {
                    pending.Push(target);
                }
            }
        }

        return sources.OrderByDescending(IndexOf).ToList();
    }

    /// <summary>The listed figures this one is built on</summary>
    List<IFigure> FindArrowTargets(IFigure figure)
    {
        // a vertex or a side of a regular polygon is listed as its polygon
        return figure.Dependencies
            .Select(dependency => rowsByFigure.ContainsKey(dependency) ? dependency : drawing.Figures.FindTopLevel(dependency))
            .Where(dependency => dependency != null && dependency != figure && rowsByFigure.ContainsKey(dependency))
            .Distinct()
            .ToList();
    }

    /// <summary>
    /// The level of each figure's trunk, 0 next to the rows: one step left of the trunks of
    /// the figures above it that run beside its own (touching counts: two trunks in one column
    /// meeting in a row would read as one), and no further.
    /// </summary>
    Dictionary<IFigure, int> LayOutTrunks(List<IFigure> sources)
    {
        var levels = new Dictionary<IFigure, int>();
        var spans = new List<(IFigure Figure, int Top, int Bottom)>();

        // from the top down: a trunk is laid out after those above it
        for (int i = sources.Count - 1; i >= 0; i--)
        {
            var source = sources[i];
            int row = IndexOf(source);
            int top = System.Math.Min(row, arrowTargets[source].Min(IndexOf));
            int bottom = System.Math.Max(row, arrowTargets[source].Max(IndexOf));
            int level = 0;
            foreach (var span in spans)
            {
                if (span.Top <= bottom && top <= span.Bottom)
                {
                    level = System.Math.Max(level, levels[span.Figure] + 1);
                }
            }

            levels[source] = level;
            spans.Add((source, top, bottom));
        }

        return levels;
    }

    static int CountLevels(Dictionary<IFigure, int> levels)
    {
        return levels.Count == 0 ? 1 : levels.Values.Max() + 1;
    }

    /// <summary>A trunk's place: at level 0 its arrowheads end at the rows, a step further left per level</summary>
    double TrunkX(int level)
    {
        return ArrowMargin - 1 - BranchLength - ArrowHeadLength - TrunkSpacing * level;
    }

    /// <summary>Room for the trunks and a pixel either side</summary>
    double ArrowMargin => 2 + BranchLength + ArrowHeadLength + TrunkSpacing * (trunkCount - 1);

    Thickness HeaderMargin => new Thickness(LeftPadding + ArrowMargin + 4, 8, 8, 6);

    Thickness RowMargin => new Thickness(ArrowMargin, 0, 4, 0);

    /// <summary>Room for as many trunks as the figure with the most needs</summary>
    void UpdateTrunkCount()
    {
        SetTrunkCount(rows.Count == 0 ? 1 : rows.Max(row => CountLevels(LayOutTrunks(FindArrowSources(row.Figure)))));
    }

    void SetTrunkCount(int count)
    {
        count = System.Math.Max(count, 1);
        if (count == trunkCount)
        {
            return;
        }

        trunkCount = count;
        header.Margin = HeaderMargin;
        foreach (var row in rows)
        {
            row.Margin = RowMargin;
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
            Margin = explorer.RowMargin;
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

        /// <summary>The row's place in the list</summary>
        public int Index { get; set; }

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
