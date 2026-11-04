using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;

namespace LiveGeometry;

/// <summary>
/// The Figure List at the left of the canvas: every figure of the drawing, by the title the
/// property grid gives it, with the icon of the tool that makes it. Its selection is the
/// drawing's (click, Ctrl+click, Shift+click, the arrow keys); for the figure the keyboard is
/// on, arrows in the left margin point up to the figures it is built on, and with
/// `RecursiveArrows` gray ones from those on up to what they are built on, all the way.
/// Hidden figures (a square's helpers, an auxiliary number) are listed faded. While the
/// construction is stepped through (<see cref="StepPlayback"/>), a row goes to its step,
/// and the foot of the list edits the step and the description of the current step's
/// figures, or of the rows clicked (<see cref="PropertyGridSection.Steps"/>): the rows the
/// side panel leaves out.
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

    /// <summary>
    /// Over the first row of each step while stepping: the steps read as groups. Safe for
    /// the arrows, which place themselves by row index at <see cref="RowHeight"/>: they
    /// show for a selected figure only, and nothing is selected while stepping.
    /// </summary>
    const double StepGap = 10;

    readonly StackPanel rowPanel = new StackPanel();
    readonly Canvas arrows = new Canvas() { IsHitTestVisible = false };
    readonly ScrollViewer scrollViewer;
    readonly TextBlock header;

    /// <summary>At the foot while stepping: the Step and Description rows of the picked figures</summary>
    readonly PropertyGrid stepGrid = new PropertyGrid()
    {
        SectionFilter = section => section == PropertyGridSection.Steps
    };

    readonly Border stepFoot;

    /// <summary>At the right of the Figures line while stepping: shows and hides the foot</summary>
    readonly Button stepRowsToggle;

    /// <summary>Whether the foot is up while stepping (for the session; off, the toggle brings it back)</summary>
    bool stepRowsShown = true;

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

    /// <summary>
    /// While stepping, the figures whose step rows the foot edits: the current step's, or
    /// the rows clicked (click, Ctrl+click, Shift+click) until the step changes
    /// </summary>
    List<IFigure> stepPicks = new List<IFigure>();

    /// <summary>The step the foot's picks were taken from, -1 while not stepping</summary>
    int pickedStep = -1;

    /// <summary>A click on a row is taking the stepping to its step: the picks are the row's, not the step's</summary>
    bool pickingStep;

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

        // the toggle at the right of the Figures line, as the side panel's eye and padlock
        // are drawn: a pencil, plain while the foot is up, faint while it is away
        stepRowsToggle = new Button()
        {
            Width = 26,
            Height = 26,
            Padding = new Thickness(0),
            Margin = new Thickness(0, 3, 6, 0),
            Background = Brushes.Transparent,
            BorderThickness = new Thickness(0),
            HorizontalContentAlignment = HorizontalAlignment.Center,
            VerticalContentAlignment = VerticalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Top,
            IsVisible = false
        };
        stepRowsToggle.Click += (s, e) =>
        {
            stepRowsShown = !stepRowsShown;
            ShowStepRows();
        };
        var headerBar = new DockPanel();
        DockPanel.SetDock(stepRowsToggle, Dock.Right);
        headerBar.Children.Add(stepRowsToggle);
        headerBar.Children.Add(header);

        // the foot under the rows, over a line, with the side panel's own room at the sides
        stepFoot = new Border()
        {
            BorderThickness = new Thickness(0, 1, 0, 0),
            Padding = new Thickness(14, 8, 14, 12),
            IsVisible = false,
            Child = stepGrid
        };
        stepFoot.BindTheme(BorderBrushProperty, nameof(AppTheme.Separator));
        var dock = new DockPanel();
        DockPanel.SetDock(headerBar, Dock.Top);
        DockPanel.SetDock(stepFoot, Dock.Bottom);
        dock.Children.Add(headerBar);
        dock.Children.Add(stepFoot);
        dock.Children.Add(scrollViewer);
        Child = dock;

        // the keyboard's row is outlined only while the keys go to the list
        GotFocus += (s, e) => UpdateLook();
        LostFocus += (s, e) => UpdateLook();
    }

    public Drawing Drawing
    {
        get => drawing;
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
                drawing.PicksChanged -= Drawing_PicksChanged;
                drawing.Playback.Changed -= Playback_Changed;
            }

            drawing = value;
            focused = null;
            anchor = null;
            constructing = false;
            lastSelection.Clear();
            stepPicks.Clear();
            pickedStep = -1;
            ShowStepRows();
            if (drawing != null)
            {
                drawing.ActionManager.CollectionChanged += ActionManager_CollectionChanged;
                drawing.SelectionChanged += Drawing_SelectionChanged;
                drawing.ConstructionStepStarted += Drawing_ConstructionStepStarted;
                drawing.ConstructionStepComplete += Drawing_ConstructionStepComplete;
                drawing.PicksChanged += Drawing_PicksChanged;
                drawing.Playback.Changed += Playback_Changed;
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
        get => recursiveArrows;
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

    /// <summary>Whether the keyboard is in the foot's rows: its keys are the editor's, not the stepping's</summary>
    public bool IsEditingSteps => stepFoot.IsKeyboardFocusWithin;

    #region Following the drawing

    // A drag of a figure or of the view records a move at every step (they merge into one
    // undo step), and a refresh reads every row again: at once, then at most every 300 ms
    // while the history keeps changing, the last change always included (the titles of
    // shapes follow the drag: a quadrilateral becomes a square)
    void ActionManager_CollectionChanged(object sender, System.EventArgs e)
    {
        Throttle.Schedule(
            this,
            explorer => Dispatcher.UIThread.Post(explorer.RequestRefresh),
            System.TimeSpan.FromMilliseconds(300),
            ThrottleOptions.RunOnceImmediatelyIfFree);
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

    // the rows of the steps not yet taken fade, as hidden figures do, and the current
    // step's rows are painted as a selection is; the foot takes the step's figures at
    // every step, and lets go of them when the stepping stops (a change that leaves the
    // step where it is, an edit of a description, leaves the picks alone: the box being
    // typed in would be built again)
    void Playback_Changed()
    {
        var playback = drawing.Playback;
        if (!playback.IsOn)
        {
            if (pickedStep >= 0)
            {
                pickedStep = -1;
                stepPicks.Clear();
                ShowStepRows();
            }
        }
        else if (playback.Current != pickedStep)
        {
            pickedStep = playback.Current;
            if (!pickingStep)
            {
                stepPicks = playback.CurrentStep?.Figures.ToList() ?? new List<IFigure>();
                ShowStepRows();
            }
        }

        if (!IsVisible)
        {
            return;
        }

        foreach (var row in rows)
        {
            row.UpdateContent();
        }

        UpdateLook();

        // the step's rows in view, however the step was reached (the keys, the bar's
        // buttons, a row): the last first, then the first, so that the whole step shows
        // when it fits and its first row always. After the layout: the foot changes height
        // with the step's rows, and took the room the row had just been scrolled into.
        var figures = playback.CurrentStep?.Figures.ToList();
        if (figures != null && figures.Count > 0)
        {
            Dispatcher.UIThread.Post(() =>
            {
                ScrollIntoView(figures[figures.Count - 1]);
                ScrollIntoView(figures[0]);
            }, DispatcherPriority.Loaded);
        }
    }

    // a tool picked a figure (IFigurePicker): the picks are selected figures, painted as the selection
    void Drawing_PicksChanged(object sender, System.EventArgs e)
    {
        if (!IsVisible)
        {
            return;
        }

        lastSelection = SelectedFigures();
        UpdateLook();
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
        IFigure replacement = null;
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
                replacement = focused;
            }

            if (anchor != null && !rowsByFigure.ContainsKey(anchor))
            {
                anchor = focused;
            }

            // a picked figure deleted leaves the foot
            if (stepPicks.RemoveAll(figure => !rowsByFigure.ContainsKey(figure)) > 0)
            {
                ShowStepRows();
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

        // the selection deleted from the list with the keyboard, the row that took its
        // place is selected, as a list's next item is: the keyboard's row alone would be
        // outlined and empty. Not while picking (the selection is the tool's), nor while
        // stepping (nothing is selected then), nor when other figures stayed selected
        if (replacement != null
            && IsFocused
            && !(drawing.Behavior is IFigurePicker)
            && !drawing.Playback.IsOn
            && lastSelection.Count == 0)
        {
            Select(replacement, range: false, toggle: false);
        }
    }

    static bool IsListed(IFigure figure)
    {
        // a figure's name is part of the figure, the grid is part of the paper
        return !(figure is CartesianGrid) && !(figure is NameLabel);
    }

    HashSet<IFigure> SelectedFigures()
    {
        return new HashSet<IFigure>(drawing != null ? drawing.GetSelectedFigures() : []);
    }

    #endregion

    #region Selecting

    /// <param name="range">From the anchor to the figure (Shift)</param>
    /// <param name="toggle">Add to or take from the selection (Ctrl)</param>
    void Select(IFigure figure, bool range, bool toggle)
    {
        // stepping through the construction, a row takes the stepping to its step, which
        // is the selection, and the foot to its figure: with Ctrl or Shift, to several
        // figures, whose Step row is set together (their descriptions are each their own)
        if (drawing.Playback.IsOn)
        {
            focused = figure;
            if (range && anchor != null && rowsByFigure.ContainsKey(anchor))
            {
                int from = IndexOf(anchor);
                int to = IndexOf(figure);
                var picked = new List<IFigure>();
                for (int i = System.Math.Min(from, to); i <= System.Math.Max(from, to); i++)
                {
                    picked.Add(rows[i].Figure);
                }

                stepPicks = toggle ? stepPicks.Union(picked).ToList() : picked;
            }
            else if (toggle)
            {
                if (!stepPicks.Remove(figure))
                {
                    stepPicks.Add(figure);
                }

                anchor = figure;
            }
            else
            {
                stepPicks = new List<IFigure>() { figure };
                anchor = figure;
            }

            pickingStep = true;
            try
            {
                drawing.Playback.GoToFigure(figure);
            }
            finally
            {
                pickingStep = false;
            }

            ShowStepRows();
            UpdateLook();
            ScrollIntoView(figure);
            return;
        }

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

        // where it is on the canvas, which may be far from the row
        if (figure.Selected)
        {
            SelectionFlash.Show(figure);
        }
    }

    /// <summary>Without raising SelectionChanged: the new selection is raised once, whole</summary>
    void ClearSelection()
    {
        foreach (var figure in drawing.Figures)
        {
            if (FigureParts.HasSelection(figure))
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

        // an expression box has the keyboard, and keeps it: the row's figure goes into the
        // formula, hidden or not (ReferencePicker), the way a click on the figure would
        if (drawing.InsertTarget != null)
        {
            ReferencePicker.TryInsert(drawing.InsertTarget, row.Figure);
            e.Handled = true;
            return;
        }

        Focus();

        // a tool that picks figures takes the row as a click on its figure, hidden or not
        if (drawing.Behavior is IFigurePicker picker)
        {
            focused = row.Figure;
            picker.TogglePick(row.Figure);
            UpdateLook();
            e.Handled = true;
            return;
        }

        Select(
            row.Figure,
            range: e.KeyModifiers.HasFlag(KeyModifiers.Shift),
            toggle: (e.KeyModifiers & (KeyModifiers.Control | KeyModifiers.Meta)) != 0);
        e.Handled = true;
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        // the keys of the foot's text box and combo are theirs (an arrow in the
        // description must not go to the next step)
        if (IsEditingSteps)
        {
            base.OnKeyDown(e);
            return;
        }

        var picker = drawing?.Behavior as IFigurePicker;
        if (picker != null && e.Key == Key.Space && focused != null && rowsByFigure.ContainsKey(focused))
        {
            picker.TogglePick(focused);
            UpdateLook();
            e.Handled = true;
            return;
        }

        if (!IsNavigationKey(e.Key) || rows.Count == 0)
        {
            base.OnKeyDown(e);
            return;
        }

        // stepping through the construction, the keys walk the steps, not the rows
        var playback = drawing.Playback;
        if (playback.IsOn)
        {
            switch (e.Key)
            {
                case Key.Up: playback.Previous(); break;
                case Key.Down: playback.Next(); break;
                case Key.Home: playback.First(); break;
                default: playback.Last(); break;
            }

            e.Handled = true;
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

        // picking, the keys walk the rows and Space picks: the selection is the tool's
        if (picker != null)
        {
            focused = rows[target].Figure;
            UpdateLook();
            ScrollIntoView(focused);
        }
        else
        {
            Select(rows[target].Figure, range: e.KeyModifiers.HasFlag(KeyModifiers.Shift), toggle: false);
        }

        e.Handled = true;
    }

    #endregion

    #region The foot

    /// <summary>
    /// The foot shows the Step and Description rows of the picked figures: both for one
    /// figure, the Step row alone for several (<see cref="PropertyGridDisallowMultiEditAttribute"/>);
    /// nothing for figures that cannot be steps, or none. Called again after an undo or a
    /// redo (the host's <see cref="DrawingControl.UndoneOrRedone"/>), as the side panel is:
    /// a description put back raises nothing the rows would hear.
    /// </summary>
    public void ShowStepRows()
    {
        bool stepping = drawing?.Playback.IsOn == true;
        stepRowsToggle.IsVisible = stepping;
        if (stepping)
        {
            var icon = PropertyGridIcons.Create(PropertyGridIcon.Pencil);
            icon.Opacity = stepRowsShown ? 1 : 0.5;
            stepRowsToggle.Content = icon;
            ToolTip.SetTip(stepRowsToggle, stepRowsShown
                ? "Hide the Step and Description rows at the foot"
                : "Show the Step and Description rows at the foot");
        }

        var figures = stepPicks
            .Where(figure => figure is FigureBase figureBase && figureBase.CanBeStep())
            .Cast<object>()
            .ToList();
        if (drawing == null || figures.Count == 0 || !stepRowsShown)
        {
            stepGrid.Show(null, null);
        }
        else if (figures.Count == 1)
        {
            stepGrid.Show(figures[0], drawing.ActionManager);
        }
        else
        {
            stepGrid.Show(figures, drawing.ActionManager);
        }

        // the list is narrow: the description's box under its label, as wide as the foot
        foreach (var editor in (stepGrid.CurrentEditors ?? []).OfType<LabeledValueEditor>())
        {
            if (editor.Value?.GetAttribute<PropertyGridMultilineAttribute>() != null)
            {
                editor.PutLabelAbove();
            }
        }

        stepFoot.IsVisible = stepGrid.Selection != null;
    }

    #endregion

    #region Look

    /// <summary>Selection, the keyboard's row, the foot's figures and the arrows</summary>
    void UpdateLook()
    {
        bool stepping = drawing?.Playback.IsOn == true;
        foreach (var row in rows)
        {
            row.UpdateLook(
                isFocused: row.Figure == focused && IsFocused,
                isPicked: stepping && stepPicks.Contains(row.Figure));
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
            row.PlaceMargin();
        }
    }

    static double RowCenter(int index)
    {
        return index * RowHeight + RowHeight / 2;
    }

    #endregion

    class Row : Border
    {
        // the number of the step the figure begins, while the construction is stepped
        // through (StepPlayback): nothing for a figure grouped with the one before it
        readonly TextBlock number = new TextBlock()
        {
            Width = 20,
            Margin = new Thickness(0, 0, 4, 0),
            TextAlignment = TextAlignment.Right,
            VerticalAlignment = VerticalAlignment.Center,
            FontSize = 12,
            IsVisible = false
        };

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

        // the title in two runs: the kind as plain text, the name bold - what the list is
        // scanned for (the kind in the accent blue made a blue column of every row)
        readonly Run kindRun = new Run();
        readonly Run nameRun = new Run() { FontWeight = FontWeight.Bold };

        // how the figure is built, faded after the title: what a narrow list cuts first
        readonly TextBlock construction = new TextBlock()
        {
            VerticalAlignment = VerticalAlignment.Center,
            TextTrimming = TextTrimming.CharacterEllipsis,
            Margin = new Thickness(4, 0, 0, 0)
        };

        string iconKey;
        bool hover;
        bool isFocused;

        // whether the foot edits this figure's step (while stepping)
        bool isPicked;

        readonly FigureExplorer explorer;

        public Row(FigureExplorer explorer, IFigure figure)
        {
            this.explorer = explorer;
            Figure = figure;
            text.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Text));
            text.Inlines = new InlineCollection() { kindRun, nameRun };
            construction.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.TextMuted));
            Height = RowHeight;
            Margin = explorer.RowMargin;
            Padding = new Thickness(4, 0, 4, 0);
            BorderThickness = new Thickness(1);
            CornerRadius = new CornerRadius(3);

            var content = new DockPanel();
            DockPanel.SetDock(number, Dock.Left);
            DockPanel.SetDock(iconHost, Dock.Left);
            DockPanel.SetDock(text, Dock.Left);
            content.Children.Add(number);
            content.Children.Add(iconHost);
            content.Children.Add(text);
            content.Children.Add(construction);
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
            var (kind, name) = Figure is FigureBase figure ? figure.TitleParts : (null, Figure.Title);
            kindRun.Text = kind == null ? "" : name == null ? kind : kind + " ";
            nameRun.Text = name ?? "";
            construction.Text = Figure.Construction;
            // a helper made for another figure (the Number of a fixed length) counts as
            // hidden, and so does a figure of a step not yet taken
            bool shown = Figure.Visible && !Figure.Auxiliary && !(Figure is FigureBase figureBase && figureBase.IsHiddenByPlayback);
            Opacity = shown ? 1 : HiddenOpacity;
            var playback = Figure.Drawing?.Playback;
            number.IsVisible = playback?.IsOn == true;
            bool beginsStep = false;
            if (number.IsVisible)
            {
                int stepNumber = playback.NumberOf(Figure);
                number.Text = stepNumber > 0 ? stepNumber.ToString() : "";
                bool current = stepNumber == playback.Current + 1;
                number.FontWeight = current ? FontWeight.Bold : FontWeight.Normal;
                number.BindTheme(TextBlock.ForegroundProperty, current ? nameof(AppTheme.Accent) : nameof(AppTheme.TextMuted));
                beginsStep = stepNumber > 1;
            }

            // a gap over each step but the first, so the steps read as groups
            gap = beginsStep ? StepGap : 0;
            PlaceMargin();
            var key = FigureIcons.GetKey(Figure);
            if (key != iconKey || iconHost.Child == null && key != null)
            {
                iconKey = key;
                iconHost.Child = FigureIcons.Create(key, IconSize);
            }
        }

        public void UpdateLook(bool isFocused, bool isPicked)
        {
            this.isFocused = isFocused;
            this.isPicked = isPicked;
            Paint();
        }

        // the gap over the row (StepGap at the start of a step while stepping)
        double gap;

        /// <summary>The list's row margin (room for the arrows at the left), with the row's own gap over it</summary>
        public void PlaceMargin()
        {
            var margin = explorer.RowMargin;
            Margin = new Thickness(margin.Left, gap, margin.Right, margin.Bottom);
        }

        void Paint()
        {
            // while the construction is stepped through, the current step's rows (nothing
            // is selected then: the halos took the eye), and the figures the foot edits
            // are outlined as the keyboard's row is
            var playback = Figure.Drawing?.Playback;
            bool highlighted = Figure.Selected
                || playback?.IsOn == true && playback.StepOf(Figure) == playback.Current;
            string background = highlighted
                ? nameof(AppTheme.ButtonChecked)
                : hover ? nameof(AppTheme.ButtonHover) : null;
            this.BindTheme(BackgroundProperty, background, whenNone: Brushes.Transparent);
            this.BindTheme(BorderBrushProperty, isFocused || isPicked ? nameof(AppTheme.ButtonCheckedBorder) : null, whenNone: Brushes.Transparent);
        }
    }
}
