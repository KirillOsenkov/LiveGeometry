using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using Avalonia;
using Avalonia.Input;
using GuiLabs.Undo;

namespace DynamicGeometry;

/// <summary>
/// Makes a show/hide box (<see cref="ShowHideControl"/>), or changes what one shows and hides
/// (<see cref="Edit"/>, the box's Edit figures). The figures selected when the tool is picked
/// are picked already; a click on a figure picks it or lets it go, and so does a click on its
/// row in the Figure List, the way to the hidden ones (<see cref="IFigurePicker"/>). A click
/// on the paper puts the box there and is done. The box starts ticked when what it holds is
/// on screen, unticked when it is hidden: nothing changes when the box appears. Editing, a
/// click on the paper, on the box, Enter or OK is done; Escape gives up either.
/// </summary>
[Category(BehaviorCategories.Misc)]
[Order(4)]
public class ShowHideCreator : Behavior, IFigurePicker
{
    // the figures picked, in the order they were picked
    readonly List<IFigure> picked = new List<IFigure>();

    // the box whose figures are picked again; null while a new box is made
    ShowHideControl editedBox;

    // the box the next start of the tool edits (Edit sets it, Started takes it)
    ShowHideControl boxToEdit;

    PickPanel panel;

    /// <summary>Picks again what the box shows and hides: the tool, started on the box</summary>
    public static void Edit(ShowHideControl box)
    {
        var drawing = box.Drawing;
        if (drawing == null)
        {
            return;
        }

        // the ribbon's instance, whose button it lights up
        var tool = FindTool<ShowHideCreator>() ?? new ShowHideCreator();
        if (drawing.Behavior == tool)
        {
            drawing.SetDefaultBehavior();
        }

        tool.boxToEdit = box;
        drawing.Behavior = tool;
    }

    public override void Started()
    {
        editedBox = boxToEdit;
        boxToEdit = null;
        picked.Clear();

        // a box's own figures as they are (one from an old file may hold what the tool would
        // not take, and keeps it); for a new box, what was selected and a box can hold
        var start = editedBox != null
            ? editedBox.Dependencies.ToList()
            : Drawing.GetSelectedFigures().Select(Pickable).Where(figure => figure != null).Distinct().ToList();
        Drawing.Figures.ClearSelection();
        foreach (var figure in start)
        {
            picked.Add(figure);
            figure.Selected = true;
        }

        panel = editedBox != null ? new EditPanel(this) : new PickPanel(this);
        Drawing.RaisePicksChanged();
    }

    public override void Stopping()
    {
        picked.Clear();
        editedBox = null;
        panel = null;
        if (Drawing != null)
        {
            Drawing.Figures.ClearSelection();
            Drawing.RaisePicksChanged();
        }

        base.Stopping();
    }

    public override object PropertyBag
    {
        get
        {
            if (panel == null)
            {
                panel = new PickPanel(this);
            }

            return panel;
        }
    }

    /// <summary>
    /// What a click on the figure picks: the figure, the one a name belongs to, the polygon
    /// of a vertex; null for what a box can't hold - another box, an axis, a Number (nothing
    /// to show), a helper that goes with the figure it was made for
    /// </summary>
    IFigure Pickable(IFigure figure)
    {
        if (figure is PointLabel || figure is FigureLabel)
        {
            figure = figure.Dependencies.FirstOrDefault();
        }

        if (figure == null)
        {
            return null;
        }

        figure = Drawing.Figures.FindTopLevel(figure);
        if (figure == null
            || figure is ShowHideControl
            || figure is AxisLine
            || figure is CartesianGrid
            || figure is Number
            || figure.Auxiliary)
        {
            return null;
        }

        return figure;
    }

    public void TogglePick(IFigure figure)
    {
        if (picked.Remove(figure))
        {
            figure.Selected = false;
        }
        else
        {
            var pickable = Pickable(figure);
            if (pickable == null)
            {
                Drawing.RaiseStatusNotification("A show/hide box can't hold that.");
                return;
            }

            if (picked.Contains(pickable))
            {
                return;
            }

            picked.Add(pickable);
            pickable.Selected = true;
        }

        panel?.Refresh();
        Drawing.RaisePicksChanged();
    }

    public override void MouseDown(object sender, MouseButtonEventArgs e)
    {
        var coordinates = Coordinates(e, false, false, false);
        var figure = FindFigureToToggle(coordinates);
        if (figure != null)
        {
            TogglePick(figure);
            return;
        }

        // (a press on a box is the tool's, not the box's: ShowHideCheckBox)
        var box = Drawing.Figures.HitTest(coordinates) as ShowHideControl;
        if (editedBox != null)
        {
            if (box == null || box == editedBox)
            {
                FinishEditing();
            }

            return;
        }

        if (box != null)
        {
            Drawing.RaiseStatusNotification("Click on the paper where there is room for the box.");
            return;
        }

        PlaceBox(Coordinates(e));
    }

    /// <summary>The figure a click here picks or lets go of: the first there is, or the one chosen with Tab</summary>
    IFigure FindFigureToToggle(Point coordinates)
    {
        return Choice.Pick(FindFiguresToToggle(coordinates));
    }

    IReadOnlyList<IFigure> FindFiguresToToggle(Point coordinates)
    {
        return Drawing.Figures
            .HitTestAll(coordinates, figure => figure.Visible && figure.IsHitTestVisible && Pickable(figure) != null)
            .Select(Pickable)
            .Distinct()
            .ToList();
    }

    protected override IReadOnlyList<object> FindClickOptions(MouseEventArgs e)
    {
        return FindFiguresToToggle(Coordinates(e, false, false, false)).ToList<object>();
    }

    /// <summary>A halo on the figure a click would pick or let go of</summary>
    protected override IFigure GetFigureToPick(MouseEventArgs e)
    {
        return FindFigureToToggle(Coordinates(e, false, false, false));
    }

    protected override Cursor GetCursor(Point coordinates)
    {
        return FindFigureToToggle(coordinates) != null ? HandCursor : ArrowCursor;
    }

    public override void KeyDown(object sender, KeyEventArgs e)
    {
        // (not Delete: what is picked is selected, and would go)
        if (e.Key == Key.Escape)
        {
            AbortAndSetDefaultTool();
            e.Handled = true;
        }
        else if (e.Key == Key.Enter && editedBox != null)
        {
            FinishEditing();
            e.Handled = true;
        }
    }

    void PlaceBox(Point coordinates)
    {
        var drawing = Drawing;
        if (picked.Count == 0)
        {
            drawing.RaiseStatusNotification("First click the figures the box should show and hide.");
            return;
        }

        // in the drawing's order, as a file lists them
        var figures = drawing.Figures.Where(picked.Contains).ToList();
        bool shown = figures.Any(figure => figure.Visible);
        var box = new ShowHideControl() { Drawing = drawing, Dependencies = figures };
        box.SetBox(shown);
        box.Text = shown ? "Show" : "Hint";
        box.MoveTo(coordinates);
        Actions.Add(drawing, box);
        AbortAndSetDefaultTool();
        ShowBox(drawing, box, focusCaption: true);
    }

    /// <summary>What the box holds becomes what is picked: one undo step</summary>
    public void FinishEditing()
    {
        var drawing = Drawing;
        var box = editedBox;
        if (box == null)
        {
            return;
        }

        if (picked.Count == 0)
        {
            drawing.RaiseStatusNotification("A box holds one figure at least. To remove the box, delete it.");
            return;
        }

        var removed = box.Dependencies.Where(figure => !picked.Contains(figure)).ToList();
        var added = drawing.Figures.Where(figure => picked.Contains(figure) && !box.Dependencies.Contains(figure)).ToList();
        if (removed.Count > 0 || added.Count > 0)
        {
            using (Transaction.Create(drawing.ActionManager, false))
            {
                foreach (var figure in removed)
                {
                    Actions.RemoveDependency(box, figure);
                }

                foreach (var figure in added)
                {
                    Actions.InsertDependency(box, box.Dependencies.Count, figure);
                }

                MoveAfter(drawing, box, added);
            }
        }

        AbortAndSetDefaultTool();
        ShowBox(drawing, box, focusCaption: false);
    }

    /// <summary>
    /// The box to the end of the drawing's list when it holds a figure that came after it:
    /// the list is in dependency order, which the file is read in
    /// </summary>
    static void MoveAfter(Drawing drawing, ShowHideControl box, List<IFigure> added)
    {
        var figures = drawing.Figures;
        int from = figures.IndexOf(box);
        if (!added.Any(figure => figures.IndexOf(figure) > from))
        {
            return;
        }

        int to = figures.Count - 1;
        drawing.ActionManager.RecordAction(new CallMethodAction(
            () => figures.Move(from, to),
            () => figures.Move(to, from)));
    }

    /// <summary>The box selected, its properties in the side panel</summary>
    static void ShowBox(Drawing drawing, ShowHideControl box, bool focusCaption)
    {
        drawing.Figures.ClearSelection();
        box.Selected = true;
        drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());
        if (focusCaption)
        {
            drawing.RaiseDisplayProperties(box, focusProperty: nameof(ShowHideControl.Text));
        }
    }

    public override string Name
    {
        get { return "Show/hide box"; }
    }

    public override string HintText
    {
        get
        {
            return editedBox != null
                ? "Click figures to add them to the box or take them out (hidden ones in the Figure List), then click the paper."
                : "Click the figures the box shows and hides (hidden ones in the Figure List), then click where the box goes.";
        }
    }

    public override FrameworkElement CreateIcon()
    {
        return IconBuilder.BuildIcon()
            .Polygon(
                nameof(AppTheme.ShapeIconFill),
                nameof(AppTheme.Ink),
                new Point(0.1, 0.32),
                new Point(0.46, 0.32),
                new Point(0.46, 0.68),
                new Point(0.1, 0.68))
            .Polyline(
                strokeThickness: 2,
                nameof(AppTheme.LineAccent),
                new[] { new Point(0.17, 0.5), new Point(0.25, 0.59), new Point(0.39, 0.4) })
            .Line(nameof(AppTheme.Ink), 0.58, 0.5, 0.9, 0.5)
            .Canvas;
    }

    /// <summary>The tool's panel: what is picked so far</summary>
    [PropertyGridName("Show/hide box")]
    public class PickPanel : ToolPanel, INotifyPropertyChanged
    {
        public PickPanel(ShowHideCreator tool)
        {
            Tool = tool;
        }

        protected ShowHideCreator Tool { get; }

        [PropertyGridVisible]
        [PropertyGridName("Shows and hides")]
        public string Figures
        {
            get
            {
                return ShowHideControl.Describe(Tool.picked);
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public void Refresh()
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Figures)));
        }
    }

    /// <summary>The panel while a box's figures are picked again: OK is done</summary>
    [PropertyGridName("Figures of the box")]
    public class EditPanel : PickPanel
    {
        public EditPanel(ShowHideCreator tool)
            : base(tool)
        {
        }

        [PropertyGridVisible]
        [PropertyGridIcon(PropertyGridIcon.Check)]
        public void OK()
        {
            Tool.FinishEditing();
        }
    }
}
