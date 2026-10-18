using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Input;
using GuiLabs.Undo;

namespace LiveGeometry;

/// <summary>
/// The right click's menus under the Drag tool: one for the paper, one for a figure or
/// the selection it is in. Every menu is made of the same sections in the same order,
/// with a separator between two that have items (<see cref="Section"/>): which figure,
/// what it shows, what it can do, its place among the others, hidden and locked, then
/// Copy and, always last, Delete. What a figure shows and can do is the property grid's
/// own (rows and buttons marked <see cref="ContextMenuAttribute"/>), asked the way the
/// grid asks, so the menu offers a verb exactly when the grid does; the rest are the
/// same for every figure. The grid is where everything is; the menu has what is handy.
/// </summary>
public class ContextMenus
{
    /// <summary>The sections of a menu, top to bottom</summary>
    public enum Section
    {
        /// <summary>Which of the figures under the cursor the menu is about</summary>
        Choose,
        /// <summary>The paper's: the view, the axes and the grid's lines</summary>
        View,
        /// <summary>What the figure shows of itself: check items (Show name)</summary>
        Show,
        /// <summary>What the figure's kind can do (Free point, Fix length, Convert to)</summary>
        Verbs,
        /// <summary>Bring to front, Send to back</summary>
        Arrange,
        /// <summary>Hide, Lock</summary>
        State,
        /// <summary>Copy, Paste, Select all; Delete, always the last item of a figure's menu</summary>
        Edit,
        /// <summary>The paper's: the drawing's own properties</summary>
        Drawing
    }

    readonly Drawing drawing;
    readonly SortedDictionary<Section, List<Control>> sections = new SortedDictionary<Section, List<Control>>();

    ContextMenus(Drawing drawing)
    {
        this.drawing = drawing;
    }

    #region Building

    MenuItem Add(
        Section section,
        string header,
        Action action,
        bool? isChecked = null,
        PropertyGridIcon? icon = null,
        KeyGesture gesture = null)
    {
        var item = CreateItem(header, action, isChecked, icon, gesture);
        Add(section, item);
        return item;
    }

    void Add(Section section, Control item)
    {
        if (!sections.TryGetValue(section, out var items))
        {
            items = new List<Control>();
            sections.Add(section, items);
        }

        items.Add(item);
    }

    static MenuItem CreateItem(
        string header,
        Action action,
        bool? isChecked = null,
        PropertyGridIcon? icon = null,
        KeyGesture gesture = null)
    {
        var item = new MenuItem() { Header = header, InputGesture = gesture };
        if (isChecked != null)
        {
            item.ToggleType = MenuItemToggleType.CheckBox;
            item.IsChecked = isChecked.Value;
        }
        else if (icon != null)
        {
            item.Icon = PropertyGridIcons.Create(icon.Value);
        }

        if (action != null)
        {
            item.Click += (s, e) => action();
        }

        return item;
    }

    ContextMenu Build()
    {
        var menu = new ContextMenu();
        foreach (var items in sections.Values.Where(items => items.Count > 0))
        {
            if (menu.Items.Count > 0)
            {
                menu.Items.Add(new Separator());
            }

            foreach (var item in items)
            {
                menu.Items.Add(item);
            }
        }

        return menu;
    }

    /// <summary>Ctrl with the key, or Cmd on a Mac, as the shortcuts are handled (<see cref="KeyNames.IsMac"/>)</summary>
    static KeyGesture Command(Key key)
    {
        return new KeyGesture(key, KeyNames.IsMac ? KeyModifiers.Meta : KeyModifiers.Control);
    }

    #endregion

    #region The paper

    /// <summary>The paper's menu: the view, Paste and Select all, the drawing's background</summary>
    public static ContextMenu ForPaper(Drawing drawing)
    {
        var menu = new ContextMenus(drawing);
        menu.Add(Section.View, "Zoom to fit", drawing.ZoomToFit, gesture: new KeyGesture(Key.Home));
        foreach (var item in GridItems(drawing))
        {
            menu.Add(Section.View, item);
        }

        menu.Add(Section.Edit, "Paste", drawing.Paste, gesture: Command(Key.V));
        menu.Add(Section.Edit, "Select all", () =>
        {
            drawing.SelectAll();
            drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());
        }, gesture: Command(Key.A));
        menu.Add(Section.Drawing, "Background…", () => drawing.RaiseDisplayProperties(drawing), icon: PropertyGridIcon.Swatches);
        return menu.Build();
    }

    /// <summary>
    /// The axes and the grid's lines, each on or off (<see cref="CartesianGrid.SetShown"/>):
    /// in the paper's menu and under the Grid button's arrow
    /// </summary>
    public static IEnumerable<MenuItem> GridItems(Drawing drawing)
    {
        var grid = drawing.CoordinateGrid;
        bool axes = grid.ShowsAxes;
        bool lines = grid.ShowsGridLines;
        yield return CreateItem("Axes", () => SetGrid(grid, !axes, lines), isChecked: axes);
        yield return CreateItem("Grid lines", () => SetGrid(grid, axes, !lines), isChecked: lines);
    }

    static void SetGrid(CartesianGrid grid, bool axes, bool lines)
    {
        grid.SetShown(axes, lines);
        CommandToolButton.UpdateToggles();
    }

    #endregion

    #region Figures

    /// <summary>
    /// The menu of a figure, which the right click has selected, or of the selection it is
    /// part of: the selection's own rows and verbs, and the common ones for all of them
    /// </summary>
    /// <param name="figure">The figure of the drawing clicked (a path, for a click on its side)</param>
    /// <param name="choices">The figures under the cursor, for Choose figure; one or none for no choice</param>
    public static ContextMenu ForFigure(Drawing drawing, IFigure figure, IList<IFigure> choices)
    {
        var menu = new ContextMenus(drawing);
        menu.AddChoices(choices);

        // A name is its figure's: what it offers is the figure's Show name (and for a
        // point Show coordinates), which hides it too; nothing else is the name's own.
        if (figure is PointLabel or FigureLabel)
        {
            if (figure.Dependencies.FirstOrDefault() is IFigure owner)
            {
                menu.AddRows(owner);
            }

            return menu.Build();
        }

        // (a side of a Bezier path is selected by itself, but what the menu offers is the
        // path's: Closed, Filled, as Hide and Delete take the path)
        var figures = menu.SelectedOrClicked(figure);
        object subject = figures.Count == 1
            ? figure
            : new FigureSelection(drawing, figures, ExcludeByDefaultValueDiscoveryStrategy.Instance);
        menu.AddRows(subject);
        menu.AddVerbs(subject, figures.Count == 1 ? figure : null);
        menu.AddArrange(figures);
        menu.AddState(figure, figures);
        menu.Add(Section.Edit, "Copy", drawing.Copy, icon: PropertyGridIcon.Copy, gesture: Command(Key.C));
        menu.Add(Section.Edit, "Delete", drawing.DeleteSelection, icon: PropertyGridIcon.Delete, gesture: new KeyGesture(Key.Delete));
        return menu.Build();
    }

    /// <summary>
    /// Where figures overlap, "Choose figure" (as in the original DG) lists them all and
    /// selects the one picked: the click selects the topmost only
    /// </summary>
    void AddChoices(IList<IFigure> choices)
    {
        if (choices == null || choices.Count < 2)
        {
            return;
        }

        var choose = new MenuItem() { Header = "Choose figure" };
        foreach (var candidate in choices)
        {
            choose.Items.Add(CreateItem(ClickChoice.Describe(candidate), () =>
            {
                drawing.Figures.ClearSelection();
                candidate.Selected = true;
                drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());
            }));
        }

        Add(Section.Choose, choose);
    }

    /// <summary>The figures selected with the one clicked, which is among them; no names of points or lines (those go with Show name)</summary>
    IList<IFigure> SelectedOrClicked(IFigure clicked)
    {
        var selected = FigureParts.Wholes(drawing.GetSelectedFigures())
            .Where(f => !(f is PointLabel) && !(f is FigureLabel))
            .ToList();
        if (!selected.Contains(clicked))
        {
            selected.Add(clicked);
        }

        return selected;
    }

    /// <summary>
    /// The on/off rows of the subject's page that the menu offers (Show name), as check
    /// items: for several figures, those they all have, ticked when all are on, and a
    /// click sets them all, one undo step, as the grid's check box does
    /// </summary>
    void AddRows(object subject)
    {
        var strategy = ValueDiscoveryStrategy.Get(subject.GetType()) ?? ExcludeByDefaultValueDiscoveryStrategy.Instance;
        foreach (var row in strategy.GetValues(subject) ?? Enumerable.Empty<IValueProvider>())
        {
            if (row.Type != typeof(bool)
                || row.GetAttribute<ContextMenuAttribute>() == null
                || !row.CanSetValue
                || !RowShows(row))
            {
                continue;
            }

            bool isChecked = row.GetValue<object>() is true;
            Add(Section.Show, row.DisplayName, () => Actions.SetProperty(drawing.ActionManager, row, !isChecked), isChecked);
        }
    }

    /// <summary>Whether the grid shows the row for every figure it stands for (<see cref="PropertyGridConditionAttribute"/>)</summary>
    static bool RowShows(IValueProvider row)
    {
        var condition = row.GetAttribute<PropertyGridConditionAttribute>();
        if (condition == null)
        {
            return true;
        }

        var owners = row is CompositeValueProvider composite
            ? composite.InnerList.Select(inner => inner.Parent)
            : new[] { row.Parent };
        return owners.All(condition.Holds);
    }

    /// <summary>
    /// The buttons of the subject's page that the menu offers, in the grid's order; the
    /// "Convert to" ones in a submenu of their own when there are two or more, as the
    /// kinds a figure can become are one choice
    /// </summary>
    /// <param name="figure">The one figure the menu is about, for what only a menu can offer (Snap to among several, Edit text); null for several</param>
    void AddVerbs(object subject, IFigure figure)
    {
        if (figure is PointBase point)
        {
            AddSnapChoices(point);
        }

        if (figure is LabelBase label && InPlaceTextEditor.CanEdit(label))
        {
            // after the menu has closed, which gives the keyboard back to the canvas
            Add(Section.Verbs, "Edit text", () => Avalonia.Threading.Dispatcher.UIThread.Post(() => InPlaceTextEditor.Open(label)), icon: PropertyGridIcon.Pencil);
        }

        const string convertTo = "Convert to ";
        var verbs = PropertyGrid.GetCallableMethods(subject)
            .Where(verb => verb.GetAttribute<ContextMenuAttribute>() != null
                && verb.Parameters.IsEmpty()
                && (subject is not IConditionalProperties conditions || conditions.CanEdit(verb.Name)))
            .ToList();
        var conversions = verbs.Where(verb => verb.DisplayName.StartsWith(convertTo, StringComparison.Ordinal)).ToList();
        if (conversions.Count < 2)
        {
            conversions.Clear();
        }

        foreach (var verb in verbs.Except(conversions))
        {
            Add(Section.Verbs, CreateVerbItem(verb, verb.DisplayName, subject));
        }

        if (conversions.Count > 0)
        {
            var submenu = new MenuItem() { Header = "Convert to" };
            foreach (var verb in conversions)
            {
                var kind = verb.DisplayName.Substring(convertTo.Length);
                submenu.Items.Add(CreateVerbItem(verb, char.ToUpperInvariant(kind[0]) + kind.Substring(1), subject));
            }

            Add(Section.Verbs, submenu);
        }
    }

    static MenuItem CreateVerbItem(IOperationDescription verb, string header, object subject)
    {
        return CreateItem(
            header,
            () => verb.Invoke(subject, Array.Empty<object>()),
            icon: verb.GetAttribute<PropertyGridIconAttribute>()?.Icon);
    }

    /// <summary>
    /// "Snap to" among the figures through a free point, when there are several: the grid
    /// offers the one figure there is (<see cref="FreePoint.SnapToFigure"/>)
    /// </summary>
    void AddSnapChoices(PointBase point)
    {
        var figures = point is FreePoint ? PointSnapping.FiguresToSnapTo(point) : new IFigure[0];
        if (figures.Count < 2)
        {
            return;
        }

        var snapTo = new MenuItem() { Header = "Snap to" };
        foreach (var target in figures)
        {
            snapTo.Items.Add(CreateItem(PointSnapping.Describe(target), () => PointSnapping.SnapTo(point, target)));
        }

        Add(Section.Verbs, snapTo);
    }

    /// <summary>The selection's place among the figures it overlaps (<see cref="ZOrders"/>)</summary>
    void AddArrange(IList<IFigure> figures)
    {
        if (ZOrders.CanBringToFront(figures))
        {
            Add(Section.Arrange, "Bring to front", () => ZOrders.BringToFront(figures), icon: PropertyGridIcon.BringToFront);
        }

        if (ZOrders.CanSendToBack(figures))
        {
            Add(Section.Arrange, "Send to back", () => ZOrders.SendToBack(figures), icon: PropertyGridIcon.SendToBack);
        }
    }

    /// <summary>Hide and Lock, for the whole selection as Delete takes it</summary>
    void AddState(IFigure figure, IList<IFigure> figures)
    {
        Add(Section.State, "Hide", () =>
        {
            // hidden, the one clicked alone went and the rest stayed selected
            var hidden = figures.Where(f => f.Visible).ToArray();
            using (Transaction.Create(drawing.ActionManager, false))
            {
                foreach (var item in hidden)
                {
                    SetProperty(item, nameof(IFigure.Visible), false);
                }
            }

            foreach (var item in figures)
            {
                item.Selected = false;
            }

            drawing.RaiseSelectionChanged(drawing.GetSelectedFigures());
        });

        if (figure is ControlBase fixedLabel && drawing.FixedLabels.Contains(fixedLabel))
        {
            // a label a gallery drawing came with is held as if locked: unlocked, it
            // drags as any label does, the caption on its own
            Add(Section.State, "Unlock", () =>
            {
                foreach (var item in figures.OfType<ControlBase>())
                {
                    drawing.FixedLabels.Remove(item);
                }
            }, icon: PropertyGridIcon.Unlock);
            return;
        }

        bool locked = !figure.Locked;
        Add(Section.State, locked ? "Lock" : "Unlock", () =>
        {
            using (Transaction.Create(drawing.ActionManager, false))
            {
                foreach (var item in figures.Where(f => f.Locked != locked))
                {
                    SetProperty(item, nameof(IFigure.Locked), locked);
                }
            }
        }, icon: locked ? PropertyGridIcon.Lock : PropertyGridIcon.Unlock);
    }

    void SetProperty(object target, string property, object value)
    {
        Actions.SetProperty(drawing.ActionManager, new PropertyValue(property, target), value);
    }

    #endregion
}
