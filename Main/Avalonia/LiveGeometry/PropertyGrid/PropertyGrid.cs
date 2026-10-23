using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using GuiLabs.Undo;

namespace LiveGeometry;

public partial class PropertyGrid : StackPanel
{
    public PropertyGrid()
    {
        ValueDiscoveryStrategy = new IncludeByDefaultValueDiscoveryStrategy();
        PropertyGridTheme.Apply(this);
        Grid.SetIsSharedSizeScope(this, true);
    }

    public IValueDiscoveryStrategy ValueDiscoveryStrategy { get; set; }

    private object mSelection;
    public object Selection
    {
        get => mSelection;
        protected set
        {
            if (mSelection != null)
            {
                UnsubscribeFromPropertyChangeNotifications(mSelection);
            }
            mSelection = value;
            IPropertyGridContentProvider customContent = mSelection as IPropertyGridContentProvider;
            if (customContent != null)
            {
                mSelection = customContent.GetContentForPropertyGrid();
            }
            if (mSelection != null)
            {
                SubscribeToPropertyChangeNotifications(value);
            }

            // the rows first: whoever hears of the change measures the page (the side panel
            // goes where it covers least of what it is about)
            UpdateContents();
            UpdateVisibility(mSelection != null);
        }
    }

    public event EventHandler VisibilityChanged;
    public void UpdateVisibility(bool visible)
    {
        this.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
        if (VisibilityChanged != null)
        {
            VisibilityChanged(this, null);
        }
    }

    void SubscribeToPropertyChangeNotifications(object instance)
    {
        INotifyPropertyChanged objectWithEvent = instance as INotifyPropertyChanged;
        if (objectWithEvent != null)
        {
            objectWithEvent.PropertyChanged += this.SelectionPropertyChanged;
        }
        IPropertyGridHost supportsHost = instance as IPropertyGridHost;
        if (supportsHost != null)
        {
            supportsHost.PropertyGrid = this;
        }

        if (instance is ToolPanel panel)
        {
            panel.PropertyError += SelectionPropertyError;
            panel.FocusRequested += SelectionFocusRequested;
        }
    }

    void UnsubscribeFromPropertyChangeNotifications(object instance)
    {
        INotifyPropertyChanged objectWithEvent = instance as INotifyPropertyChanged;
        if (objectWithEvent != null)
        {
            objectWithEvent.PropertyChanged -= SelectionPropertyChanged;
        }
        IPropertyGridHost supportsHost = instance as IPropertyGridHost;
        if (supportsHost != null)
        {
            supportsHost.PropertyGrid = null;
        }

        if (instance is ToolPanel panel)
        {
            panel.PropertyError -= SelectionPropertyError;
            panel.FocusRequested -= SelectionFocusRequested;
        }
    }

    void SelectionPropertyChanged(object sender, PropertyChangedEventArgs e)
    {
        // no name: the rows changed shape (what is editable, a caption, which buttons -
        // a translated point freed), so everything is built again; posted, since this
        // may come from inside an editor's own set
        if (string.IsNullOrEmpty(e.PropertyName))
        {
            Dispatcher.UIThread.Post(() =>
            {
                if (Selection == sender)
                {
                    UpdateContents();
                }
            });
            return;
        }

        FindAndUpdatePropertyEditor(e.PropertyName);
        UpdateLiveButtons();
        FollowObjectTab();

        // the title may say what changed: the name ("Segment AB", also when a point is
        // renamed), the number of sides of a regular polygon ("Regular pentagon")
        UpdateHeader();
    }

    /// <summary>What a tool panel's command found wrong, under the row it is about</summary>
    void SelectionPropertyError(string propertyName, string error)
    {
        var editor = CurrentEditors?
            .OfType<StringEditor>()
            .FirstOrDefault(e => e.Value?.Name == propertyName);
        if (editor != null)
        {
            editor.ErrorText = error;
        }
    }

    /// <summary>A tool panel's command wants the keyboard in one of its rows (Add point: back to X)</summary>
    void SelectionFocusRequested(string propertyName)
    {
        var editor = CurrentEditors?
            .OfType<LabeledValueEditor>()
            .FirstOrDefault(e => e.Value?.Name == propertyName);
        editor?.TakeFocus();
    }

    void FindAndUpdatePropertyEditor(string propertyName)
    {
        if (CurrentProperties == null)
        {
            return;
        }
        foreach (var value in CurrentProperties)
        {
            if (value.Name == propertyName)
            {
                value.RaiseValueChanged();
                return;
            }
        }
    }

    /// <summary>
    /// Figures moved (a drag, undo): the numbers shown are read again, since some follow
    /// from where things are (a segment's length, a circle's radius) and raise no change
    /// of their own. Only numbers - a move changes no text, color or choice - and never
    /// the box being typed in.
    /// </summary>
    public void RefreshNumbers()
    {
        // posted from a throttle (DrawingHost): by now the grid may show nothing
        if (Selection == null || CurrentEditors == null)
        {
            return;
        }

        // what a quadrilateral is called depends on where its vertices are
        UpdateHeader();

        foreach (var editor in CurrentEditors)
        {
            if (editor is IValueEditor valueEditor
                && valueEditor.Value?.Type == typeof(double)
                && !editor.IsKeyboardFocusWithin)
            {
                valueEditor.Value.RaiseValueChanged();
            }

            // the value under an expression box follows the figures it names
            if (editor is StringEditor stringEditor)
            {
                stringEditor.RefreshLiveValue();
            }

            // a figure's lengths, area and angles, and whether each is on the drawing
            if (editor is QuantitiesEditor quantities)
            {
                quantities.Refresh();
            }
        }
    }

    private void Show(object newSelection)
    {
        Selection = newSelection;
    }

    /// <summary>Whose history the rows record their sets in; null for none</summary>
    public virtual ActionManager ActionManager { get; set; }

    /// <param name="focusProperty">
    /// The property whose editor takes the keyboard, this once (a new label's text), besides
    /// those marked <see cref="PropertyGridFocusAttribute"/>; null for none
    /// </param>
    public void Show(object newSelection, ActionManager actionManager, string focusProperty = null)
    {
        ActionManager = PropertyGridNoUndoAttribute.IsOn(newSelection) ? null : actionManager;
        Selection = newSelection;
        if (focusProperty != null && CurrentEditors != null)
        {
            var editor = CurrentEditors
                .OfType<LabeledValueEditor>()
                .FirstOrDefault(e => e.Value?.Name == focusProperty);
            if (editor != null)
            {
                editor.FocusWhenLoaded = true;
            }
        }
    }

    public void Show(IEnumerable<object> objects, ActionManager actionManager)
    {
        if (objects == null || objects.Count() == 0)
        {
            Selection = null;
            return;
        }
        ActionManager = actionManager;
        var aggregate = new CompositePropertyProvider(this.ValueDiscoveryStrategy, objects);
        var properties = aggregate.GetProperties();
        if (properties.IsEmpty())
        {
            aggregate = null;
        }
        Selection = aggregate;
    }

    protected bool mExpanded = true;
    public bool Expanded
    {
        get => mExpanded;
        set
        {
            if (mExpanded == value)
            {
                return;
            }
            mExpanded = value;
            UpdateContents();
        }
    }

    void UpdateContents()
    {
        this.Children.Clear();
        if (HeaderHost != null)
        {
            HeaderHost.Child = null;
        }

        // the editors of what was shown before must not be refreshed after it's gone (a
        // refresh posted on a move can arrive after the grid was emptied or changed)
        CurrentProperties = null;
        CurrentEditors = null;
        if (hoistedValues != null)
        {
            hoistedValues.PropertyChanged -= HoistedValuesPropertyChanged;
            hoistedValues = null;
        }
        tabSwitcher = null;
        tabPages.Clear();
        if (Selection == null)
        {
            return;
        }
        AddHeader();
        if (!Expanded)
        {
            return;
        }
        AddChildren();
    }

    protected virtual void AddChildren()
    {
        var controls = CreateObjectControls(Selection);
        if (controls == null)
        {
            return;
        }

        ShowHeaderRows();

        if (Selection is IPropertyGridTabs tabs)
        {
            AddTabs(tabs, controls);
            return;
        }

        foreach (var control in Arrange(controls))
        {
            this.Children.Add(control);
        }
    }

    SegmentSwitcher tabSwitcher;
    readonly Dictionary<string, Control> tabPages = new Dictionary<string, Control>();

    // the object's own tab as last seen: when it changes (undo, a pick), the grid follows
    string shownObjectTab;

    /// <summary>
    /// The strip of tabs, the page of the current one, then what is on every tab (OK)
    /// </summary>
    void AddTabs(IPropertyGridTabs tabs, IEnumerable<UIElement> controls)
    {
        tabSwitcher = new SegmentSwitcher()
        {
            Margin = new Thickness(0, 0, 0, 8)
        };
        tabPages.Clear();
        var pages = new Panel();

        // a row on more than one tab gets an editor of its own on each
        var placed = new HashSet<UIElement>();
        var copies = new List<UIElement>();
        foreach (var tab in tabs.Tabs)
        {
            var page = new StackPanel();
            var members = new List<UIElement>();
            foreach (var control in controls.Where(c => tabs.GetTabs(GetMemberName(c)).Contains(tab)))
            {
                if (placed.Add(control))
                {
                    members.Add(control);
                    continue;
                }

                var copy = CreateCopy(control);
                if (copy != null)
                {
                    copies.Add(copy);
                    members.Add(copy);
                }
            }

            foreach (var control in Arrange(members))
            {
                page.Children.Add(control);
            }

            tabPages[tab] = page;
            pages.Children.Add(page);
            tabSwitcher.Add(tab, tab);
        }

        tabSwitcher.Selected += tab =>
        {
            ShowTab((string)tab);
            tabs.OnTabSelected((string)tab, ActionManager);
            shownObjectTab = tabs.CurrentTab;
        };

        this.Children.Add(tabSwitcher);
        this.Children.Add(pages);
        foreach (var control in Arrange(controls.Where(c => tabs.GetTabs(GetMemberName(c)).Count == 0)))
        {
            this.Children.Add(control);
        }

        CurrentEditors = CurrentEditors.Concat(copies).ToArray();
        shownObjectTab = tabs.CurrentTab;
        ShowTab(shownObjectTab);
    }

    void ShowTab(string tab)
    {
        tabSwitcher.Current = tab;
        foreach (var page in tabPages)
        {
            page.Value.IsVisible = page.Key == tab;
        }
    }

    /// <summary>The object changed: if what it is now belongs on another tab, show that one</summary>
    void FollowObjectTab()
    {
        if (Selection is not IPropertyGridTabs tabs || tabSwitcher == null)
        {
            return;
        }

        var tab = tabs.CurrentTab;
        if (tab != shownObjectTab)
        {
            shownObjectTab = tab;
            ShowTab(tab);
        }
    }

    /// <summary>Another editor of the same property, or another button of the same method</summary>
    UIElement CreateCopy(UIElement control)
    {
        if (control is MethodCallerButton button)
        {
            return CreateMethodCallerControl(button.OperationDescription, button.Target);
        }

        var value = (control as IValueEditor)?.Value;
        return value == null ? null : CreatePropertyEditorControl(value, Selection, ActionManager);
    }

    static string GetMemberName(UIElement control)
    {
        if (control is MethodCallerButton button)
        {
            return button.OperationDescription?.Name;
        }

        return (control as IValueEditor)?.Value?.Name;
    }

    /// <summary>
    /// Where a row or a button goes on the page (<see cref="PropertyGridSectionAttribute"/>):
    /// unmarked, a row is what defines the figure, a button a verb, a destructive button last
    /// </summary>
    static (PropertyGridSection Section, int Order) GetPlace(IMetadataDescription metadata, bool isButton)
    {
        var attribute = metadata?.GetAttribute<PropertyGridSectionAttribute>();
        if (attribute != null)
        {
            return (attribute.Section, attribute.Order);
        }

        if (isButton && metadata?.GetAttribute<PropertyGridDestructiveAttribute>() != null)
        {
            return (PropertyGridSection.Delete, 0);
        }

        return (isButton ? PropertyGridSection.Verbs : PropertyGridSection.Definition, 0);
    }

    /// <summary>Whether the row goes over the page, into its header (<see cref="PropertyGridSection.Header"/>)</summary>
    static bool IsHeaderRow(IValueProvider value)
    {
        return GetPlace(value, isButton: false).Section == PropertyGridSection.Header;
    }

    /// <summary>
    /// Rows and buttons by their sections and places (<see cref="PropertyGridSectionAttribute"/>),
    /// those of one place in the order they came. Members that carry the same
    /// <see cref="PropertyGridGroupAttribute"/> are boxed together (editors, then their
    /// buttons in a row), the box where its first member goes; the Delete section goes last,
    /// below a divider.
    /// </summary>
    static IEnumerable<UIElement> Arrange(IEnumerable<UIElement> controls)
    {
        var entries = new List<(PropertyGridSection Section, int Order, int Index, UIElement Control)>();
        var groups = new List<(string Name, StackPanel Editors, WrapPanel Buttons, int Entry)>();
        int index = 0;

        foreach (var control in controls)
        {
            var button = control as MethodCallerButton;
            IMetadataDescription metadata = button != null
                ? button.OperationDescription
                : (control as IValueEditor)?.Value;
            var (section, order) = GetPlace(metadata, button != null);
            index++;

            string groupName = section == PropertyGridSection.Delete
                ? null
                : metadata?.GetAttribute<PropertyGridGroupAttribute>()?.Name;
            if (groupName == null)
            {
                entries.Add((section, order, index, control));
                continue;
            }

            var group = groups.FirstOrDefault(g => g.Name == groupName);
            if (group.Name == null)
            {
                group = (groupName, new StackPanel(), new WrapPanel() { Margin = new Thickness(0, 4, 0, 0) }, entries.Count);
                groups.Add(group);

                var content = new StackPanel();
                content.Children.Add(group.Editors);
                content.Children.Add(group.Buttons);
                var box = new Border()
                {
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(6),
                    Padding = new Thickness(8, 6, 8, 6),
                    Margin = new Thickness(-8, 8, -8, 4),
                    Child = content
                };
                box.BindTheme(Border.BorderBrushProperty, nameof(AppTheme.Separator));
                box.BindTheme(Border.BackgroundProperty, nameof(AppTheme.GroupBackground));
                entries.Add((section, order, index, box));
            }
            else if (button == null)
            {
                // the box goes where its first row goes: the rows came before the buttons
                var first = entries[group.Entry];
                if ((section, order, index).CompareTo((first.Section, first.Order, first.Index)) < 0)
                {
                    entries[group.Entry] = (section, order, index, first.Control);
                }
            }

            if (button != null)
            {
                button.Margin = new Thickness(0, 2, 6, 2);
                group.Buttons.Children.Add(button);
            }
            else
            {
                group.Editors.Children.Add(control);
            }
        }

        var sorted = entries
            .OrderBy(entry => entry.Section)
            .ThenBy(entry => entry.Order)
            .ThenBy(entry => entry.Index)
            .ToList();
        var result = sorted
            .Where(entry => entry.Section != PropertyGridSection.Delete)
            .Select(entry => entry.Control)
            .ToList();
        var destructive = sorted
            .Where(entry => entry.Section == PropertyGridSection.Delete)
            .Select(entry => entry.Control)
            .ToList();

        if (destructive.Count > 0)
        {
            var divider = new Border()
            {
                Height = 1,
                Margin = new Thickness(-8, 12, -8, 8)
            };
            divider.BindTheme(Border.BackgroundProperty, nameof(AppTheme.Separator));

            // a live button (PropertyGridLiveCondition) may hide: no divider over nothing
            void UpdateDivider() => divider.IsVisible = destructive.Any(button => button.IsVisible);
            foreach (var button in destructive)
            {
                button.PropertyChanged += (s, e) =>
                {
                    if (e.Property == Visual.IsVisibleProperty)
                    {
                        UpdateDivider();
                    }
                };
            }

            UpdateDivider();
            result.Add(divider);
            result.AddRange(destructive);
        }

        return result;
    }

    protected virtual void AddHeader()
    {
        CreateHeader();
        if (HeaderHost != null)
        {
            HeaderHost.Child = Header;
        }
        else
        {
            this.Children.Add(Header);
        }
    }

    /// <summary>
    /// Where the title goes instead of above the rows: a place of the host's that stays put
    /// while the rows scroll (the side panel, with its close cross). Null: the first row.
    /// </summary>
    public Decorator HeaderHost { get; set; }

    public void UpdateHeader()
    {
        // a refresh can come after the grid was emptied (it is posted to the dispatcher)
        if (Selection == null || headerTitle == null)
        {
            return;
        }

        ShowHeader();
    }

    StackPanel Header { get; set; }

    TextBlock headerTitle;

    // what the figure is ("Segment"), before its name, when the name is a box of its own
    TextBlock headerKind;

    // the name's box, in the title's place
    HeaderNameEditor headerName;

    // Visible and Locked, at the right of the title
    StackPanel headerToggles;

    // the rows that go into the header (PropertyGridSection.Header), out of the page
    IValueProvider[] headerRows = [];

    // how the figure is built ("of CD"), under its title, smaller and faded
    TextBlock headerConstruction;

    void CreateHeader()
    {
        headerName = null;
        // in the title's size, as "Segment" read in "Segment AB", but plain and faded: the
        // name after it is the part that is edited
        headerKind = new TextBlock()
        {
            FontSize = HeaderNameEditor.TitleFontSize,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 12, 0),
            IsVisible = false,
            IsHitTestVisible = false
        };
        headerKind.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.TextMuted));
        DockPanel.SetDock(headerKind, Dock.Left);
        headerTitle = new TextBlock()
        {
            FontSize = HeaderNameEditor.TitleFontSize,
            FontWeight = FontWeight.SemiBold,
            TextWrapping = TextWrapping.Wrap,
            VerticalAlignment = VerticalAlignment.Center,
            IsHitTestVisible = false
        };
        headerTitle.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Text));
        headerToggles = new StackPanel()
        {
            Orientation = Orientation.Horizontal,
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(8, 0, 0, 0)
        };
        DockPanel.SetDock(headerToggles, Dock.Right);
        var titleRow = new DockPanel() { Children = { headerToggles, headerKind, headerTitle } };
        headerConstruction = new TextBlock()
        {
            FontSize = 12,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 2, 0, 0),
            IsHitTestVisible = false
        };
        headerConstruction.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.TextMuted));
        // a long construction (the expressions of a point by coordinates) wraps where the
        // text boxes below do, rather than stretching the panel
        Header = new StackPanel()
        {
            Margin = new Thickness(0, 0, 0, 10),
            MaxWidth = PropertyGridTheme.MaxTextWidth,
            HorizontalAlignment = HorizontalAlignment.Left,
            Children = { titleRow, headerConstruction }
        };
        ShowHeader();
    }

    /// <summary>
    /// The rows of the header (<see cref="PropertyGridSection.Header"/>), on the title's
    /// line: what the figure is, then the name as a box as wide as the name, then each
    /// on/off row (Visible, Locked) as a small toggle.
    /// </summary>
    void ShowHeaderRows()
    {
        headerToggles.Children.Clear();
        foreach (var row in headerRows.OrderBy(row => GetPlace(row, isButton: false).Order))
        {
            if (row.Type == typeof(bool))
            {
                headerToggles.Children.Add(new HeaderToggle(row, ActionManager));
            }
            else if (row.Type == typeof(string) && headerName == null)
            {
                headerName = new HeaderNameEditor() { ActionManager = ActionManager };
                headerName.Value = row;
                var titleRow = (DockPanel)headerTitle.Parent;
                titleRow.Children.Remove(headerTitle);
                titleRow.Children.Add(headerName);
                headerName.VerticalAlignment = VerticalAlignment.Top;
                headerKind.VerticalAlignment = VerticalAlignment.Top;
                headerToggles.VerticalAlignment = VerticalAlignment.Top;
                headerName.TextBox.SizeChanged += (s, e) => AlignHeaderLine();
                headerKind.SizeChanged += (s, e) => AlignHeaderLine();
                headerToggles.SizeChanged += (s, e) => AlignHeaderLine();
            }
        }

        // the toggles right after the name, or after the title of several figures, all on
        // the left: the title no longer takes the rest of the line
        if (headerToggles.Children.Count > 0)
        {
            var titleRow = (DockPanel)headerToggles.Parent;
            titleRow.Children.Remove(headerToggles);
            DockPanel.SetDock(headerToggles, Dock.Left);
            titleRow.Children.Add(headerToggles);
            titleRow.LastChildFill = false;
        }

        ShowHeader();
    }

    /// <summary>
    /// The kind, the name's box and the toggles share a middle line: the box's, not the
    /// middle of the name's editor, which grows downwards when an error shows under the box
    /// (centered, the box went up over the kind)
    /// </summary>
    void AlignHeaderLine()
    {
        if (headerName == null)
        {
            return;
        }

        double kind = headerKind.Bounds.Height;
        double box = headerName.TextBox.Bounds.Height;
        double toggles = headerToggles.Bounds.Height;
        double line = System.Math.Max(box, System.Math.Max(kind, toggles));
        var margin = headerKind.Margin;
        headerKind.Margin = new Thickness(margin.Left, (line - kind) / 2, margin.Right, 0);
        headerName.Margin = new Thickness(0, (line - box) / 2, 0, 0);
        margin = headerToggles.Margin;
        headerToggles.Margin = new Thickness(margin.Left, (line - toggles) / 2, margin.Right, 0);
    }

    void ShowHeader()
    {
        string title = Title;
        string construction = null;
        if (string.IsNullOrEmpty(title))
        {
            title = GetTitleString(Selection);
            construction = GetConstruction(Selection);
        }

        // the name has a box of its own: what the figure is goes before it
        string kind = headerName != null ? (Selection as FigureBase)?.KindCaption : null;
        headerKind.Text = kind;
        headerKind.IsVisible = !string.IsNullOrEmpty(kind);

        headerTitle.Text = title;
        headerConstruction.Text = construction;
        headerConstruction.IsVisible = !string.IsNullOrEmpty(construction);
    }

    /// <summary>A figure's construction, unless its page is named otherwise ([PropertyGridName])</summary>
    static string GetConstruction(object editableObject)
    {
        if (editableObject is not IFigure figure || editableObject.GetType().GetAttribute<PropertyGridNameAttribute>() != null)
        {
            return null;
        }

        return figure is FigureBase onPage ? onPage.ConstructionOnPage : figure.Construction;
    }

    public string Title { get; set; }

    protected static string GetTitleString(object editableObject)
    {
        var result = editableObject is IFigure figure ? figure.Title : editableObject.ToString();
        var type = editableObject.GetType();
        var attribute = type.GetAttribute<PropertyGridNameAttribute>();
        if (attribute != null)
        {
            result = attribute.Name;
        }
        return result;
    }

    public IEnumerable<IValueProvider> CurrentProperties { get; set; }

    public IEnumerable<UIElement> CurrentEditors { get; set; }

    public IEnumerable<UIElement> CreateObjectControls<T>(T editableObject)
    {
        CurrentProperties = GetEditableProperties(editableObject).ToArray();
        headerRows = CurrentProperties.Where(IsHeaderRow).ToArray();
        var currentEditors = CurrentProperties
            .Except(headerRows)
            .Select(p => CreatePropertyEditorControl(p, editableObject, ActionManager))
            .Where(c => c != null).ToArray();
        var currentMethods = GetCallableMethods(editableObject);

        // the values of the transformation a figure is the image of (the angle of the
        // rotation that made this triangle), after the rows that define it (a control
        // without a section of its own is Definition, and comes after the rows that came
        // first): the rows and buttons of one of its points, which share their sources with
        // the others, so an edit through them is an undo step on the point that moves the
        // whole image
        var image = editableObject as IFigure;
        var transformation = FigureTransformation.Of(image);
        if (transformation?.Values != null)
        {
            var values = transformation.Values;
            var transformationRows = TiedValues.Rows(values)
                .Select(p => CreatePropertyEditorControl(p, editableObject, ActionManager))
                .Where(c => c != null);
            currentEditors = currentEditors.Concat(transformationRows).ToArray();
            currentMethods = currentMethods.Concat(TiedValues.Buttons(values, () => image.Drawing.RaiseDisplayProperties(image)));

            // the image says the point's values in its header ("by 45°"), and the point,
            // not the image, says when they change
            hoistedValues = values as INotifyPropertyChanged;
            if (hoistedValues != null)
            {
                hoistedValues.PropertyChanged += HoistedValuesPropertyChanged;
            }
        }

        // what can be measured of a figure
        var quantities = editableObject is IFigure measured ? FigureQuantities.Of(measured) : new FigureQuantity[0];
        if (quantities.Count > 0)
        {
            var quantitiesEditor = new QuantitiesEditor(quantities, currentEditors.OfType<LabeledValueEditor>());
            currentEditors = currentEditors.Append(quantitiesEditor).ToArray();
        }

        CurrentEditors = currentEditors;
        var currentMethodButtons = currentMethods
            .Select(m => CreateMethodCallerControl(m, editableObject)).ToArray();
        liveButtons = currentMethodButtons
            .OfType<MethodCallerButton>()
            .Where(b => b.OperationDescription.GetAttribute<PropertyGridLiveConditionAttribute>() != null)
            .ToArray();
        liveRows = currentEditors
            .Where(editor => (editor as IValueEditor)?.Value?.GetAttribute<PropertyGridConditionAttribute>() != null)
            .ToArray();
        UpdateLiveButtons();

        return currentEditors.Concat(currentMethodButtons).ToArray();
    }

    /// <summary>The point whose tied-value rows the page borrowed (the image of a transformation); null when none</summary>
    INotifyPropertyChanged hoistedValues;

    void HoistedValuesPropertyChanged(object sender, PropertyChangedEventArgs e)
    {
        if (!string.IsNullOrEmpty(e.PropertyName))
        {
            RefreshNumbers();
        }
    }

    /// <summary>The buttons that show and hide as their object changes (<see cref="PropertyGridLiveConditionAttribute"/>)</summary>
    MethodCallerButton[] liveButtons = [];

    /// <summary>The rows that show and hide as their object changes (<see cref="PropertyGridConditionAttribute"/> on a property)</summary>
    UIElement[] liveRows = [];

    void UpdateLiveButtons()
    {
        foreach (var button in liveButtons)
        {
            button.IsVisible = button.Target is not IConditionalProperties conditions
                || conditions.CanEdit(button.OperationDescription.Name);
        }

        foreach (var row in liveRows)
        {
            var value = ((IValueEditor)row).Value;
            var condition = value.GetAttribute<PropertyGridConditionAttribute>();

            // a row of several figures shows while any of them would show it
            var owners = value is CompositeValueProvider composite
                ? composite.InnerList.Select(inner => inner.Parent)
                : new[] { value.Parent };
            row.IsVisible = owners.Any(condition.Holds);
        }
    }

    protected virtual IEnumerable<IValueProvider> GetEditableProperties<T>(T editableObject)
    {
        IValueDiscoveryStrategy discoveryStrategy = 
            LiveGeometry.ValueDiscoveryStrategy.Get(editableObject.GetType())
            ?? this.ValueDiscoveryStrategy;
        var result = discoveryStrategy.GetValues(editableObject);
        return result;
    }

    static UIElement CreateMethodCallerControl(IOperationDescription m, object obj)
    {
        return new MethodCallerButton() { OperationDescription = m, Target = obj };
    }

    /// <summary>The buttons of the object's page (also what its context menu offers of them, <see cref="ContextMenus"/>)</summary>
    public static IEnumerable<IOperationDescription> GetCallableMethods(object editableObject)
    {
        if (editableObject is ICustomMethodProvider)
        {
            return (editableObject as ICustomMethodProvider).GetMethods();
        }

        IEnumerable<MethodInfo> allMethods = editableObject.GetType().GetMethods();
        allMethods = allMethods
            .Where(m => m.ReturnType == typeof(void)
                && !m.IsSpecialName
                && m.IsPublic
                && m.HasAttribute<PropertyGridVisibleAttribute>()
                && PropertyGridConditionAttribute.Holds(m, editableObject));

        // a figure can veto a button by name the way it vetoes editing a row
        // (a segment shows Fix length or Free length, whichever applies), and caption it
        // ("Snap to Segment1")
        if (editableObject is IConditionalProperties conditions)
        {
            return allMethods
                .Where(m => conditions.CanEdit(m.Name) || m.HasAttribute<PropertyGridLiveConditionAttribute>())
                .Select(m => (IOperationDescription)new CaptionedMethod(MethodDescription.Create(m), conditions))
                .ToArray();
        }

        var result = allMethods.Select(m => (IOperationDescription)MethodDescription.Create(m));

        return result;
    }

    static IEnumerable<IValueEditorFactory> Factories = Reflector.DiscoverTypesAndInstantiate<IValueEditorFactory>();

    static UIElement CreatePropertyEditorControl(IValueProvider p, object obj, ActionManager actionManager)
    {
        // a style's color under the dark theme is its Dark override: that is what is edited
        p = ThemedValue.ForCurrentTheme(p);
        var factory = SelectProperFactory(p);
        if (factory != null)
        {
            try
            {
                var valueEditor = factory.CreateEditor(p);
                valueEditor.ActionManager = actionManager;
                UIElement result = valueEditor as UIElement;
                if (result != null)
                {
                    HookupEvents(p, result, obj);
                    return result;
                }
            }
            catch (Exception)
            {
            }
        }
        return null;
    }

    private static IValueEditorFactory SelectProperFactory(IValueProvider p)
    {
        var candidates = Factories.Where(f => f.SupportsValue(p)).OrderBy(f => f.LoadOrder);
        var candidate = candidates.FirstOrDefault();
        var attribute = p.GetAttribute<PropertyGridPreferredEditorAttribute>();
        if (attribute != null && !attribute.EditorTypeName.IsEmpty())
        {
            var substringCandidates = candidates.Where(f => f.GetType().Name.Contains(attribute.EditorTypeName));
            if (substringCandidates.Count() > 0)
            {
                candidate = substringCandidates.First();
            }
        }
        return candidate;
    }

    static void HookupEvents(IValueProvider p, UIElement result, object obj)
    {
        var eventAttributes = p
            .GetAttributes<PropertyGridEventAttribute>();
        if (eventAttributes != null)
        {
            foreach (var eventAttribute in eventAttributes)
            {
                HookupEvent(result, eventAttribute, obj);
            }
        }
    }

    static void HookupEvent(UIElement control, PropertyGridEventAttribute eventAttribute, object model)
    {
        var foundEvent = Reflector.FindEventByName(control.GetType(), eventAttribute.EventName);
        if (foundEvent != null)
        {
            try
            {
                Delegate d = Delegate.CreateDelegate(foundEvent.EventHandlerType, model, eventAttribute.HandlerName);
                foundEvent.AddEventHandler(control, d);
            }
            catch (Exception)
            {
            }
        }
    }
}

public class MessageBoxDialog : IPropertyGridHost
{
    [PropertyGridVisible]
    public virtual string Message => MessageText;

    public string MessageText { get; set; }

    [PropertyGridVisible]
    [PropertyGridIcon(PropertyGridIcon.Check)]
    public void OK()
    {
        PropertyGrid.Show(null, null);
        OKClicked();
    }

    protected virtual void OKClicked()
    {
    }

    public PropertyGrid PropertyGrid { get; set; }
}
