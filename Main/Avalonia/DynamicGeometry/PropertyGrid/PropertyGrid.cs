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

namespace DynamicGeometry
{
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
            get
            {
                return mSelection;
            }
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
                    UpdateVisibility(true);
                }
                else
                {
                    UpdateVisibility(false);
                }
                UpdateContents();
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
            get
            {
                return mExpanded;
            }
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
        /// Plain editors and buttons stay in their order. Members that carry the same
        /// <see cref="PropertyGridGroupAttribute"/> are boxed together (editors, then their
        /// buttons in a row), and destructive buttons go last, below a divider.
        /// </summary>
        static IEnumerable<UIElement> Arrange(IEnumerable<UIElement> controls)
        {
            // in this order: plain fields, group boxes, plain buttons, destructive buttons
            var result = new List<UIElement>();
            var groupBoxes = new List<UIElement>();
            var plainButtons = new List<UIElement>();
            var groups = new List<(string Name, StackPanel Editors, WrapPanel Buttons)>();
            var destructive = new List<UIElement>();

            foreach (var control in controls)
            {
                var button = control as MethodCallerButton;
                IMetadataDescription metadata = button != null
                    ? button.OperationDescription
                    : (control as IValueEditor)?.Value;

                if (button != null && metadata?.GetAttribute<PropertyGridDestructiveAttribute>() != null)
                {
                    destructive.Add(control);
                    continue;
                }

                string groupName = metadata?.GetAttribute<PropertyGridGroupAttribute>()?.Name;
                if (groupName == null)
                {
                    (button != null ? plainButtons : result).Add(control);
                    continue;
                }

                var group = groups.FirstOrDefault(g => g.Name == groupName);
                if (group.Name == null)
                {
                    group = (groupName, new StackPanel(), new WrapPanel() { Margin = new Thickness(0, 4, 0, 0) });
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
                    groupBoxes.Add(box);
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

            result.AddRange(groupBoxes);
            result.AddRange(plainButtons);

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

        UIElement Header { get; set; }

        TextBlock headerTitle;

        // how the figure is built ("of CD"), under its title, smaller and faded
        TextBlock headerConstruction;

        void CreateHeader()
        {
            headerTitle = new TextBlock()
            {
                FontSize = 15,
                FontWeight = FontWeight.SemiBold,
                TextWrapping = TextWrapping.Wrap,
                IsHitTestVisible = false
            };
            headerTitle.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Text));
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
                Children = { headerTitle, headerConstruction }
            };
            ShowHeader();
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

            return figure.Construction;
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
            var currentEditors = CurrentProperties
                .Select(p => CreatePropertyEditorControl(p, editableObject, ActionManager))
                .Where(c => c != null).ToArray();
            CurrentEditors = currentEditors;
            var currentMethods = GetCallableMethods(editableObject);
            var currentMethodButtons = currentMethods
                .Select(m => CreateMethodCallerControl(m, editableObject)).ToArray();
            liveButtons = currentMethodButtons
                .OfType<MethodCallerButton>()
                .Where(b => b.OperationDescription.GetAttribute<PropertyGridLiveConditionAttribute>() != null)
                .ToArray();
            UpdateLiveButtons();

            return currentEditors.Concat(currentMethodButtons).ToArray();
        }

        /// <summary>The buttons that show and hide as their object changes (<see cref="PropertyGridLiveConditionAttribute"/>)</summary>
        MethodCallerButton[] liveButtons = new MethodCallerButton[0];

        void UpdateLiveButtons()
        {
            foreach (var button in liveButtons)
            {
                button.IsVisible = button.Target is not IConditionalProperties conditions
                    || conditions.CanEdit(button.OperationDescription.Name);
            }
        }

        protected virtual IEnumerable<IValueProvider> GetEditableProperties<T>(T editableObject)
        {
            IValueDiscoveryStrategy discoveryStrategy = 
                DynamicGeometry.ValueDiscoveryStrategy.Get(editableObject.GetType())
                ?? this.ValueDiscoveryStrategy;
            var result = discoveryStrategy.GetValues(editableObject);
            return result;
        }

        static UIElement CreateMethodCallerControl(IOperationDescription m, object obj)
        {
            return new MethodCallerButton() { OperationDescription = m, Target = obj };
        }

        private static IEnumerable<IOperationDescription> GetCallableMethods(object editableObject)
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
                    && m.HasAttribute<PropertyGridVisibleAttribute>());

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
        public virtual string Message
        {
            get
            {
                return MessageText;
            }
        }

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
}
