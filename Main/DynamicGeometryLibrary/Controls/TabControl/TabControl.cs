﻿// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Public License (Ms-PL).
// Please see http://go.microsoft.com/fwlink/?LinkID=131993 for details.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Windows.Automation.Peers;
using System.Windows.Controls.Primitives;
using System.Windows.Input;

namespace System.Windows.Controls
{
    /// <summary>
    /// Represents a control that contains multiple items that share the same
    /// space on the screen.
    /// </summary>
    /// <remarks>
    /// TabControl allows a developer to arrange visual content in a compacted
    /// and organized form.  The real-world analog of the control might be a
    /// tabbed notebook, in which visual content is displayed in discreet pages
    /// which are accessed by selecting the appropriate tab.  Each tab/page is
    /// encapsulated by a TabItem, the generated item of TabControl.  A TabItem
    /// has a Header property which corresponds to the content in the tab button
    /// and a Content property which corresponds to the content in the tab page.
    /// This control is useful for minimizing screen space usage while allowing
    /// an application to expose a large amount of data.  The user navigates
    /// through TabItems by clicking on a tab button using the mouse or by using
    /// the keyboard.
    /// </remarks>
    /// <QualityBand>Mature</QualityBand>
    [TemplatePart(Name = TabControl.ElementTemplateTopName, Type = typeof(FrameworkElement))]
    [TemplatePart(Name = TabControl.ElementTemplateBottomName, Type = typeof(FrameworkElement))]
    [TemplatePart(Name = TabControl.ElementTemplateLeftName, Type = typeof(FrameworkElement))]
    [TemplatePart(Name = TabControl.ElementTemplateRightName, Type = typeof(FrameworkElement))]
    [TemplatePart(Name = TabControl.ElementTabPanelTopName, Type = typeof(TabPanel))]
    [TemplatePart(Name = TabControl.ElementTabPanelBottomName, Type = typeof(TabPanel))]
    [TemplatePart(Name = TabControl.ElementTabPanelLeftName, Type = typeof(TabPanel))]
    [TemplatePart(Name = TabControl.ElementTabPanelRightName, Type = typeof(TabPanel))]
    [TemplatePart(Name = TabControl.ElementContentTopName, Type = typeof(ContentPresenter))]
    [TemplatePart(Name = TabControl.ElementContentBottomName, Type = typeof(ContentPresenter))]
    [TemplatePart(Name = TabControl.ElementContentLeftName, Type = typeof(ContentPresenter))]
    [TemplatePart(Name = TabControl.ElementContentRightName, Type = typeof(ContentPresenter))]
    [TemplateVisualState(Name = VisualStates.StateNormal, GroupName = VisualStates.GroupCommon)]
    [TemplateVisualState(Name = VisualStates.StateDisabled, GroupName = VisualStates.GroupCommon)]
    public partial class TabControl : ItemsControl
    {
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:System.Windows.Controls.TabControl" /> class.
        /// </summary>
        public TabControl()
        {
            SelectedIndex = -1;
            KeyDown += delegate(object sender, KeyEventArgs e) { OnKeyDown(e); };
            SelectionChanged += delegate(object sender, SelectionChangedEventArgs e) { OnSelectionChanged(e); };
            IsEnabledChanged += new DependencyPropertyChangedEventHandler(OnIsEnabledChanged);
            DefaultStyleKey = typeof(TabControl);
        }

        /// <summary>
        /// Builds the visual tree for the
        /// <see cref="T:System.Windows.Controls.TabControl" /> when a new
        /// template is applied.
        /// </summary>
        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();

            // if previous items exist, we want to clear
            if (ElementTabPanelTop != null)
            {
                ElementTabPanelTop.Children.Clear();
            }
            if (ElementTabPanelBottom != null)
            {
                ElementTabPanelBottom.Children.Clear();
            }
            if (ElementTabPanelLeft != null)
            {
                ElementTabPanelLeft.Children.Clear();
            }
            if (ElementTabPanelRight != null)
            {
                ElementTabPanelRight.Children.Clear();
            }

            // also clear the content if it is set to a ContentPresenter
            ContentPresenter contentHost = GetContentHost(TabStripPlacement);
            if (contentHost != null)
            {
                contentHost.Content = null;
            }

            // Get the parts
            ElementTemplateTop = GetTemplateChild(ElementTemplateTopName) as FrameworkElement;
            ElementTemplateBottom = GetTemplateChild(ElementTemplateBottomName) as FrameworkElement;
            ElementTemplateLeft = GetTemplateChild(ElementTemplateLeftName) as FrameworkElement;
            ElementTemplateRight = GetTemplateChild(ElementTemplateRightName) as FrameworkElement;

            ElementTabPanelTop = GetTemplateChild(ElementTabPanelTopName) as TabPanel;
            ElementTabPanelBottom = GetTemplateChild(ElementTabPanelBottomName) as TabPanel;
            ElementTabPanelLeft = GetTemplateChild(ElementTabPanelLeftName) as TabPanel;
            ElementTabPanelRight = GetTemplateChild(ElementTabPanelRightName) as TabPanel;

            ElementContentTop = GetTemplateChild(ElementContentTopName) as ContentPresenter;
            ElementContentBottom = GetTemplateChild(ElementContentBottomName) as ContentPresenter;
            ElementContentLeft = GetTemplateChild(ElementContentLeftName) as ContentPresenter;
            ElementContentRight = GetTemplateChild(ElementContentRightName) as ContentPresenter;

            foreach (object item in Items)
            {
                TabItem tabItem = item as TabItem;
                if (tabItem == null)
                {
                    throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, "TabControl_InvalidChild", item.GetType().ToString()));
                }
                AddToTabPanel(tabItem);
            }
            if (SelectedIndex >= 0) 
            {
                UpdateSelectedContent(SelectedContent);
            }

            UpdateTabPanelLayout(TabStripPlacement, TabStripPlacement);
            ChangeVisualState(false);
        }

        /// <summary>
        /// Returns 
        /// <see cref="T:System.Windows.Automation.Peers.TabControlAutomationPeer" />
        /// for use by the Silverlight automation infrastructure.
        /// </summary>
        /// <returns>
        /// A
        /// <see cref="T:System.Windows.Automation.Peers.TabControlAutomationPeer" />
        /// for the <see cref="T:System.Windows.Controls.TabControl" /> object.
        /// </returns>
        protected override AutomationPeer OnCreateAutomationPeer()
        {
            return new TabControlAutomationPeer(this);
        }

        #region SelectedItem
        /// <summary>
        /// Gets or sets the currently selected
        /// <see cref="T:System.Windows.Controls.TabItem" />.
        /// </summary>
        /// <value>
        /// The currently selected
        /// <see cref="T:System.Windows.Controls.TabItem" />, or null if a
        /// <see cref="T:System.Windows.Controls.TabItem" /> is not selected.
        /// </value>
        public object SelectedItem
        {
            get { return (object)GetValue(SelectedItemProperty); }
            set 
            { 
                if (value == null || Items.Contains(value)) 
                {
                    SetValue(SelectedItemProperty, value);
                } 
            }
        }

        /// <summary>
        /// Identifies the
        /// <see cref="P:System.Windows.Controls.TabControl.SelectedItem" />
        /// dependency property.
        /// </summary>
        /// <value>
        /// The identifier for the
        /// <see cref="P:System.Windows.Controls.TabControl.SelectedItem" />
        /// dependency property.
        /// </value>
        public static readonly DependencyProperty SelectedItemProperty =
            DependencyProperty.Register(
                "SelectedItem",
                typeof(object),
                typeof(TabControl),
                new PropertyMetadata(OnSelectedItemChanged));

        /// <summary>
        /// SelectedItem property changed handler.
        /// </summary>
        /// <param name="d">TabControl that changed its SelectedItem.</param>
        /// <param name="e">The DependencyPropertyChangedEventArgs.</param>
        private static void OnSelectedItemChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            TabControl tc = (TabControl)d;

            TabItem oldItem = e.OldValue as TabItem;
            TabItem tabItem = e.NewValue as TabItem;

            // if you select an item not in the TabControl keep the old
            // selection
            int index = (tabItem == null ? -1 : tc.Items.IndexOf(tabItem));
            if (tabItem != null && index == -1)
            {
                tc.SelectedItem = oldItem;
                return;
            }
            else
            {
                tc.SelectedIndex = index;
                tc.SelectItem(oldItem, tabItem);
            }
        }
        #endregion SelectedItem

        #region SelectedIndex
        /// <summary>
        /// Gets or sets the index of the currently selected
        /// <see cref="T:System.Windows.Controls.TabItem" />.
        /// </summary>
        /// <value>
        /// The index of the currently selected
        /// <see cref="T:System.Windows.Controls.TabItem" />, or -1 if a
        /// <see cref="T:System.Windows.Controls.TabItem" /> is not selected.
        /// </value>
        public int SelectedIndex
        {
            get { return (int)GetValue(SelectedIndexProperty); }
            set { SetValue(SelectedIndexProperty, value); }
        }

        /// <summary>
        /// Identifies the
        /// <see cref="P:System.Windows.Controls.TabControl.SelectedIndex" />
        /// dependency property.
        /// </summary>
        /// <value>
        /// The identifier for the
        /// <see cref="P:System.Windows.Controls.TabControl.SelectedIndex" />
        /// dependency property.
        /// </value>
        public static readonly DependencyProperty SelectedIndexProperty =
            DependencyProperty.Register(
                "SelectedIndex",
                typeof(int),
                typeof(TabControl),
                new PropertyMetadata(OnSelectedIndexChanged));

        /// <summary>
        /// SelectedIndex property changed handler.
        /// </summary>
        /// <param name="d">TabControl that changed its SelectedIndex.</param>
        /// <param name="e">The DependencyPropertyChangedEventArgs.</param>
        private static void OnSelectedIndexChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            TabControl tc = (TabControl)d;

            int newIndex = (int)e.NewValue;
            int oldIndex = (int)e.OldValue;

            if (newIndex < -1)
            {
                throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, "TabControl_InvalidIndex", newIndex.ToString(CultureInfo.CurrentCulture)));
            }

            // Coercion Workaround
            // -------------------
            if (tc._updateIndex)
            {
                tc._desiredIndex = newIndex;
            }
            else if (!tc._updateIndex)
            {
                tc._updateIndex = true;
            }
            if (newIndex >= tc.Items.Count)
            {
                tc._updateIndex = false;
                tc.SelectedIndex = oldIndex;
                return;
            }
            // -------------------

            TabItem item = tc.GetItemAtIndex(newIndex);
            if (tc.SelectedItem != item)
            {
                tc.SelectedItem = item;
            }
        }

        /// <summary>
        /// Given the TabItem in the list of Items, we will set that item as the
        /// currently selected item, and un-select the rest of the items.
        /// </summary>
        /// <param name="oldItem">Previous TabItem that was unselected.</param>
        /// <param name="newItem">New TabItem to set as the SelectedItem.</param>
        private void SelectItem(object oldItem, object newItem)
        {
            if (newItem == null)
            {
                ContentPresenter contentHost = GetContentHost(TabStripPlacement);
                if (contentHost != null) 
                {
                    contentHost.Content = null;
                }
                SetValue(SelectedContentProperty, null);
            }

            foreach (object item in Items)
            {
                TabItem tabItem = item as TabItem;
                if (tabItem == null)
                {
                    throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, "TabControl_InvalidChild", item.GetType().ToString()));
                }
                if (tabItem != newItem && tabItem.IsSelected)
                {
                    tabItem.IsSelected = false;
                }
                else if (tabItem == newItem)
                {
                    tabItem.IsSelected = true;
                    SetValue(SelectedContentProperty, tabItem.Content);
                }
            }

            // Fire SelectionChanged Event
            SelectionChangedEventHandler handler = SelectionChanged;
            if (handler != null)
            {
                TabItem oldTabItem = oldItem as TabItem;
                TabItem newTabItem = newItem as TabItem;

                SelectionChangedEventArgs args = new SelectionChangedEventArgs(
                    (oldTabItem == null ? new List<TabItem> { } : new List<TabItem> { oldTabItem }),
                    (newTabItem == null ? new List<TabItem> { } : new List<TabItem> { newTabItem }));
                handler(this, args);
            }
        }
        #endregion SelectedIndex

        #region SelectedContent
        /// <summary>
        /// Gets the content of the currently selected
        /// <see cref="T:System.Windows.Controls.TabItem" />.
        /// </summary>
        /// <value>
        /// The content of the currently selected
        /// <see cref="T:System.Windows.Controls.TabItem" />. The default is
        /// null.
        /// </value>
        public object SelectedContent
        {
            get { return GetValue(SelectedContentProperty); }
            internal set { SetValue(SelectedContentProperty, value); }
        }

        /// <summary>
        /// Identifies the
        /// <see cref="P:System.Windows.Controls.TabControl.SelectedContent" />
        /// dependency property.
        /// </summary>
        /// <value>
        /// The identifier for the
        /// <see cref="P:System.Windows.Controls.TabControl.SelectedContent" />
        /// dependency property.
        /// </value>
        public static readonly DependencyProperty SelectedContentProperty =
            DependencyProperty.Register(
                "SelectedContent",
                typeof(object),
                typeof(TabControl),
                new PropertyMetadata(OnSelectedContentChanged));

        /// <summary>
        /// SelectedContent property changed handler.
        /// </summary>
        /// <param name="d">TabControl that changed its SelectedContent.</param>
        /// <param name="e">The DependencyPropertyChangedEventArgs.</param>
        private static void OnSelectedContentChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            TabControl tc = (TabControl)d;

            tc.UpdateSelectedContent(e.NewValue);
        }
        #endregion SelectedContent

        #region TabStripPlacement
        /// <summary>
        /// Gets or sets how <see cref="T:System.Windows.Controls.TabItem" />
        /// headers align relative to the
        /// <see cref="T:System.Windows.Controls.TabItem" /> content.
        /// </summary>
        /// <value>
        /// The alignment of <see cref="T:System.Windows.Controls.TabItem" />
        /// headers relative to the
        /// <see cref="T:System.Windows.Controls.TabItem" /> content. The
        /// default is <see cref="F:System.Windows.Controls.Dock.Top" />.
        /// </value>
        public Dock TabStripPlacement
        {
            get { return (Dock)GetValue(TabStripPlacementProperty); }
            set { SetValue(TabStripPlacementProperty, value); }
        }

        /// <summary>
        /// Identifies the
        /// <see cref="P:System.Windows.Controls.TabControl.TabStripPlacement" />
        /// dependency property.
        /// </summary>
        /// <value>
        /// The identifier for the
        /// <see cref="P:System.Windows.Controls.TabControl.TabStripPlacement" />
        /// dependency property.
        /// </value>
        public static readonly DependencyProperty TabStripPlacementProperty =
            DependencyProperty.Register(
                "TabStripPlacement",
                typeof(Dock),
                typeof(TabControl),
                new PropertyMetadata(Dock.Top, OnTabStripPlacementPropertyChanged));

        /// <summary>
        /// TabStripPlacement property changed handler.
        /// </summary>
        /// <param name="d">
        /// TabControl that changed its TabStripPlacement.
        /// </param>
        /// <param name="e">The DependencyPropertyChangedEventArgs.</param>
        private static void OnTabStripPlacementPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            TabControl tc = (TabControl)d;
            Debug.Assert(tc != null, "TabControl should not be null!");
            tc.UpdateTabPanelLayout((Dock)e.OldValue, (Dock)e.NewValue);

            foreach (object o in tc.Items)
            {
                TabItem tabItem = o as TabItem;
                if (tabItem != null)
                {
                    tabItem.UpdateVisualState();
                }
            }
        }
        #endregion TabStripPlacement

        /// <summary>
        /// Called when the IsEnabled property changes.
        /// </summary>
        /// <param name="sender">
        /// Control that triggers this property change.
        /// </param>
        /// <param name="e">Property changed args.</param>
        private void OnIsEnabledChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            UpdateVisualState();
        }

        /// <summary>
        /// This will hide the templates for the TabStripPlacements that are not
        /// being displayed, and show the template for the TabStripPlacement
        /// that is currently selected.
        /// </summary>
        /// <param name="oldValue">Inherited code: Requires comment.</param>
        /// <param name="newValue">Inherited code: Requires comment 1.</param>
        private void UpdateTabPanelLayout(Dock oldValue, Dock newValue)
        {
            FrameworkElement oldTemplate = GetTemplate(oldValue);
            FrameworkElement newTemplate = GetTemplate(newValue);
            TabPanel oldPanel = GetTabPanel(oldValue);
            TabPanel newPanel = GetTabPanel(newValue);
            ContentPresenter oldContentHost = GetContentHost(oldValue);
            ContentPresenter newContentHost = GetContentHost(newValue);

            if (oldValue != newValue) 
            {
                if (oldTemplate != null) 
                {
                    oldTemplate.Visibility = Visibility.Collapsed;
                }
                if (oldPanel != null) 
                {
                    oldPanel.Children.Clear();
                }
                if (newPanel != null)
                {
                    foreach (object item in Items)
                    {
                        TabItem tabItem = item as TabItem;
                        if (tabItem == null)
                        {
                            throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, "TabControl_InvalidChild", item.GetType().ToString()));
                        }
                        AddToTabPanel(tabItem);
                    }
                }
                if (oldContentHost != null) 
                {
                    oldContentHost.Content = null;
                }
                if (newContentHost != null) 
                {
                    newContentHost.Content = SelectedContent;
                }
            }
            if (newTemplate != null) 
            {
                newTemplate.Visibility = Visibility.Visible;
            }
        }

        /// <summary>
        /// Update the current visual state of the TabControl.
        /// </summary>
        internal void UpdateVisualState()
        {
            ChangeVisualState(true);
        }

        /// <summary>
        /// Change to the correct visual state for the TabControl.
        /// </summary>
        /// <param name="useTransitions">
        /// True to use transitions when updating the visual state, false to
        /// snap directly to the new visual state.
        /// </param>
        internal void ChangeVisualState(bool useTransitions)
        {
            if (!IsEnabled)
            {
                VisualStates.GoToState(this, useTransitions, VisualStates.StateDisabled, VisualStates.StateNormal);
            }
            else
            {
                VisualStates.GoToState(this, useTransitions, VisualStates.StateNormal);
            }
        }
        
        /// <summary>
        /// Updates the current selection when Items has changed.
        /// </summary>
        /// <param name="e">Data used by the event.</param>
        [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity", Justification = "Need to handle Add/Remove/Replace")]
        [SuppressMessage("Microsoft.Naming", "CA1704:IdentifiersShouldBeSpelledCorrectly", MessageId = "e", Justification = "Compat with WPF.")]
        protected override void OnItemsChanged(NotifyCollectionChangedEventArgs e)
        {
            base.OnItemsChanged(e);

            switch (e.Action)
            {
                case NotifyCollectionChangedAction.Add:
                    int newSelectedIndex = -1;
                    foreach (object o in e.NewItems)
                    {
                        TabItem tabItem = o as TabItem;
                        if (tabItem == null)
                        {
                            throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, "TabControl_InvalidChild", o.GetType().ToString()));
                        }
                        int index = Items.IndexOf(tabItem);
                        InsertIntoTabPanel(index, tabItem);

                        // If we are adding a selected item
                        if (tabItem.IsSelected)
                        {
                            newSelectedIndex = index;
                        }
                        else if (SelectedItem != GetItemAtIndex(SelectedIndex))
                        {
                            newSelectedIndex = Items.IndexOf(SelectedItem);
                        }
                        else if ((_desiredIndex < Items.Count) && (_desiredIndex >= 0))
                        {
                            // Coercion Workaround
                            newSelectedIndex = _desiredIndex;
                        }

                        tabItem.UpdateVisualState();
                    }

                    if (newSelectedIndex == -1)
                    {
                        // If we are adding many items through xaml, one could
                        // already be specified as selected. If so, we don't
                        // want to override the value.
                        foreach (object item in Items)
                        {
                            TabItem tabItem = item as TabItem;
                            if (tabItem == null)
                            {
                                throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, "TabControl_InvalidChild", item.GetType().ToString()));
                            }
                            
                            if (tabItem.IsSelected)
                            {
                                return;
                            }
                        }

                        // To follow WPF behavior, we only select the item if
                        // the user has not explicitly set the IsSelected field
                        // to false, or if there are 2 or more items in the
                        // TabControl.
                        if (Items.Count > 1 || ((Items[0] as TabItem).ReadLocalValue(TabItem.IsSelectedProperty) as bool?) != false)
                        {
                            newSelectedIndex = 0;
                        }
                    }

                    // When we add a new item into the selected position, 
                    // SelectedIndex does not change, so we need to update both
                    // the SelectedItem and the SelectedIndex.
                    SelectedItem = GetItemAtIndex(newSelectedIndex);
                    SelectedIndex = newSelectedIndex;
                    break;

                case NotifyCollectionChangedAction.Remove:
                    foreach (object o in e.OldItems)
                    {
                        TabItem tabItem = o as TabItem;
                        RemoveFromTabPanel(tabItem);
                        
                        // if there are no items, the selected index is set to
                        // -1
                        if (Items.Count == 0)
                        {
                            SelectedIndex = -1;
                        }
                        else if (Items.Count <= SelectedIndex)
                        {
                            SelectedIndex = Items.Count - 1;
                        }
                        else
                        {
                            SelectedItem = GetItemAtIndex(SelectedIndex);
                        }
                    }
                    break;

                case NotifyCollectionChangedAction.Reset:
                    ClearTabPanel();
                    SelectedIndex = -1;

                    // For Setting the ItemsSource
                    foreach (object item in Items)
                    {
                        TabItem tabItem = item as TabItem;
                        if (tabItem == null)
                        {
                            throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, "TabControl_InvalidChild", item.GetType().ToString()));
                        }
                        AddToTabPanel(tabItem);
                        if (tabItem.IsSelected) 
                        {
                            SelectedItem = tabItem;
                        }
                    }
                    if (SelectedIndex == -1 && Items.Count > 0) 
                    {
                        SelectedIndex = 0;
                    }
                    break;

                case NotifyCollectionChangedAction.Replace:
                    break;
            }
        }

        /// <summary>
        /// Occurs when the selected
        /// <see cref="T:System.Windows.Controls.TabItem" /> changes.
        /// </summary>
        public event SelectionChangedEventHandler SelectionChanged;

        /// <summary>
        /// Raises the
        /// <see cref="E:System.Windows.Controls.TabControl.SelectionChanged" />
        /// event.
        /// </summary>
        /// <param name="args">
        /// Provides data for the
        /// <see cref="E:System.Windows.Controls.TabControl.SelectionChanged" />
        /// event.
        /// </param>
        protected virtual void OnSelectionChanged(SelectionChangedEventArgs args)
        {
        }

        /// <summary>
        /// This is the method that responds to the KeyDown event.
        /// </summary>
        /// <param name="e">Data used by the event.</param>
        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.Handled)
            {
                return;
            }
            TabItem nextTabItem = null;

            int direction = 0;
            int startIndex = -1;
            switch (e.Key)
            {
                case Key.Home:
                    direction = 1;
                    startIndex = -1;
                    break;
                case Key.End:
                    direction = -1;
                    startIndex = Items.Count;
                    break;
                default:
                    return;
            }

            nextTabItem = FindNextTabItem(startIndex, direction);

            if (nextTabItem != null && nextTabItem != SelectedItem)
            {
                e.Handled = true;
                SelectedItem = nextTabItem;
                nextTabItem.Focus();
            }
        }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        /// <param name="startIndex">Inherited code: Requires comment 1.</param>
        /// <param name="direction">Inherited code: Requires comment 2.</param>
        /// <returns>Inherited code: Requires comment 3.</returns>
        internal TabItem FindNextTabItem(int startIndex, int direction)
        {
            Debug.Assert(direction != 0, "The direction should not be zero!");
            TabItem nextTabItem = null;

            int index = startIndex;
            for (int i = 0; i < Items.Count; i++)
            {
                index += direction;
                if (index >= Items.Count)
                {
                    index = 0;
                }
                else if (index < 0)
                {
                    index = Items.Count - 1;
                }

                TabItem tabItem = GetItemAtIndex(index);
                if (tabItem != null && tabItem.IsEnabled && tabItem.Visibility == Visibility.Visible)
                {
                    nextTabItem = tabItem;
                    break;
                }
            }
            return nextTabItem;
        }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        /// <param name="index">Inherited code: Requires comment 1.</param>
        /// <returns>Inherited code: Requires comment 2.</returns>
        private TabItem GetItemAtIndex(int index)
        {
            if (index < 0 || index >= Items.Count)
            {
                return null;
            }
            else
            {
                return Items[index] as TabItem;
            }
        }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        /// <param name="content">Inherited code: Requires comment 1.</param>
        private void UpdateSelectedContent(object content)
        {
            TabItem tabItem = SelectedItem as TabItem;
            if (tabItem != null)
            {
                ContentPresenter contentHost = GetContentHost(TabStripPlacement);
                if (contentHost != null)
                {
                    contentHost.HorizontalAlignment = tabItem.HorizontalContentAlignment;
                    contentHost.VerticalAlignment = tabItem.VerticalContentAlignment;
                    contentHost.ContentTemplate = tabItem.ContentTemplate;
                    contentHost.Content = content;
                }
            }
        }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        /// <param name="tabItem">Inherited code: Requires comment 1.</param>
        private void AddToTabPanel(TabItem tabItem)
        {
            TabPanel panel = GetTabPanel(TabStripPlacement);
            if (panel != null && !panel.Children.Contains(tabItem)) 
            {
                panel.Children.Add(tabItem);
            }
        }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        /// <param name="index">Inherited code: Requires comment 1.</param>
        /// <param name="tabItem">Inherited code: Requires comment 2.</param>
        private void InsertIntoTabPanel(int index, TabItem tabItem)
        {
            TabPanel panel = GetTabPanel(TabStripPlacement);
            if (panel != null && !panel.Children.Contains(tabItem)) 
            {
                panel.Children.Insert(index, tabItem);
            }
        }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        /// <param name="tabItem">Inherited code: Requires comment 1.</param>
        private void RemoveFromTabPanel(TabItem tabItem)
        {
            TabPanel panel = GetTabPanel(TabStripPlacement);
            if (panel != null && panel.Children.Contains(tabItem)) 
            {
                panel.Children.Remove(tabItem);
            }
        }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        private void ClearTabPanel()
        {
            TabPanel panel = GetTabPanel(TabStripPlacement);
            if (panel != null) 
            {
                panel.Children.Clear();
            }
        }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        /// <param name="tabPlacement">Inherited code: Requires comment 1.</param>
        /// <returns>Inherited code: Requires comment 2.</returns>
        internal TabPanel GetTabPanel(Dock tabPlacement)
        {
            switch (tabPlacement)
            {
                case Dock.Top:
                    return ElementTabPanelTop;
                case Dock.Bottom:
                    return ElementTabPanelBottom;
                case Dock.Left:
                    return ElementTabPanelLeft;
                case Dock.Right:
                    return ElementTabPanelRight;
            }
            return null;
        }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        /// <param name="tabPlacement">Inherited code: Requires comment 1.</param>
        /// <returns>Inherited code: Requires comment 2.</returns>
        internal FrameworkElement GetTemplate(Dock tabPlacement)
        {
            switch (tabPlacement)
            {
                case Dock.Top:
                    return ElementTemplateTop;
                case Dock.Bottom:
                    return ElementTemplateBottom;
                case Dock.Left:
                    return ElementTemplateLeft;
                case Dock.Right:
                    return ElementTemplateRight;
            }
            return null;
        }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        /// <param name="tabPlacement">Inherited code: Requires comment 1.</param>
        /// <returns>Inherited code: Requires comment 2.</returns>
        internal ContentPresenter GetContentHost(Dock tabPlacement)
        {
            switch (tabPlacement)
            {
                case Dock.Top:
                    return ElementContentTop;
                case Dock.Bottom:
                    return ElementContentBottom;
                case Dock.Left:
                    return ElementContentLeft;
                case Dock.Right:
                    return ElementContentRight;
            }
            return null;
        }

        /// <summary>
        /// Gets or sets the TabStripPlacement Top template.
        /// </summary>
        internal FrameworkElement ElementTemplateTop { get; set; }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        internal const string ElementTemplateTopName = "TemplateTop";

        /// <summary>
        /// Gets or sets the TabStripPlacement Bottom template.
        /// </summary>
        internal FrameworkElement ElementTemplateBottom { get; set; }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        internal const string ElementTemplateBottomName = "TemplateBottom";
        
        /// <summary>
        /// Gets or sets the TabStripPlacement Left template.
        /// </summary>
        internal FrameworkElement ElementTemplateLeft { get; set; }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        internal const string ElementTemplateLeftName = "TemplateLeft";
        
        /// <summary>
        /// Gets or sets the TabStripPlacement Right template.
        /// </summary>
        internal FrameworkElement ElementTemplateRight { get; set; }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        internal const string ElementTemplateRightName = "TemplateRight";

        /// <summary>
        /// Gets or sets the TabPanel of the TabStripPlacement Top template.
        /// </summary>
        internal TabPanel ElementTabPanelTop { get; set; }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        internal const string ElementTabPanelTopName = "TabPanelTop";

        /// <summary>
        /// Gets or sets the TabPanel of the TabStripPlacement Bottom template.
        /// </summary>
        internal TabPanel ElementTabPanelBottom { get; set; }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        internal const string ElementTabPanelBottomName = "TabPanelBottom";

        /// <summary>
        /// Gets or sets the TabPanel of the TabStripPlacement Left template.
        /// </summary>
        internal TabPanel ElementTabPanelLeft { get; set; }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        internal const string ElementTabPanelLeftName = "TabPanelLeft";

        /// <summary>
        /// Gets or sets the TabPanel of the TabStripPlacement Right template.
        /// </summary>
        internal TabPanel ElementTabPanelRight { get; set; }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        internal const string ElementTabPanelRightName = "TabPanelRight";

        /// <summary>
        /// Gets or sets the ContentHost of the TabStripPlacement Top template.
        /// </summary>
        internal ContentPresenter ElementContentTop { get; set; }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        internal const string ElementContentTopName = "ContentTop";

        /// <summary>
        /// Gets or sets the ContentHost of the TabStripPlacement Bottom
        /// template.
        /// </summary>
        internal ContentPresenter ElementContentBottom { get; set; }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        internal const string ElementContentBottomName = "ContentBottom";

        /// <summary>
        /// Gets or sets the ContentHost of the TabStripPlacement Left template.
        /// </summary>
        internal ContentPresenter ElementContentLeft { get; set; }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        internal const string ElementContentLeftName = "ContentLeft";

        /// <summary>
        /// Gets or sets the ContentHost of the TabStripPlacement Right
        /// template.
        /// </summary>
        internal ContentPresenter ElementContentRight { get; set; }

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        internal const string ElementContentRightName = "ContentRight";

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        private int _desiredIndex;

        /// <summary>
        /// Inherited code: Requires comment.
        /// </summary>
        private bool _updateIndex = true;
    }
}