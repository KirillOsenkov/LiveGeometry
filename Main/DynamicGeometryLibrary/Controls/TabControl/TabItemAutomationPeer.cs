﻿// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Public License (Ms-PL).
// Please see http://go.microsoft.com/fwlink/?LinkID=131993 for details.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Windows.Automation.Provider;
using System.Windows.Controls;

[assembly: SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member", Target = "System.Windows.Automation.Peers.TabItemAutomationPeer.#System.Windows.Automation.Provider.ISelectionItemProvider.AddToSelection()", Justification = "WPF Compatibility")]
[assembly: SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member", Target = "System.Windows.Automation.Peers.TabItemAutomationPeer.#System.Windows.Automation.Provider.ISelectionItemProvider.IsSelected", Justification = "WPF Compatibility")]
[assembly: SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member", Target = "System.Windows.Automation.Peers.TabItemAutomationPeer.#System.Windows.Automation.Provider.ISelectionItemProvider.RemoveFromSelection()", Justification = "WPF Compatibility")]
[assembly: SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member", Target = "System.Windows.Automation.Peers.TabItemAutomationPeer.#System.Windows.Automation.Provider.ISelectionItemProvider.Select()", Justification = "WPF Compatibility")]
[assembly: SuppressMessage("Microsoft.Design", "CA1033:InterfaceMethodsShouldBeCallableByChildTypes", Scope = "member", Target = "System.Windows.Automation.Peers.TabItemAutomationPeer.#System.Windows.Automation.Provider.ISelectionItemProvider.SelectionContainer", Justification = "WPF Compatibility")]

namespace System.Windows.Automation.Peers
{
    /// <summary>
    /// Exposes <see cref="T:System.Windows.Controls.TabItem" /> types to  UI
    /// automation.
    /// </summary>
    /// <QualityBand>Mature</QualityBand>
    public partial class TabItemAutomationPeer : ItemAutomationPeer, ISelectionItemProvider
    {
        /// <summary>
        /// Initializes a new instance of the
        /// <see cref="T:System.Windows.Automation.Peers.TabItemAutomationPeer" />
        /// class.
        /// </summary>
        /// <param name="owner">
        /// The <see cref="T:System.Windows.Controls.TabItem" /> to associate
        /// with the
        /// <see cref="T:System.Windows.Automation.Peers.TabItemAutomationPeer" />.
        /// </param>
        public TabItemAutomationPeer(object owner)
            : base(owner as UIElement)
        {
        }

        /// <summary>
        /// Gets the control type for the element that is associated with the UI
        /// Automation peer.
        /// </summary>
        /// <returns>The control type.</returns>
        protected override AutomationControlType GetAutomationControlTypeCore()
        {
            return AutomationControlType.TabItem;
        }

        /// <summary>
        /// Returns the ChildrenCore.
        /// </summary>
        /// <returns>List of AutomationPeers for the Children.</returns>
        [SuppressMessage("Microsoft.Design", "CA1002:DoNotExposeGenericLists", Justification = "Requires a return list of AutomationPeers")]
        protected override List<AutomationPeer> GetChildrenCore()
        {
            List<AutomationPeer> childrenCore = base.GetChildrenCore();
            TabItem wrapper = Item as TabItem;
            if ((wrapper != null) && wrapper.IsSelected)
            {
                TabControl control = TabOwner;
                if (control == null)
                {
                    return childrenCore;
                }
                ContentPresenter selectedContentPresenter = control.GetContentHost(control.TabStripPlacement);
                if (selectedContentPresenter == null)
                {
                    return childrenCore;
                }
                List<AutomationPeer> children = new FrameworkElementAutomationPeer(selectedContentPresenter).GetChildren();
                if (children == null)
                {
                    return childrenCore;
                }
                if (childrenCore == null)
                {
                    return children;
                }
                childrenCore.AddRange(children);
            }
            return childrenCore;
        }

        /// <summary>
        /// Called by GetClassName that gets a human readable name that, in
        /// addition to AutomationControlType, differentiates the control
        /// represented by this AutomationPeer.
        /// </summary>
        /// <returns>The string that contains the name.</returns>
        protected override string GetClassNameCore()
        {
            return Owner.GetType().Name;
        }

        /// <summary>
        /// Gets the control pattern for the
        /// <see cref="T:System.Windows.Controls.TabItem" /> that is associated
        /// with this
        /// <see cref="T:System.Windows.Automation.Peers.TabItemAutomationPeer" />.
        /// </summary>
        /// <param name="patternInterface">
        /// One of the enumeration values.
        /// </param>
        /// <returns>
        /// The object that implements the pattern interface, or null if the
        /// specified pattern interface is not implemented by this peer.
        /// </returns>
        public override object GetPattern(PatternInterface patternInterface)
        {
            if (patternInterface == PatternInterface.SelectionItem)
            {
                return this;
            }
            return base.GetPattern(patternInterface);
        }

        /// <summary>
        /// Gets the text label of the TabItem that is associated with this
        /// TabItemAutomationPeer. Called by GetName. 
        /// </summary>
        /// <returns>
        /// The string that contains the label. If set, this method returns the
        /// value of the Name property; otherwise this method will return the
        /// value of the Header property. 
        /// </returns>
        protected override string GetNameCore()
        {
            TabItem wrapper = this.Owner as TabItem;
            Debug.Assert(wrapper != null, "Owner should be a non-null TabItem!");
            return (wrapper.Header as string) ?? String.Empty;
        }
        
        /// <summary>
        /// Adds the current element to the collection of selected items.
        /// </summary>
        /// <exception cref="T:System.Windows.Automation.ElementNotEnabledException">
        /// Owner element is not enabled.
        /// </exception>
        void ISelectionItemProvider.AddToSelection()
        {
            if (!IsEnabled())
            {
                throw new ElementNotEnabledException();
            }

            TabControl control = TabOwner;
            if (control == null)
            {
                // TODO: update when Jolt 23302 is fixed
                // throw new InvalidOperationException();
                return;
            }
            control.SelectedItem = Item;
        }

        /// <summary>
        /// Gets a value indicating whether an item is selected.
        /// </summary>
        /// <value>True if the element is selected; otherwise, false.</value>
        bool ISelectionItemProvider.IsSelected
        {
            get
            {
                TabControl control = TabOwner;
                return control != null && Item.Equals(control.SelectedItem);
            }
        }

        /// <summary>
        /// Removes the current element from the collection of selected items.
        /// </summary>
        void ISelectionItemProvider.RemoveFromSelection()
        {
            // Even though we can set the SelectedItem to null programmatically
            // through code, we cannot do this through normal interaction with
            // the TabControl, and so we will not allow Automation to
            // RemoveFromSelection either.

            // TODO: update when Jolt 23302 is fixed
            // throw new InvalidOperationException();
            return;
        }

        /// <summary>
        /// Deselects any selected items and then selects the current element.
        /// </summary>
        /// <exception cref="T:System.Windows.Automation.ElementNotEnabledException">
        /// Owner element is not enabled.
        /// </exception>
        void ISelectionItemProvider.Select()
        {
            if (!IsEnabled())
            {
                throw new ElementNotEnabledException();
            }

            TabControl control = TabOwner;
            if (control == null)
            {
                // TODO: update when Jolt 23302 is fixed
                // throw new InvalidOperationException();
                return;
            }
            control.SelectedItem = Item;
        }

        /// <summary>
        /// Gets the UI automation provider that implements
        /// <see cref="T:System.Windows.Automation.Provider.ISelectionProvider" />
        /// and acts as the container for the calling object.
        /// </summary>
        /// <value>
        /// The provider that supports
        /// <see cref="T:System.Windows.Automation.Provider.ISelectionProvider" />.
        /// </value>
        IRawElementProviderSimple ISelectionItemProvider.SelectionContainer
        {
            get { return ProviderFromPeer(ItemsControlAutomationPeer); }
        }

        /// <summary>
        /// Gets Inherited code: Requires comment.
        /// </summary>
        private TabControl TabOwner
        {
            get
            {
                if (ItemsControlAutomationPeer != null)
                {
                    return ItemsControlAutomationPeer.Owner as TabControl;
                }
                else
                {
                    return null;
                }
            }
        }
        
        /// <summary>
        /// Raise the event for when the IsSelectedProperty changes.
        /// </summary>
        /// <param name="isSelected">Inherited code: Requires comment.</param>
        internal void RaiseAutomationIsSelectedChanged(bool isSelected)
        {
            RaisePropertyChangedEvent(
                SelectionItemPatternIdentifiers.IsSelectedProperty,
                !isSelected,
                isSelected);
        }
    }
}