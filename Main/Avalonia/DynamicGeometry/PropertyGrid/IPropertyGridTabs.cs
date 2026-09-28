using System.Collections.Generic;
using GuiLabs.Undo;

namespace DynamicGeometry;

/// <summary>
/// An object whose rows the property grid spreads over tabs (a point style: Shape | Emoji).
/// The tabs are two ways of being the same thing: the tab shown is the one that says what the
/// object is now, and picking another may change the object.
/// </summary>
public interface IPropertyGridTabs
{
    /// <summary>The captions, in order</summary>
    IReadOnlyList<string> Tabs { get; }

    /// <summary>
    /// The tabs a property or a button is on, by member name (a size can be on two); none:
    /// below the tabs, whichever is shown (Done)
    /// </summary>
    IReadOnlyList<string> GetTabs(string memberName);

    /// <summary>The tab that says what the object is now (the style draws a character: Emoji)</summary>
    string CurrentTab { get; }

    /// <summary>
    /// The user picked a tab. Change the object if the tab means it (Shape: the character
    /// goes), through the action manager when there is one.
    /// </summary>
    void OnTabSelected(string tab, ActionManager actionManager);
}
