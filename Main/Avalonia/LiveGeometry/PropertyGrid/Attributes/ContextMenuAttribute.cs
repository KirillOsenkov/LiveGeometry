using System;

namespace LiveGeometry;

/// <summary>
/// Offers a row or a button of the property grid in the figure's context menu too: an on/off
/// row (a bool, "Show name") as a check item, a button ("Free point") as an item. The menu
/// asks the same questions the grid does - the row or button must be on the grid's page,
/// editable, its condition holding (<see cref="IConditionalProperties"/>,
/// <see cref="PropertyGridConditionAttribute"/>) - and takes the grid's caption and icon, so
/// the grid stays where a figure's verbs are defined. Only for what is handy at the cursor:
/// the menu is not a second grid. (<see cref="ContextMenus"/>)
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Method, Inherited = true, AllowMultiple = false)]
public class ContextMenuAttribute : Attribute
{
}
