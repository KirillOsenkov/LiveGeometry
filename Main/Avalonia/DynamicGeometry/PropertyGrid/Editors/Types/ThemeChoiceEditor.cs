using System.Linq;

namespace DynamicGeometry;

/// <summary>
/// Picked by <c>[PropertyGridPreferredEditor("ThemeChoice")]</c> on a string: a combo box of
/// <see cref="AppTheme.SystemChoice"/> and the names of every theme there is.
/// </summary>
public class ThemeChoiceEditorFactory : BaseValueEditorFactory<ThemeChoiceEditor, string>
{
    public ThemeChoiceEditorFactory()
    {
        // after the plain string editor: only the attribute chooses this one
        LoadOrder = 1;
    }
}

public class ThemeChoiceEditor : SelectorValueEditor, IValueEditor
{
    public override void FillList()
    {
        Items = new[] { AppTheme.SystemChoice }.Concat(AppTheme.All.Select(theme => theme.Name)).ToArray();
        base.FillList();
    }
}
