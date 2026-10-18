using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using GuiLabs.Undo;

namespace LiveGeometry;

/// <summary>
/// An on/off row of the header (<see cref="PropertyGridSection.Header"/>) as a small button
/// with an icon: Visible an eye, open or shut, Locked a padlock, shut or open. The state a
/// figure usually has (shown, not locked) is faint, the other one plain; several figures
/// that differ show the other one fainter still, and a click gives them all that one.
/// Each click is an undo step, as the check box's was.
/// </summary>
public class HeaderToggle : Button
{
    // Avalonia matches theme templates by concrete type; keep using the Button template.
    protected override System.Type StyleKeyOverride => typeof(Button);

    readonly IValueProvider value;
    readonly ActionManager actionManager;

    public HeaderToggle(IValueProvider value, ActionManager actionManager)
    {
        this.value = value;
        this.actionManager = actionManager;
        Width = 26;
        Height = 26;
        Padding = new Thickness(0);
        Background = Brushes.Transparent;
        BorderThickness = new Thickness(0);
        HorizontalContentAlignment = HorizontalAlignment.Center;
        VerticalContentAlignment = VerticalAlignment.Center;
        Click += (s, e) => Toggle();
        AttachedToVisualTree += (s, e) => value.ValueChanged += Update;
        DetachedFromVisualTree += (s, e) => value.ValueChanged -= Update;
        Update();
    }

    bool IsLock => value.Name == nameof(IFigure.Locked);

    // null: several figures that differ
    bool? State => value.GetValue<object>() as bool?;

    void Toggle()
    {
        bool next = State != true;
        if (actionManager != null)
        {
            Actions.SetProperty(actionManager, value, next);
        }
        else
        {
            value.SetValue(next);
        }

        Update();
    }

    void Update()
    {
        var state = State;
        bool usual = IsLock ? state == false : state == true;
        var icon = IsLock
            ? state == true ? PropertyGridIcon.Lock : PropertyGridIcon.Unlock
            : state == false ? PropertyGridIcon.EyeHidden : PropertyGridIcon.Eye;
        var content = PropertyGridIcons.Create(icon);
        content.Opacity = state == null ? 0.35 : usual ? 0.5 : 1;
        Content = content;
        IsEnabled = value.CanSetValue;
        ToolTip.SetTip(this, IsLock
            ? state == true ? "Locked: click to unlock" : state == false ? "Click to lock" : "Some are locked: click to lock all"
            : state == true ? "Shown: click to hide" : state == false ? "Hidden: click to show" : "Some are hidden: click to show all");
    }
}
