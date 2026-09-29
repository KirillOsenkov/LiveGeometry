using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Reactive;

namespace DynamicGeometry;

/// <summary>
/// How the chrome takes its colors from the <see cref="AppTheme"/>: a property is bound to the
/// resource of a color's name (the code side of DynamicResource), and follows the theme
/// from then on - a switch of theme, an edit of the color. Where a control changes color
/// with its state (a button's plate) it binds again to another name; the earlier binding
/// of the same property is let go of.
/// </summary>
public static class ThemeBinding
{
    static readonly ConditionalWeakTable<Control, Dictionary<AvaloniaProperty, IDisposable>> bindings =
        new ConditionalWeakTable<Control, Dictionary<AvaloniaProperty, IDisposable>>();

    /// <summary>
    /// Binds the property to the theme color of that name (<c>nameof(AppTheme.Text)</c>). A null
    /// key unbinds, and the property is then <paramref name="whenNone"/>, or cleared.
    /// </summary>
    public static void BindTheme(this Control control, AvaloniaProperty property, string key, object whenNone = null)
    {
        var table = bindings.GetOrCreateValue(control);
        if (table.Remove(property, out var previous))
        {
            previous.Dispose();
        }

        if (key == null)
        {
            if (whenNone != null)
            {
                control.SetValue(property, whenNone);
            }
            else
            {
                control.ClearValue(property);
            }

            return;
        }

        table[property] = control.Bind(property, control.GetResourceObservable(key));
    }

    /// <summary>
    /// Runs the action with the theme color of that name, now and whenever it changes - for
    /// what can't be bound: a gradient built from theme colors, a pen made while drawing.
    /// </summary>
    public static IDisposable ObserveTheme(this Control control, string key, Action<Color> action)
    {
        return control.GetResourceObservable(key).Subscribe(new AnonymousObserver<object>(value =>
        {
            if (value is ISolidColorBrush brush)
            {
                action(brush.Color);
            }
        }));
    }
}
