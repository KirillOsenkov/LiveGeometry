using System;
using System.Collections.Generic;

namespace DynamicGeometry;

/// <summary>
/// A property of an <see cref="IThemeOverridable"/> as seen under one theme: reading gives
/// the override for that theme when there is one, else the property's own value; writing
/// sets the override. What the property grid edits under a theme other than the base one,
/// so that a change lands in the theme on screen (and undo puts the override back), and
/// what the serializer reads and writes a theme's child element through.
/// </summary>
public class ThemedValue : IValueProvider
{
    readonly IValueProvider inner;
    readonly IThemeOverridable target;
    readonly string theme;

    public ThemedValue(IValueProvider inner, IThemeOverridable target, string theme)
    {
        this.inner = inner;
        this.target = target;
        this.theme = theme;
        inner.ValueChanged += RaiseValueChanged;
    }

    /// <summary>Whether the theme has a value of its own for the property</summary>
    public bool HasOverride
    {
        get
        {
            return target.Overrides.TryGetValue(theme, out var values) && values.ContainsKey(inner.Name);
        }
    }

    public event Action ValueChanged;

    public void RaiseValueChanged()
    {
        ValueChanged?.Invoke();
    }

    public T GetValue<T>()
    {
        if (target.Overrides.TryGetValue(theme, out var values) && values.TryGetValue(inner.Name, out var value))
        {
            return (T)value;
        }

        return inner.GetValue<T>();
    }

    public bool CanSetValue => inner.CanSetValue;

    public void SetValue<T>(T value)
    {
        target.SetOverride(theme, inner.Name, value);
    }

    public Type Type => inner.Type;

    public object Parent => inner.Parent;

    public string Name => inner.Name;

    public string DisplayName => inner.DisplayName;

    public T GetAttribute<T>() where T : Attribute
    {
        return inner.GetAttribute<T>();
    }

    public IEnumerable<T> GetAttributes<T>() where T : Attribute
    {
        return inner.GetAttributes<T>();
    }

    public string GetSignature()
    {
        return inner.GetSignature();
    }

    /// <summary>
    /// The value the grid edits: the property itself under the base theme, its override
    /// under any other, when the property belongs to something that has overrides
    /// </summary>
    public static IValueProvider ForCurrentTheme(IValueProvider value)
    {
        if (value.Parent is IThemeOverridable target && !AppTheme.IsBase(AppTheme.Current))
        {
            return new ThemedValue(value, target, AppTheme.Current.Name);
        }

        return value;
    }
}
