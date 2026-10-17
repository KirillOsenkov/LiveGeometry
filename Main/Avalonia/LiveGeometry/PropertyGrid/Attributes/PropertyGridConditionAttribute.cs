using System;
using System.Reflection;

namespace LiveGeometry;

/// <summary>
/// On a button's method: the button is made only while the named public method of the
/// object (no parameters, a bool) answers true. For a verb every figure has but not every
/// figure can do right now (Bring to front), where <see cref="IConditionalProperties"/>
/// would have to be implemented by every figure.
/// On a property: the row shows while the method answers true, asked again whenever the
/// object raises a change with a property name - the row is hidden, not taken away, so a
/// box being typed in keeps the keyboard (a label's Decimals, which appear with the text's
/// first [...] part).
/// </summary>
[AttributeUsage(AttributeTargets.Method | AttributeTargets.Property, AllowMultiple = false)]
public class PropertyGridConditionAttribute : Attribute
{
    public PropertyGridConditionAttribute(string methodName)
    {
        MethodName = methodName;
    }

    public string MethodName { get; }

    /// <summary>Whether the object's condition for the method holds; true for a method without one</summary>
    public static bool Holds(MethodInfo method, object target)
    {
        var condition = method.GetCustomAttribute<PropertyGridConditionAttribute>();
        return condition == null || condition.Holds(target);
    }

    /// <summary>Whether the named method of the object answers true; true for an object without it</summary>
    public bool Holds(object target)
    {
        var answer = target?.GetType().GetMethod(MethodName, Type.EmptyTypes);
        return answer == null || (bool)answer.Invoke(target, null);
    }
}
