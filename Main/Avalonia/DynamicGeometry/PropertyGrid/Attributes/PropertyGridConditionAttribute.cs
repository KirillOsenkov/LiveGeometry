using System;
using System.Reflection;

namespace DynamicGeometry;

/// <summary>
/// On a button's method: the button is made only while the named public method of the
/// object (no parameters, a bool) answers true. For a verb every figure has but not every
/// figure can do right now (Bring to front), where <see cref="IConditionalProperties"/>
/// would have to be implemented by every figure.
/// </summary>
[AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
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
        if (condition == null)
        {
            return true;
        }

        var answer = target.GetType().GetMethod(condition.MethodName, Type.EmptyTypes);
        return answer == null || (bool)answer.Invoke(target, null);
    }
}
