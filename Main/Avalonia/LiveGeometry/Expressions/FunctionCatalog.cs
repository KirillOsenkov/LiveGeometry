using System;
using System.Collections.Generic;
using System.Linq;

namespace LiveGeometry;

/// <summary>A function of the expression language: its name, what it takes, a line of help, and the other names it goes by</summary>
public record FunctionInfo(string Name, string Parameters, string Help, string[] Aliases = null)
{
    /// <summary>"sin(x)", "angle(A, B, C)"</summary>
    public string Signature => Name + "(" + Parameters + ")";

    /// <summary>The names it is also called by, which the completion list matches too: asin for arcsin</summary>
    public IReadOnlyList<string> OtherNames => Aliases ?? [];

    /// <summary>It takes any number of arguments, two at least: max(a, b, ...)</summary>
    public bool IsVariadic => Parameters.EndsWith("...", StringComparison.Ordinal);

    /// <summary>How many arguments it is called with, for the check that the language has it ("A, B, C, ...": the three it names)</summary>
    public int ArgumentCount
    {
        get
        {
            var parameters = Parameters.Split(", ");
            return parameters[^1] == "..." ? parameters.Length - 1 : parameters.Length;
        }
    }
}

/// <summary>
/// The functions of the expression language (2026-10-18): this list is the language. A name
/// here, or one of its other names, is the method of <see cref="Functions"/> by that name;
/// nothing else is a function (System.Math was searched by name, and BitIncrement or
/// IEEERemainder compiled). The names are the American classroom's and Desmos's where they
/// differ from others: log is base 10 and ln is e's, arcsin, ceil, distance. log_2(x) and
/// any log_ with a whole number say the base in the name (<see cref="ExpressionReflection.LogBaseOf"/>).
/// In the order of the groups, a line of help each, for the completion list.
/// </summary>
public static class FunctionCatalog
{
    public static IReadOnlyList<FunctionInfo> All { get; } =
    [
        new("sin", "x", "The sine of an angle x in radians: sin(30°) for 30 degrees."),
        new("cos", "x", "The cosine of an angle x in radians: cos(60°) for 60 degrees."),
        new("tan", "x", "The tangent of an angle x in radians: tan(45°) for 45 degrees."),
        new("sec", "x", "The secant of an angle x in radians: 1 / cos(x)."),
        new("csc", "x", "The cosecant of an angle x in radians: 1 / sin(x)."),
        new("cot", "x", "The cotangent of an angle x in radians: 1 / tan(x)."),
        new("arcsin", "x", "The angle whose sine is x, in radians: deg(arcsin(x)) for degrees.", ["asin"]),
        new("arccos", "x", "The angle whose cosine is x, in radians: deg(arccos(x)) for degrees.", ["acos"]),
        new("arctan", "x", "The angle whose tangent is x, in radians: deg(arctan(x)) for degrees.", ["atan"]),
        new("atan2", "y, x", "The angle of the point (x, y) from the x axis, in radians."),
        new("sinh", "x", "The hyperbolic sine of x."),
        new("cosh", "x", "The hyperbolic cosine of x."),
        new("tanh", "x", "The hyperbolic tangent of x."),
        new("arcsinh", "x", "The inverse hyperbolic sine of x.", ["asinh"]),
        new("arccosh", "x", "The inverse hyperbolic cosine of x.", ["acosh"]),
        new("arctanh", "x", "The inverse hyperbolic tangent of x.", ["atanh"]),
        new("sqrt", "x", "The square root of x."),
        new("cbrt", "x", "The cube root of x: cbrt(-8) is -2."),
        new("nthroot", "x, n", "The nth root of x: nthroot(16, 4) is 2."),
        new("exp", "x", "e to the power x, the same as e^x."),
        new("ln", "x", "The natural logarithm of x, to base e."),
        new("log", "x", "The logarithm of x to base 10: log(1000) is 3. Another base goes after _: log_2(8) is 3."),
        new("abs", "x", "x without its sign, the same as |x|: abs(-3) is 3."),
        new("round", "x, digits", "x rounded, a half up: round(2.5) is 3, round(3.14159, 2) is 3.14."),
        new("floor", "x", "The whole number at or below x: floor(2.7) is 2."),
        new("ceil", "x", "The whole number at or above x: ceil(2.1) is 3.", ["ceiling"]),
        new("sign", "x", "1 for a positive x, -1 for a negative one, 0 for 0."),
        new("mod", "a, b", "What is left when a is divided by b: mod(7, 3) is 1, mod(-1, 3) is 2."),
        new("min", "a, b, ...", "The smallest of the numbers."),
        new("max", "a, b, ...", "The largest of the numbers."),
        new("clamp", "x, low, high", "x, but no less than low and no more than high."),
        new("gcd", "a, b, ...", "The greatest common factor of whole numbers: gcd(12, 18) is 6."),
        new("lcm", "a, b, ...", "The least common multiple of whole numbers: lcm(4, 6) is 12."),
        new("deg", "x", "An angle x in radians, in degrees: deg(pi) is 180."),
        new("rad", "x", "An angle x in degrees, in radians, the same as x°: rad(180) is pi."),
        new("if", "condition, a, b", "a where the condition holds, else b: if(x < 0, -x, x). Without b, nothing where it doesn't."),
        new("distance", "A, B", "The distance between points A and B; AB says the same."),
        new("angle", "A, B, C", "The angle ABC at B in degrees, the smaller one, as the Angle tool measures it."),
        new("directedangle", "A, B, C", "The angle ABC at B in degrees, counterclockwise from BA to BC: from 0 up to 360."),
        new("direction", "A, B", "The direction from A to B in degrees, counterclockwise from the x axis."),
        new("area", "A, B, C, ...", "The area of the polygon on these points.")
    ];

    static readonly Dictionary<string, FunctionInfo> byName = All
        .SelectMany(function => function.OtherNames.Prepend(function.Name).Select(name => (name, function)))
        .ToDictionary(pair => pair.name, pair => pair.function, StringComparer.OrdinalIgnoreCase);

    /// <summary>The function a name calls, by its name or another, in any case; null when the language has none</summary>
    public static FunctionInfo Find(string name)
    {
        return name != null && byName.TryGetValue(name, out var function) ? function : null;
    }

    /// <summary>The trigonometric functions, which take radians: a degree-valued argument there is a mistake (<see cref="GivesDegrees"/>)</summary>
    public static bool TakesRadians(FunctionInfo function)
    {
        return function?.Name is "sin" or "cos" or "tan" or "sec" or "csc" or "cot";
    }

    /// <summary>The functions of points that give degrees, as the drawing shows angles</summary>
    public static bool GivesDegrees(FunctionInfo function)
    {
        return function?.Name is "angle" or "directedangle" or "direction";
    }
}
