using System.Collections.Generic;

namespace LiveGeometry;

/// <summary>A function the completion list offers: its name, what it takes, and a line of help</summary>
public record FunctionInfo(string Name, string Parameters, string Help)
{
    /// <summary>"sin(x)", "angle(A, B, C)"</summary>
    public string Signature => Name + "(" + Parameters + ")";

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
/// The functions of the expression language as the completion list offers them
/// (docs/completion-plan.md), written by hand: a line of help each, in the order of the
/// groups. Not every method of System.Math the language can call (BigMul, ILogB, ScaleB...),
/// and not the names VB6's drawings use for the same thing (distance, sgn, int, sqr, todeg,
/// torad, xang): those still compile. The regression suite checks that each one resolves
/// (<see cref="ExpressionReflection.ResolveMethod"/>), so the list can't offer what isn't there.
/// </summary>
public static class FunctionCatalog
{
    public static IReadOnlyList<FunctionInfo> All { get; } =
    [
        new("sin", "x", "The sine of an angle x in radians: sin(rad(30)) for 30°."),
        new("cos", "x", "The cosine of an angle x in radians."),
        new("tan", "x", "The tangent of an angle x in radians."),
        new("asin", "x", "The angle whose sine is x, in radians."),
        new("acos", "x", "The angle whose cosine is x, in radians."),
        new("atan", "x", "The angle whose tangent is x, in radians."),
        new("atan2", "y, x", "The angle of the point (x, y) from the x axis, in radians."),
        new("sinh", "x", "The hyperbolic sine of x."),
        new("cosh", "x", "The hyperbolic cosine of x."),
        new("tanh", "x", "The hyperbolic tangent of x."),
        new("asinh", "x", "The inverse hyperbolic sine of x."),
        new("acosh", "x", "The inverse hyperbolic cosine of x."),
        new("atanh", "x", "The inverse hyperbolic tangent of x."),
        new("sqrt", "x", "The square root of x."),
        new("cbrt", "x", "The cube root of x."),
        new("pow", "x, y", "x to the power y, the same as x ^ y."),
        new("exp", "x", "e to the power x."),
        new("ln", "x", "The natural logarithm of x (to base e)."),
        new("log", "x", "The natural logarithm of x, the same as ln(x)."),
        new("lg", "x", "The logarithm of x to base 10."),
        new("log2", "x", "The logarithm of x to base 2."),
        new("abs", "x", "x without its sign: abs(-3) is 3."),
        new("round", "x", "x rounded to a whole number, a half up: round(2.5) is 3."),
        new("floor", "x", "The whole number at or below x: floor(2.7) is 2."),
        new("ceiling", "x", "The whole number at or above x: ceiling(2.1) is 3."),
        new("truncate", "x", "x without what follows the point: truncate(-2.7) is -2."),
        new("sign", "x", "1 for a positive x, -1 for a negative one, 0 for 0."),
        new("min", "a, b", "The smaller of a and b."),
        new("max", "a, b", "The larger of a and b."),
        new("clamp", "x, low, high", "x, but no less than low and no more than high."),
        new("deg", "x", "An angle x in radians, in degrees: deg(pi) is 180."),
        new("rad", "x", "An angle x in degrees, in radians: rad(180) is pi."),
        new("angle", "A, B, C", "The angle ABC at B in degrees, the smaller one, as the Angle tool measures it."),
        new("ang", "A, B, C", "The angle ABC at B in radians, the smaller one."),
        new("oang", "A, B, C", "The angle ABC at B in radians, counterclockwise from BA to BC."),
        new("dist", "A, B", "The distance between points A and B; AB says the same."),
        new("area", "A, B, C, ...", "The area of the polygon on these points.")
    ];
}
