namespace DynamicGeometry;

/// <summary>
/// How an expression's bound tree (<see cref="BoundExpression"/>, what
/// <see cref="ExpressionTreeBuilder"/> makes of a text) becomes the delegate that is run at
/// every recalculation: <see cref="Compiler.Strategy"/>. The benchmark
/// (LiveGeometry's ExpressionBenchmark, "--bench-expressions") prints what each costs to make
/// and to evaluate; they all give the same values, which the regression suite checks.
/// </summary>
public enum ExpressionStrategy
{
    /// <summary>The bound tree evaluates itself: no System.Linq.Expressions at all</summary>
    OwnTree,

    /// <summary>
    /// A System.Linq.Expressions tree run by that library's own interpreter
    /// (Compile(preferInterpretation: true)): what every expression did until 2026-10-07
    /// </summary>
    LightCompiler,

    /// <summary>A System.Linq.Expressions tree walked by <see cref="ExpressionTreeInterpreter"/></summary>
    LinqInterpreter,

    /// <summary>
    /// A System.Linq.Expressions tree compiled to IL (Compile()): a dynamic method for the
    /// JIT on the desktop; the browser has no JIT and interprets it as LightCompiler does
    /// </summary>
    Compiled
}
