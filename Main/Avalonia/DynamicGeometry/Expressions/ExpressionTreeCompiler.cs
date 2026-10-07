using System.Linq.Expressions;

namespace DynamicGeometry
{
    /// <summary>
    /// A System.Linq.Expressions tree made into a delegate by that library: compiled to IL
    /// (<see cref="ExpressionStrategy.Compiled"/>: a dynamic method for the JIT on the
    /// desktop; the browser has no JIT and interprets it instead), or run by its interpreter
    /// (<see cref="ExpressionStrategy.LightCompiler"/>: Compile(preferInterpretation: true),
    /// which builds an instruction list and boxes each value as it runs).
    /// </summary>
    public class ExpressionTreeCompiler : IExpressionTreeEvaluatorProvider
    {
        readonly bool preferInterpretation;

        public ExpressionTreeCompiler(bool preferInterpretation)
        {
            this.preferInterpretation = preferInterpretation;
        }

        public T InterpretFunction<T>(Expression<T> node)
        {
            return node.Compile(preferInterpretation);
        }

        public T InterpretExpression<T>(Expression<T> node)
        {
            return node.Compile(preferInterpretation);
        }
    }
}
