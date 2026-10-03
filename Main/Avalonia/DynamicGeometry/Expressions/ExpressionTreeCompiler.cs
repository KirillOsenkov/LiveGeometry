using System.Linq.Expressions;

namespace DynamicGeometry
{
    public class ExpressionTreeCompiler : IExpressionTreeEvaluatorProvider
    {
        public T InterpretFunction<T>(Expression<T> node)
        {
            return node.Compile();
        }

        // Interpreted, as the browser always does: compiled, each expression was a dynamic
        // method for the JIT, and a gallery drawing with hundreds of points by coordinates
        // spent most of its load there (a function is evaluated far more often, and stays
        // compiled)
        public T InterpretExpression<T>(Expression<T> node)
        {
            return node.Compile(preferInterpretation: true);
        }
    }
}
