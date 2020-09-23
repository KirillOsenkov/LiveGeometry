using System.Linq.Expressions;

namespace DynamicGeometry
{
    public class ExpressionTreeCompiler : IExpressionTreeEvaluatorProvider
    {
        public T InterpretFunction<T>(Expression<T> node)
        {
            return node.Compile();
        }

        public T InterpretExpression<T>(Expression<T> node)
        {
            return node.Compile();
        }
    }
}
