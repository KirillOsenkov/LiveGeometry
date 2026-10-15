using System.Linq.Expressions;

namespace LiveGeometry;

public interface IExpressionTreeEvaluatorProvider
{
    T InterpretFunction<T>(Expression<T> node);
    T InterpretExpression<T>(Expression<T> node);
}
