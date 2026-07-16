using System;
using System.Linq.Expressions;

namespace DynamicGeometry
{
    public interface IExpressionTreeEvaluatorProvider
    {
        T InterpretFunction<T>(Expression<T> node);
        T InterpretExpression<T>(Expression<T> node);
    }
}
