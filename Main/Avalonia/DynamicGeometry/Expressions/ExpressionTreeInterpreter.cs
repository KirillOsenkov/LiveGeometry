using System;
using System.Linq.Expressions;
using System.Reflection;

namespace DynamicGeometry
{
    /// <summary>
    /// Walks a System.Linq.Expressions tree (the one <see cref="BoundExpression.ToLinq"/>
    /// makes) at every evaluation: <see cref="ExpressionStrategy.LinqInterpreter"/>. Nothing
    /// is compiled, so the delegate costs nothing to make; evaluating boxes each value and
    /// looks each node up. Methods and properties are called through the delegates of
    /// <see cref="ExpressionReflection"/>, not by MethodInfo.Invoke. A tree with a node it
    /// doesn't know is refused when the delegate is made (NotSupportedException, which the
    /// compiler turns into an error of the expression), not at the first evaluation.
    /// </summary>
    public class ExpressionTreeInterpreter : IExpressionTreeEvaluatorProvider
    {
        public T InterpretFunction<T>(Expression<T> node)
        {
            Validate(node.Body);
            var parameter = node.Parameters.Count > 0 ? node.Parameters[0] : null;
            var body = node.Body;
            Func<double, double> function = x => ExpressionReflection.ToDouble(Evaluate(body, parameter, x));
            return (T)(object)function;
        }

        public T InterpretExpression<T>(Expression<T> node)
        {
            Validate(node.Body);
            var body = node.Body;
            Func<double> expression = () => ExpressionReflection.ToDouble(Evaluate(body, null, 0));
            return (T)(object)expression;
        }

        static void Validate(Expression expression)
        {
            switch (expression)
            {
                case BinaryExpression binary when IsArithmetic(binary.NodeType):
                    Validate(binary.Left);
                    Validate(binary.Right);
                    return;
                case UnaryExpression unary when unary.NodeType == ExpressionType.Negate || unary.NodeType == ExpressionType.Convert:
                    Validate(unary.Operand);
                    return;
                case ConstantExpression _:
                case ParameterExpression _:
                    return;
                case MemberExpression member when member.Member is PropertyInfo || member.Member is FieldInfo:
                    if (member.Expression != null)
                    {
                        Validate(member.Expression);
                    }

                    return;
                case MethodCallExpression call when call.Object == null:
                    foreach (var argument in call.Arguments)
                    {
                        Validate(argument);
                    }

                    return;
                default:
                    throw new NotSupportedException("The interpreter can't evaluate " + expression.NodeType + ".");
            }
        }

        static bool IsArithmetic(ExpressionType type)
        {
            return type == ExpressionType.Add
                || type == ExpressionType.Subtract
                || type == ExpressionType.Multiply
                || type == ExpressionType.Divide
                || type == ExpressionType.Power;
        }

        /// <param name="parameter">The x of a function, null for a plain expression</param>
        /// <param name="value">What x is</param>
        static object Evaluate(Expression expression, ParameterExpression parameter, double value)
        {
            switch (expression.NodeType)
            {
                case ExpressionType.Add:
                    Operands(expression, parameter, value, out var addLeft, out var addRight);
                    return addLeft + addRight;
                case ExpressionType.Subtract:
                    Operands(expression, parameter, value, out var subtractLeft, out var subtractRight);
                    return subtractLeft - subtractRight;
                case ExpressionType.Multiply:
                    Operands(expression, parameter, value, out var multiplyLeft, out var multiplyRight);
                    return multiplyLeft * multiplyRight;
                case ExpressionType.Divide:
                    Operands(expression, parameter, value, out var divideLeft, out var divideRight);
                    return divideLeft / divideRight;
                case ExpressionType.Power:
                    Operands(expression, parameter, value, out var powerLeft, out var powerRight);
                    return System.Math.Pow(powerLeft, powerRight);
                case ExpressionType.Negate:
                    return -ExpressionReflection.ToDouble(Evaluate(((UnaryExpression)expression).Operand, parameter, value));
                case ExpressionType.Convert:
                    var converted = Evaluate(((UnaryExpression)expression).Operand, parameter, value);
                    return expression.Type == typeof(double) ? ExpressionReflection.ToDouble(converted) : Convert.ChangeType(converted, expression.Type);
                case ExpressionType.Constant:
                    return ((ConstantExpression)expression).Value;
                case ExpressionType.Parameter:
                    if (expression != parameter)
                    {
                        throw new InvalidOperationException("An expression has no x.");
                    }

                    return value;
                case ExpressionType.MemberAccess:
                    return Member((MemberExpression)expression, parameter, value);
                case ExpressionType.Call:
                    var call = (MethodCallExpression)expression;
                    var arguments = new object[call.Arguments.Count];
                    for (int i = 0; i < arguments.Length; i++)
                    {
                        arguments[i] = Evaluate(call.Arguments[i], parameter, value);
                    }

                    return ExpressionReflection.Invoke(call.Method, arguments);
                default:
                    throw new NotSupportedException("The interpreter can't evaluate " + expression.NodeType + ".");
            }
        }

        static void Operands(Expression expression, ParameterExpression parameter, double value, out double left, out double right)
        {
            var binary = (BinaryExpression)expression;
            left = ExpressionReflection.ToDouble(Evaluate(binary.Left, parameter, value));
            right = ExpressionReflection.ToDouble(Evaluate(binary.Right, parameter, value));
        }

        static object Member(MemberExpression member, ParameterExpression parameter, double value)
        {
            var instance = member.Expression == null ? null : Evaluate(member.Expression, parameter, value);
            if (member.Member is PropertyInfo property)
            {
                // a number through its getter's delegate; anything else (a point's
                // coordinates, on their way into dist or ang) as it is
                return ExpressionReflection.IsNumberType(property.PropertyType)
                    ? ExpressionReflection.GetNumberGetter(property)(instance)
                    : property.GetValue(instance);
            }

            return ((FieldInfo)member.Member).GetValue(instance);
        }
    }
}
