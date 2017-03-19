using System;
using System.ComponentModel.Composition;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace DynamicGeometry
{
    [Export(typeof(IExpressionTreeEvaluatorProvider))]
    public class ExpressionTreeInterpreter : IExpressionTreeEvaluatorProvider
    {
        private object parameterValue = null;

        public T InterpretFunction<T>(Expression<T> node)
        {
            Func<double, double> result = x => Evaluate(x, node);
            return (T)(object)result;
        }

        public T InterpretExpression<T>(Expression<T> node)
        {
            var body = node.Body;
            Func<double> result = () => (double)Evaluate(body);
            return (T)(object)result;
        }

        private double Evaluate(double parameter, LambdaExpression lambda)
        {
            parameterValue = parameter;
            return (double)Evaluate(lambda.Body);
        }

        private object Evaluate(Expression expression)
        {
            object result = null;

            if (expression == null)
            {
                return null;
            }

            switch (expression.NodeType)
            {
                case ExpressionType.Add:
                    var addExpression = (BinaryExpression)expression;
                    var addLeft = Evaluate(addExpression.Left);
                    var addRight = Evaluate(addExpression.Right);
                    result = (double)addLeft + (double)addRight;
                    break;
                case ExpressionType.AddChecked:
                    break;
                case ExpressionType.And:
                    break;
                case ExpressionType.AndAlso:
                    break;
                case ExpressionType.ArrayIndex:
                    break;
                case ExpressionType.ArrayLength:
                    break;
                case ExpressionType.Call:
                    var methodCallExpression = (MethodCallExpression)expression;
                    var instance = Evaluate(methodCallExpression.Object);
                    var arguments = methodCallExpression.Arguments.Select(a => Evaluate(a)).ToArray();
                    result = methodCallExpression.Method.Invoke(instance, arguments);
                    break;
                case ExpressionType.Coalesce:
                    break;
                case ExpressionType.Conditional:
                    break;
                case ExpressionType.Constant:
                    var constant = (ConstantExpression)expression;
                    result = constant.Value;
                    break;
                case ExpressionType.Convert:
                    break;
                case ExpressionType.ConvertChecked:
                    break;
                case ExpressionType.Divide:
                    var divideExpression = (BinaryExpression)expression;
                    var divideLeft = Evaluate(divideExpression.Left);
                    var divideRight = Evaluate(divideExpression.Right);
                    result = (double)divideLeft / (double)divideRight;
                    break;
                case ExpressionType.Equal:
                    break;
                case ExpressionType.ExclusiveOr:
                    break;
                case ExpressionType.GreaterThan:
                    break;
                case ExpressionType.GreaterThanOrEqual:
                    break;
                case ExpressionType.Invoke:
                    break;
                case ExpressionType.Lambda:
                    break;
                case ExpressionType.LeftShift:
                    break;
                case ExpressionType.LessThan:
                    break;
                case ExpressionType.LessThanOrEqual:
                    break;
                case ExpressionType.ListInit:
                    break;
                case ExpressionType.MemberAccess:
                    var memberExpression = (MemberExpression)expression;
                    var memberParent = Evaluate(memberExpression.Expression);
                    PropertyInfo property = memberExpression.Member as PropertyInfo;
                    if (property != null)
                    {
                        result = property.GetValue(memberParent, null);
                    }

                    FieldInfo fieldInfo = memberExpression.Member as FieldInfo;
                    if (fieldInfo != null)
                    {
                        result = fieldInfo.GetValue(memberParent);
                    }

                    break;
                case ExpressionType.MemberInit:
                    break;
                case ExpressionType.Modulo:
                    break;
                case ExpressionType.Multiply:
                    var multiplyExpression = (BinaryExpression)expression;
                    var multiplyLeft = Evaluate(multiplyExpression.Left);
                    var multiplyRight = Evaluate(multiplyExpression.Right);
                    result = (double)multiplyLeft * (double)multiplyRight;
                    break;
                case ExpressionType.MultiplyChecked:
                    break;
                case ExpressionType.Negate:
                    var unaryExpression = (UnaryExpression)expression;
                    var operand = Evaluate(unaryExpression.Operand);
                    result = unaryExpression.Method.Invoke(null, new[] { operand });
                    break;
                case ExpressionType.NegateChecked:
                    break;
                case ExpressionType.New:
                    break;
                case ExpressionType.NewArrayBounds:
                    break;
                case ExpressionType.NewArrayInit:
                    break;
                case ExpressionType.Not:
                    break;
                case ExpressionType.NotEqual:
                    break;
                case ExpressionType.Or:
                    break;
                case ExpressionType.OrElse:
                    break;
                case ExpressionType.Parameter:
                    result = parameterValue;
                    break;
                case ExpressionType.Power:
                    var powerExpression = (BinaryExpression)expression;
                    var powerLeft = Evaluate(powerExpression.Left);
                    var powerRight = Evaluate(powerExpression.Right);
                    result = (double)System.Math.Pow((double)powerLeft, (double)powerRight);
                    break;
                case ExpressionType.Quote:
                    break;
                case ExpressionType.RightShift:
                    break;
                case ExpressionType.Subtract:
                    var subtractExpression = (BinaryExpression)expression;
                    var subtractLeft = Evaluate(subtractExpression.Left);
                    var subtractRight = Evaluate(subtractExpression.Right);
                    result = (double)subtractLeft - (double)subtractRight;
                    break;
                case ExpressionType.SubtractChecked:
                    break;
                case ExpressionType.TypeAs:
                    break;
                case ExpressionType.TypeIs:
                    break;
                case ExpressionType.UnaryPlus:
                    break;
                default:
                    break;
            }

            return result;
        }
    }
}
