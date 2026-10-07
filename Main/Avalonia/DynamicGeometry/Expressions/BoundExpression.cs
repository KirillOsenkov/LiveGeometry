using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace DynamicGeometry;

/// <summary>
/// An expression with its names resolved: what <see cref="ExpressionTreeBuilder"/> makes of
/// a text, and what every <see cref="ExpressionStrategy"/> starts from. It evaluates itself
/// (<see cref="Evaluate"/>: the OwnTree strategy, no System.Linq.Expressions anywhere), or is
/// turned into a System.Linq.Expressions tree (<see cref="ToLinq"/>) for the strategies that
/// compile or interpret one. Every value is a double: what is no number was refused when the
/// tree was bound (<see cref="ExpressionTreeBuilder.AsNumber"/>), and a whole number is
/// converted.
/// </summary>
public abstract class BoundExpression
{
    /// <param name="parameter">The x of a function; a plain expression has none and ignores it</param>
    public abstract double Evaluate(double parameter);

    /// <param name="parameter">The x of a function; null for a plain expression</param>
    public abstract Expression ToLinq(ParameterExpression parameter);

    /// <summary>The delegate of a plain expression, evaluating this tree</summary>
    public Func<double> ToExpressionDelegate()
    {
        return () => Evaluate(0);
    }

    /// <summary>The delegate of a function of x, evaluating this tree</summary>
    public Func<double, double> ToFunctionDelegate()
    {
        return parameter => Evaluate(parameter);
    }
}

public class BoundConstant : BoundExpression
{
    public BoundConstant(double value)
    {
        Value = value;
    }

    public double Value { get; }

    public override double Evaluate(double parameter)
    {
        return Value;
    }

    public override Expression ToLinq(ParameterExpression parameter)
    {
        return Expression.Constant(Value);
    }
}

/// <summary>The x of a function</summary>
public class BoundParameter : BoundExpression
{
    public override double Evaluate(double parameter)
    {
        return parameter;
    }

    public override Expression ToLinq(ParameterExpression parameter)
    {
        return parameter ?? throw new InvalidOperationException("An expression has no x.");
    }
}

/// <summary>A number of the drawing (a slider, a Number) named in the expression: its value</summary>
public class BoundNumber : BoundExpression
{
    public BoundNumber(INumber number)
    {
        Number = number;
    }

    public INumber Number { get; }

    public override double Evaluate(double parameter)
    {
        return Number.Value;
    }

    public override Expression ToLinq(ParameterExpression parameter)
    {
        return Expression.Property(Expression.Constant(Number, typeof(INumber)), ExpressionReflection.NumberValue);
    }
}

/// <summary>Two points' names run together (AB): their distance</summary>
public class BoundDistance : BoundExpression
{
    public BoundDistance(PointBase first, PointBase second)
    {
        First = first;
        Second = second;
    }

    public PointBase First { get; }

    public PointBase Second { get; }

    public override double Evaluate(double parameter)
    {
        return Math.Distance(First, Second);
    }

    public override Expression ToLinq(ParameterExpression parameter)
    {
        return Expression.Call(ExpressionReflection.DistanceOfPoints, Expression.Constant(First), Expression.Constant(Second));
    }
}

/// <summary>A property of a figure (A.X, AB.Length, a.Value): a number</summary>
public class BoundProperty : BoundExpression
{
    readonly Func<object, double> getter;

    public BoundProperty(IFigure figure, PropertyInfo property)
    {
        Figure = figure;
        Property = property;
        getter = ExpressionReflection.GetNumberGetter(property);
    }

    public IFigure Figure { get; }

    public PropertyInfo Property { get; }

    public override double Evaluate(double parameter)
    {
        return getter(Figure);
    }

    public override Expression ToLinq(ParameterExpression parameter)
    {
        return ExpressionReflection.AsDouble(Expression.Property(Expression.Constant(Figure), Property));
    }
}

/// <summary>A function of numbers (sin, max, atan2, round): each argument is an expression</summary>
public class BoundCall : BoundExpression
{
    readonly BoundExpression[] arguments;
    readonly NumberFunction function;

    public BoundCall(MethodInfo method, IReadOnlyList<BoundExpression> arguments)
    {
        Method = method;
        this.arguments = arguments.ToArray();
        function = ExpressionReflection.GetNumberFunction(method);
    }

    public MethodInfo Method { get; }

    public IReadOnlyList<BoundExpression> Arguments => arguments;

    public override double Evaluate(double parameter)
    {
        switch (arguments.Length)
        {
            case 1:
                return function.Of1(arguments[0].Evaluate(parameter));
            case 2:
                return function.Of2(arguments[0].Evaluate(parameter), arguments[1].Evaluate(parameter));
            case 3:
                return function.Of3(arguments[0].Evaluate(parameter), arguments[1].Evaluate(parameter), arguments[2].Evaluate(parameter));
            default:
                var values = new double[arguments.Length];
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = arguments[i].Evaluate(parameter);
                }

                return function.OfMany(values);
        }
    }

    public override Expression ToLinq(ParameterExpression parameter)
    {
        return ExpressionReflection.AsDouble(Expression.Call(Method, arguments.Select(argument => argument.ToLinq(parameter))));
    }
}

/// <summary>A function of points (dist, ang, area), called with the names of points</summary>
public class BoundPointCall : BoundExpression
{
    readonly PointFunction function;

    public BoundPointCall(MethodInfo method, IReadOnlyList<IPoint> points)
    {
        Method = method;
        Points = points.ToArray();
        function = ExpressionReflection.GetPointFunction(method);
    }

    public MethodInfo Method { get; }

    public IPoint[] Points { get; }

    public override double Evaluate(double parameter)
    {
        return function.Invoke(Points);
    }

    public override Expression ToLinq(ParameterExpression parameter)
    {
        if (function.TakesArray)
        {
            return ExpressionReflection.AsDouble(Expression.Call(Method, Expression.Constant(Points)));
        }

        var coordinates = Points.Select(point => (Expression)Expression.Property(Expression.Constant(point), ExpressionReflection.PointCoordinates));
        return ExpressionReflection.AsDouble(Expression.Call(Method, coordinates));
    }
}

public class BoundNegation : BoundExpression
{
    public BoundNegation(BoundExpression operand)
    {
        Operand = operand;
    }

    public BoundExpression Operand { get; }

    public override double Evaluate(double parameter)
    {
        return -Operand.Evaluate(parameter);
    }

    public override Expression ToLinq(ParameterExpression parameter)
    {
        return Expression.Negate(Operand.ToLinq(parameter));
    }
}

public enum BoundOperator
{
    Add,
    Subtract,
    Multiply,
    Divide,
    Power
}

public class BoundBinary : BoundExpression
{
    public BoundBinary(BoundOperator @operator, BoundExpression left, BoundExpression right)
    {
        Operator = @operator;
        Left = left;
        Right = right;
    }

    public BoundOperator Operator { get; }

    public BoundExpression Left { get; }

    public BoundExpression Right { get; }

    public override double Evaluate(double parameter)
    {
        double left = Left.Evaluate(parameter);
        double right = Right.Evaluate(parameter);
        switch (Operator)
        {
            case BoundOperator.Add:
                return left + right;
            case BoundOperator.Subtract:
                return left - right;
            case BoundOperator.Multiply:
                return left * right;
            case BoundOperator.Divide:
                return left / right;
            default:
                return System.Math.Pow(left, right);
        }
    }

    public override Expression ToLinq(ParameterExpression parameter)
    {
        var left = Left.ToLinq(parameter);
        var right = Right.ToLinq(parameter);
        switch (Operator)
        {
            case BoundOperator.Add:
                return Expression.Add(left, right);
            case BoundOperator.Subtract:
                return Expression.Subtract(left, right);
            case BoundOperator.Multiply:
                return Expression.Multiply(left, right);
            case BoundOperator.Divide:
                return Expression.Divide(left, right);
            default:
                return Expression.Power(left, right);
        }
    }
}
