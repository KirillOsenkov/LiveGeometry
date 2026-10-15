using System;
using System.Collections.Generic;
using System.Linq;

namespace LiveGeometry;

/// <summary>
/// What a term of a <see cref="LinearForm"/> multiplies: a number a figure gives - a point's X
/// or Y, a Number's value, a segment's length - or the distance between two points
/// </summary>
public readonly record struct FormVariable(object Owner, string Member, object Other = null);

/// <summary>
/// An expression worked out exactly as a constant plus a sum of terms, a coefficient times a
/// variable (<see cref="FormVariable"/>), the coefficients exact (<see cref="ExactNumber"/>).
/// Two expressions with equal forms are equal however the figures move: the same sum of the
/// same variables, whatever else ties the variables together - which only means that some
/// equal expressions have unequal forms, never the other way round.
/// </summary>
public class LinearForm : IEquatable<LinearForm>
{
    readonly Dictionary<FormVariable, ExactNumber> terms;

    LinearForm(Dictionary<FormVariable, ExactNumber> terms, ExactNumber constant)
    {
        this.terms = terms;
        Constant = constant;
    }

    public ExactNumber Constant { get; }

    public bool IsConstant => terms.Count == 0;

    public static LinearForm Of(ExactNumber constant)
    {
        return new LinearForm(new Dictionary<FormVariable, ExactNumber>(), constant);
    }

    public static LinearForm Of(FormVariable variable)
    {
        return new LinearForm(new Dictionary<FormVariable, ExactNumber>() { [variable] = ExactNumber.One }, default);
    }

    /// <summary>This plus the other times <paramref name="factor"/> (1 to add, -1 to subtract)</summary>
    public LinearForm Plus(LinearForm other, ExactNumber factor)
    {
        var sum = new Dictionary<FormVariable, ExactNumber>(terms);
        foreach (var (variable, coefficient) in other.terms)
        {
            var total = (sum.TryGetValue(variable, out var existing) ? existing : default) + coefficient * factor;
            if (total.IsZero)
            {
                sum.Remove(variable);
            }
            else
            {
                sum[variable] = total;
            }
        }

        return new LinearForm(sum, Constant + other.Constant * factor);
    }

    public LinearForm Times(ExactNumber factor)
    {
        if (factor.IsZero)
        {
            return Of(default(ExactNumber));
        }

        return new LinearForm(terms.ToDictionary(term => term.Key, term => term.Value * factor), Constant * factor);
    }

    public bool Equals(LinearForm other)
    {
        return other != null
            && Constant.Equals(other.Constant)
            && terms.Count == other.terms.Count
            && terms.All(term => other.terms.TryGetValue(term.Key, out var coefficient) && coefficient.Equals(term.Value));
    }

    public override bool Equals(object obj)
    {
        return Equals(obj as LinearForm);
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(Constant, terms.Count);
    }
}

/// <summary>
/// Works out the exact forms of expressions (<see cref="LinearForm"/>), through the points by
/// coordinates they name: their X and Y are not variables but forms of what they name in
/// turn, down to points of other kinds, so that a corner of a lattice built on corners built
/// on corners comes out in the two points the lattice grew from. Each point once (worked out
/// again for every corner that names it, a lattice took time exponential in its size). An
/// expression has no form - null - where it isn't such a sum: a product of two variables, a
/// function of one, a constant that isn't exact (sqrt(5)), the x of a function.
/// </summary>
public class LinearForms
{
    readonly Dictionary<(PointByCoordinates Point, bool IsX), LinearForm> coordinates = new Dictionary<(PointByCoordinates Point, bool IsX), LinearForm>();
    readonly HashSet<(PointByCoordinates Point, bool IsX)> inProgress = new HashSet<(PointByCoordinates Point, bool IsX)>();

    /// <summary>
    /// Whether the two points are the same point however the figures move: their X and Y have
    /// equal forms (<see cref="OfCoordinate"/>). A point of any other kind is its own X and Y,
    /// so a point by coordinates is that point only where its form comes to exactly them.
    /// </summary>
    public bool AreSamePoint(IPoint first, IPoint second)
    {
        var firstX = OfCoordinate(first, isX: true);
        var secondX = OfCoordinate(second, isX: true);
        if (firstX == null || !firstX.Equals(secondX))
        {
            return false;
        }

        var firstY = OfCoordinate(first, isX: false);
        return firstY != null && firstY.Equals(OfCoordinate(second, isX: false));
    }

    /// <summary>The form of a point's X or Y: worked out for a point by coordinates, the coordinate itself, a variable, for any other</summary>
    public LinearForm OfCoordinate(IPoint point, bool isX)
    {
        return point is PointByCoordinates byCoordinates
            ? Of(byCoordinates, isX)
            : LinearForm.Of(new FormVariable(point, isX ? "X" : "Y"));
    }

    /// <summary>The form of the point's X or Y, as its expression is compiled now</summary>
    public LinearForm Of(PointByCoordinates point, bool isX)
    {
        var key = (point, isX);
        if (coordinates.TryGetValue(key, out var form))
        {
            return form;
        }

        // (expressions can't name each other round in a circle; should they, there is no form)
        if (!inProgress.Add(key))
        {
            return null;
        }

        var bound = (isX ? point.XExpression : point.YExpression)?.Bound;
        form = bound == null ? null : Of(bound);
        inProgress.Remove(key);
        coordinates[key] = form;
        return form;
    }

    public LinearForm Of(BoundExpression expression)
    {
        switch (expression)
        {
            case BoundConstant constant:
                return Fraction.FromDouble(constant.Value) is Fraction value ? LinearForm.Of(new ExactNumber(value)) : null;
            case BoundNegation negation:
                return Of(negation.Operand)?.Times(-ExactNumber.One);
            case BoundBinary binary:
                return Of(binary);
            case BoundCall call when call.Method.Name == nameof(System.Math.Sqrt) && call.Arguments.Count == 1:
                return Of(call.Arguments[0]) is { IsConstant: true } argument && ExactNumber.Sqrt(argument.Constant) is ExactNumber root
                    ? LinearForm.Of(root)
                    : null;
            case BoundProperty property:
                if (property.Figure is PointByCoordinates point && property.Property.Name is "X" or "Y")
                {
                    return Of(point, isX: property.Property.Name == "X");
                }

                return LinearForm.Of(new FormVariable(property.Figure, property.Property.Name));
            case BoundNumber number:
                return LinearForm.Of(new FormVariable(number.Number, nameof(INumber.Value)));
            case BoundDistance distance:
                return LinearForm.Of(new FormVariable(distance.First, "Distance", distance.Second));
            default:
                return null;
        }
    }

    LinearForm Of(BoundBinary binary)
    {
        var left = Of(binary.Left);
        var right = left == null ? null : Of(binary.Right);
        if (right == null)
        {
            return null;
        }

        switch (binary.Operator)
        {
            case BoundOperator.Add:
                return left.Plus(right, ExactNumber.One);
            case BoundOperator.Subtract:
                return left.Plus(right, -ExactNumber.One);
            case BoundOperator.Multiply:
                if (left.IsConstant)
                {
                    return right.Times(left.Constant);
                }

                return right.IsConstant ? left.Times(right.Constant) : null;
            case BoundOperator.Divide:
                return right.IsConstant && right.Constant.IsRational && !right.Constant.IsZero
                    ? left.Times(new ExactNumber(new Fraction(right.Constant.Rational.Denominator, right.Constant.Rational.Numerator)))
                    : null;
            default:
                return null;
        }
    }
}
