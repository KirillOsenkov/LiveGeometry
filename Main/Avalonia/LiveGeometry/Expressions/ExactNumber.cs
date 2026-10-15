using System;
using System.Numerics;

namespace LiveGeometry;

/// <summary>
/// A number worked out exactly: a + b sqrt(2) + c sqrt(3) + d sqrt(6), a to d fractions
/// (<see cref="Fraction"/>) - what the coordinates of the regular polygons that tile need
/// (<see cref="RegularPolygonBuilder"/>), and closed under adding and multiplying. The
/// default is 0.
/// </summary>
public readonly struct ExactNumber : IEquatable<ExactNumber>
{
    public ExactNumber(Fraction rational, Fraction root2 = default, Fraction root3 = default, Fraction root6 = default)
    {
        Rational = rational;
        Root2 = root2;
        Root3 = root3;
        Root6 = root6;
    }

    public static readonly ExactNumber One = new ExactNumber(new Fraction(BigInteger.One));

    public Fraction Rational { get; }

    /// <summary>What multiplies sqrt(2)</summary>
    public Fraction Root2 { get; }

    /// <summary>What multiplies sqrt(3)</summary>
    public Fraction Root3 { get; }

    /// <summary>What multiplies sqrt(6)</summary>
    public Fraction Root6 { get; }

    public bool IsRational => Root2.IsZero && Root3.IsZero && Root6.IsZero;

    public bool IsZero => Rational.IsZero && IsRational;

    public static ExactNumber operator +(ExactNumber a, ExactNumber b)
    {
        return new ExactNumber(a.Rational + b.Rational, a.Root2 + b.Root2, a.Root3 + b.Root3, a.Root6 + b.Root6);
    }

    public static ExactNumber operator -(ExactNumber a, ExactNumber b)
    {
        return new ExactNumber(a.Rational - b.Rational, a.Root2 - b.Root2, a.Root3 - b.Root3, a.Root6 - b.Root6);
    }

    public static ExactNumber operator -(ExactNumber a)
    {
        return new ExactNumber(-a.Rational, -a.Root2, -a.Root3, -a.Root6);
    }

    /// <summary>sqrt(2) sqrt(2) = 2, sqrt(2) sqrt(3) = sqrt(6), sqrt(2) sqrt(6) = 2 sqrt(3), sqrt(3) sqrt(6) = 3 sqrt(2)...</summary>
    public static ExactNumber operator *(ExactNumber a, ExactNumber b)
    {
        var two = new Fraction(2);
        var three = new Fraction(3);
        var six = new Fraction(6);
        return new ExactNumber(
            a.Rational * b.Rational + two * a.Root2 * b.Root2 + three * a.Root3 * b.Root3 + six * a.Root6 * b.Root6,
            a.Rational * b.Root2 + a.Root2 * b.Rational + three * (a.Root3 * b.Root6 + a.Root6 * b.Root3),
            a.Rational * b.Root3 + a.Root3 * b.Rational + two * (a.Root2 * b.Root6 + a.Root6 * b.Root2),
            a.Rational * b.Root6 + a.Root6 * b.Rational + a.Root2 * b.Root3 + a.Root3 * b.Root2);
    }

    /// <summary>
    /// The square root of a fraction that is r squared times 1, 2, 3 or 6; null for any other
    /// number (sqrt(5), a negative one, one with a root of its own)
    /// </summary>
    public static ExactNumber? Sqrt(ExactNumber value)
    {
        if (!value.IsRational || value.Rational.Sign < 0)
        {
            return null;
        }

        // sqrt(p / q) = sqrt(p q) / q, and p q = square^2 rest; a number of hundreds of digits
        // (a tiny decimal typed) is no root of anything here
        var product = value.Rational.Numerator * value.Rational.Denominator;
        if (product.GetBitLength() > 1000)
        {
            return null;
        }

        var square = BigInteger.One;
        var rest = product;
        for (int factor = 2; factor <= 1000 && !rest.IsZero; factor++)
        {
            var factorSquared = new BigInteger(factor * factor);
            if (factorSquared > rest)
            {
                break;
            }

            while ((rest % factorSquared).IsZero)
            {
                rest /= factorSquared;
                square *= factor;
            }
        }

        var root = IntegerSquareRoot(rest);
        if (root * root == rest)
        {
            square *= root;
            rest = BigInteger.One;
        }

        var coefficient = new Fraction(rest.IsZero ? BigInteger.Zero : square, value.Rational.Denominator);
        if (rest.IsZero || rest.IsOne)
        {
            return new ExactNumber(coefficient);
        }

        if (rest == 2)
        {
            return new ExactNumber(default, root2: coefficient);
        }

        if (rest == 3)
        {
            return new ExactNumber(default, root3: coefficient);
        }

        if (rest == 6)
        {
            return new ExactNumber(default, root6: coefficient);
        }

        return null;
    }

    /// <summary>The largest whole number whose square is no more than the value (Newton's method)</summary>
    static BigInteger IntegerSquareRoot(BigInteger value)
    {
        if (value < 2)
        {
            return value;
        }

        var root = value;
        var next = (root + 1) / 2;
        while (next < root)
        {
            root = next;
            next = (root + value / root) / 2;
        }

        return root;
    }

    public bool Equals(ExactNumber other)
    {
        return Rational == other.Rational && Root2 == other.Root2 && Root3 == other.Root3 && Root6 == other.Root6;
    }

    public override bool Equals(object obj)
    {
        return obj is ExactNumber other && Equals(other);
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(Rational, Root2, Root3, Root6);
    }

    public override string ToString()
    {
        return Rational + " + " + Root2 + " sqrt(2) + " + Root3 + " sqrt(3) + " + Root6 + " sqrt(6)";
    }
}
