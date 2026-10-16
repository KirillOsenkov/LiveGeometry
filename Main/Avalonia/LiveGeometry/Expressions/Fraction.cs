using System;
using System.Globalization;
using System.Numerics;

namespace LiveGeometry;

/// <summary>
/// An exact fraction of whole numbers, what <see cref="ExactNumber"/> calculates with: in
/// lowest terms, the denominator positive. The default is 0.
/// </summary>
public readonly struct Fraction : IEquatable<Fraction>
{
    readonly BigInteger denominator;

    /// <param name="denominator">Not 0: whoever divides checks first</param>
    public Fraction(BigInteger numerator, BigInteger denominator)
    {
        if (denominator.Sign < 0)
        {
            numerator = -numerator;
            denominator = -denominator;
        }

        var divisor = BigInteger.GreatestCommonDivisor(numerator, denominator);
        if (!divisor.IsZero && !divisor.IsOne)
        {
            numerator /= divisor;
            denominator /= divisor;
        }

        Numerator = numerator;
        this.denominator = denominator;
    }

    public Fraction(BigInteger whole)
        : this(whole, BigInteger.One)
    {
    }

    public BigInteger Numerator { get; }

    public BigInteger Denominator => denominator.IsZero ? BigInteger.One : denominator;

    public bool IsZero => Numerator.IsZero;

    public int Sign => Numerator.Sign;

    /// <summary>
    /// The number as it reads, the shortest decimals that give the double back: 0.1 is a tenth
    /// (the double is a hair off it), as it was typed in an expression. Null for no number.
    /// </summary>
    public static Fraction? FromDouble(double value)
    {
        if (!double.IsFinite(value))
        {
            return null;
        }

        // [-]digits[.digits][E[+-]digits]
        var text = value.ToString("R", CultureInfo.InvariantCulture);
        int exponent = 0;
        int exponentStart = text.IndexOfAny(['E', 'e']);
        if (exponentStart >= 0)
        {
            exponent = int.Parse(text.Substring(exponentStart + 1), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
            text = text.Substring(0, exponentStart);
        }

        bool negative = text.StartsWith("-", StringComparison.Ordinal);
        if (negative)
        {
            text = text.Substring(1);
        }

        int point = text.IndexOf('.');
        if (point >= 0)
        {
            exponent -= text.Length - point - 1;
            text = text.Remove(point, 1);
        }

        var numerator = BigInteger.Parse(text, NumberStyles.None, CultureInfo.InvariantCulture);
        var denominator = BigInteger.One;
        if (exponent >= 0)
        {
            numerator *= BigInteger.Pow(10, exponent);
        }
        else
        {
            denominator = BigInteger.Pow(10, -exponent);
        }

        return new Fraction(negative ? -numerator : numerator, denominator);
    }

    public static Fraction operator +(Fraction a, Fraction b)
    {
        return new Fraction(a.Numerator * b.Denominator + b.Numerator * a.Denominator, a.Denominator * b.Denominator);
    }

    public static Fraction operator -(Fraction a, Fraction b)
    {
        return new Fraction(a.Numerator * b.Denominator - b.Numerator * a.Denominator, a.Denominator * b.Denominator);
    }

    public static Fraction operator -(Fraction a)
    {
        return new Fraction(-a.Numerator, a.Denominator);
    }

    public static Fraction operator *(Fraction a, Fraction b)
    {
        return new Fraction(a.Numerator * b.Numerator, a.Denominator * b.Denominator);
    }

    /// <param name="b">Not 0: whoever divides checks first</param>
    public static Fraction operator /(Fraction a, Fraction b)
    {
        return new Fraction(a.Numerator * b.Denominator, a.Denominator * b.Numerator);
    }

    public bool Equals(Fraction other)
    {
        return Numerator == other.Numerator && Denominator == other.Denominator;
    }

    public override bool Equals(object obj)
    {
        return obj is Fraction other && Equals(other);
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(Numerator, Denominator);
    }

    public static bool operator ==(Fraction a, Fraction b)
    {
        return a.Equals(b);
    }

    public static bool operator !=(Fraction a, Fraction b)
    {
        return !a.Equals(b);
    }

    public override string ToString()
    {
        return Denominator.IsOne ? Numerator.ToString(CultureInfo.InvariantCulture) : Numerator.ToString(CultureInfo.InvariantCulture) + "/" + Denominator.ToString(CultureInfo.InvariantCulture);
    }
}
