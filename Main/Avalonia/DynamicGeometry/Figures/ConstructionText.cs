using System;
using System.Globalization;
using System.Linq;
using System.Text;

namespace DynamicGeometry;

/// <summary>
/// The words of a figure's <see cref="FigureBase.Construction"/>: what it says of the figures
/// it is built on. A point goes by its name (E), anything else by its plain noun and name
/// (segment CD, circle k); a value typed as a number by the number (45°), one taken from a
/// figure by what says it (a, AB, angle ABC).
/// </summary>
public static class ConstructionText
{
    /// <summary>"circle k", "E", "side 2 of regular pentagon p"</summary>
    public static string Of(IFigure figure)
    {
        if (figure == null)
        {
            return "?";
        }

        // a part has no name: it says which part of what it is
        if (figure is IFigurePart)
        {
            return LowerFirst(figure.ToString());
        }

        if (figure is FigureBase figureBase)
        {
            return figureBase.Reference;
        }

        return NameDisplay.Format(figure.Name);
    }

    /// <summary>The names of points run together: CD, ABC</summary>
    public static string Points(params IFigure[] points)
    {
        return string.Concat(points.Select(point => point == null ? "?" : NameDisplay.Format(point.Name)));
    }

    /// <summary>"angle ABC" for the angle at <paramref name="vertex"/> (no ∠: the browser's only font has none)</summary>
    public static string Angle(IFigure vertex, IFigure side1, IFigure side2)
    {
        return "angle " + Points(side1, vertex, side2);
    }

    /// <summary>A number as the property grid shows it: 3, 2.5, 1.41</summary>
    public static string Number(double value)
    {
        return Math.Round(value, Settings.DisplayDecimals).ToStringInvariant();
    }

    /// <summary>A length or a factor: 3 (typed), a (a slider), AB (a segment), "AB" (a distance measured)</summary>
    public static string Length(IFigure source)
    {
        return Value(source, Number, number => NameDisplay.Format(number.Name));
    }

    /// <summary>An angle: 45° (typed), angle a (a slider), angle ABC (a measurement)</summary>
    public static string AngleValue(IFigure source)
    {
        return Value(source, value => Number(value) + "°", number => "angle " + NameDisplay.Format(number.Name));
    }

    static string Value(IFigure source, Func<double, string> typed, Func<IFigure, string> named)
    {
        switch (source)
        {
            case DynamicGeometry.Number number when number.Auxiliary:
                return typed(number.Value);
            case INumber:
                return named(source);
            case AngleMeasurementBase:
            case AngleArc:
                return "angle " + source.Construction;
            case DistanceMeasurement:
                return source.Construction;
            case PerimeterMeasurement:
                return "perimeter " + source.Construction;
            case Segment:
            case Vector:
                // a segment's name is its points: AB, not "segment AB" (radius AB)
                return NameDisplay.Format(source.Name);
            default:
                return Of(source);
        }
    }

    /// <summary>
    /// The terms of an equation, each a coefficient's text and its variable ("" for the
    /// constant), as a school book writes them: 2x - y + 3, a·x, (a + 1)x; terms of 0 left out
    /// </summary>
    public static string Sum(params (string Coefficient, string Variable)[] terms)
    {
        var sb = new StringBuilder();
        foreach (var (coefficient, variable) in terms)
        {
            var term = Term((coefficient ?? "").Trim(), variable);
            if (term == null)
            {
                continue;
            }

            if (sb.Length == 0)
            {
                sb.Append(term);
            }
            else if (term.StartsWith('-'))
            {
                sb.Append(" - ").Append(term.Substring(1).TrimStart());
            }
            else
            {
                sb.Append(" + ").Append(term);
            }
        }

        return sb.Length == 0 ? "0" : sb.ToString();
    }

    static string Term(string coefficient, string variable)
    {
        if (coefficient.Length == 0 || coefficient == "0")
        {
            return null;
        }

        if (variable.Length == 0)
        {
            return coefficient;
        }

        if (coefficient == "1")
        {
            return variable;
        }

        if (coefficient == "-1")
        {
            return "-" + variable;
        }

        if (IsNumber(coefficient))
        {
            return coefficient + variable;
        }

        // a·x: ax would read as one name
        if (IsName(coefficient.TrimStart('-')))
        {
            return coefficient + "·" + variable;
        }

        return "(" + coefficient + ")" + variable;
    }

    static bool IsNumber(string text)
    {
        return double.TryParse(text, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out _);
    }

    static bool IsName(string text)
    {
        return text.Length > 0 && char.IsLetter(text[0]) && text.All(c => char.IsLetterOrDigit(c) || c == '_' || c == '.' || c == '\'');
    }

    /// <summary>The start of a text, in quotes: “Drag the point…”</summary>
    public static string Quote(string text)
    {
        const int maxLength = 30;
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        var line = text.Trim();
        int lineEnd = line.IndexOfAny(new[] { '\r', '\n' });
        bool cut = false;
        if (lineEnd >= 0)
        {
            line = line.Substring(0, lineEnd).TrimEnd();
            cut = true;
        }

        if (line.Length > maxLength)
        {
            line = line.Substring(0, maxLength).TrimEnd();
            cut = true;
        }

        return "“" + line + (cut ? "…" : "") + "”";
    }

    static string LowerFirst(string text)
    {
        return string.IsNullOrEmpty(text) ? text : char.ToLowerInvariant(text[0]) + text.Substring(1);
    }
}
