using System.Text;

namespace DynamicGeometry;

/// <summary>
/// How a name is drawn: A_1 as A₁, the way GeoGebra (and TeX) read an underscore, and A1 as
/// A₁ too - the trailing digits of a name are its index, which is how a textbook writes it.
/// Only the text on the screen changes - the name itself, in files, in expressions and in
/// the Name box, stays as typed. What follows an underscore is the run of letters and
/// digits, or anything in braces (A_{12}); Unicode has subscript digits and a few subscript
/// letters, and a part it can't write stays as typed.
/// </summary>
public static class NameDisplay
{
    const string Plain = "0123456789aeoxhklmnpst";
    const string Subscript = "₀₁₂₃₄₅₆₇₈₉ₐₑₒₓₕₖₗₘₙₚₛₜ";

    public static string Format(string name)
    {
        if (string.IsNullOrEmpty(name))
        {
            return name;
        }

        if (name.IndexOf('_') < 0)
        {
            return FormatTrailingDigits(name);
        }

        var sb = new StringBuilder();
        int i = 0;
        while (i < name.Length)
        {
            char c = name[i];
            if (c != '_' || i == 0)
            {
                sb.Append(c);
                i++;
                continue;
            }

            int start = i + 1;
            int next;
            string part;
            if (start < name.Length && name[start] == '{')
            {
                int close = name.IndexOf('}', start);
                if (close < 0)
                {
                    sb.Append(c);
                    i++;
                    continue;
                }

                part = name.Substring(start + 1, close - start - 1);
                next = close + 1;
            }
            else
            {
                int end = start;
                while (end < name.Length && char.IsLetterOrDigit(name[end]))
                {
                    end++;
                }

                part = name.Substring(start, end - start);
                next = end;
            }

            var subscript = ToSubscript(part);
            if (subscript == null)
            {
                sb.Append(name, i, next - i);
            }
            else
            {
                sb.Append(subscript);
            }

            i = next;
        }

        return sb.ToString();
    }

    /// <summary>A1 as A₁; a name that is all digits, or ends in none, stays</summary>
    static string FormatTrailingDigits(string name)
    {
        int start = name.Length;
        while (start > 0 && char.IsDigit(name[start - 1]))
        {
            start--;
        }

        if (start == 0 || start == name.Length)
        {
            return name;
        }

        return name.Substring(0, start) + ToSubscript(name.Substring(start));
    }

    /// <summary>The text in subscript characters, or null when a character has none</summary>
    static string ToSubscript(string text)
    {
        if (text.Length == 0)
        {
            return null;
        }

        var sb = new StringBuilder(text.Length);
        foreach (var c in text)
        {
            int index = Plain.IndexOf(c);
            if (index < 0)
            {
                return null;
            }

            sb.Append(Subscript[index]);
        }

        return sb.ToString();
    }
}
