using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LiveGeometry;

/// <summary>
/// Expressions written for the language as it was until 2026-10-18 - VB6's names, which DG
/// 1.0's files (.dgf) have, and the ones this language kept from them: dist, ang, oang and
/// xang in radians, log for e's logarithm, lg, sqr for the square root, int, sgn, ceiling -
/// in the language as it is (<see cref="FunctionCatalog"/>), with the same values: ang(A, B, C)
/// is rad(angle(A, B, C)), and deg(ang(A, B, C)) is angle(A, B, C). What the .dgf reader
/// passes every expression through; the gallery's drawings were rewritten with it once.
/// Anything else is left as it is, and a text that doesn't scan comes back unchanged.
/// </summary>
public static class LegacyExpressions
{
    /// <param name="fromDG">VB6's angle(A, B, C) is ang's twin, in radians; this language's was in degrees</param>
    public static string Translate(string text, bool fromDG)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text;
        }

        var scan = Scanner.Scan(text);
        if (scan.Errors.Count > 0)
        {
            return text;
        }

        var tokens = scan.Tokens;
        var sb = new StringBuilder();
        int copied = 0;
        for (int i = 0; i + 1 < tokens.Count; i++)
        {
            var token = tokens[i];
            if (token.Kind != TokenType.Identifier || token.IsQuoted || tokens[i + 1].Kind != TokenType.OpenParen)
            {
                continue;
            }

            int close = ClosingParen(tokens, i + 1);
            if (close < 0)
            {
                return text;
            }

            var arguments = Arguments(text, tokens, i + 1, close).Select(argument => Translate(argument, fromDG)).ToList();
            var replacement = Rewrite(token.Text.ToLowerInvariant(), arguments, fromDG);
            if (replacement == null)
            {
                continue;
            }

            sb.Append(text, copied, token.Start - copied);
            sb.Append(replacement);
            copied = tokens[close].Start + tokens[close].Length;
            i = close;
        }

        sb.Append(text, copied, text.Length - copied);
        return sb.ToString();
    }

    /// <summary>The text of a label: each [...] part translated</summary>
    public static string TranslateLabelText(string text, bool fromDG)
    {
        if (string.IsNullOrEmpty(text))
        {
            return text;
        }

        var sb = new StringBuilder(text);
        foreach (var part in LabelBase.FindExpressionParts(text).OrderByDescending(part => part.Start))
        {
            string inner = text.Substring(part.Start + 1, part.Length - 2);
            sb.Remove(part.Start + 1, part.Length - 2);
            sb.Insert(part.Start + 1, Translate(inner, fromDG));
        }

        return sb.ToString();
    }

    /// <summary>The call in the language as it is, or null for one that stays as it was</summary>
    static string Rewrite(string name, List<string> arguments, bool fromDG)
    {
        string all = string.Join(", ", arguments);
        string first = arguments.Count > 0 ? arguments[0] : "";
        switch (name)
        {
            case "dist":
            case "distance":
                return "distance(" + all + ")";
            case "ang":
                return "rad(angle(" + all + "))";
            case "angle" when fromDG:
                return "rad(angle(" + all + "))";
            case "oang":
            case "oangle":
                return "rad(directedangle(" + all + "))";
            case "xang":
            case "xangle":
                return "rad(direction(" + all + "))";
            case "arg" when arguments.Count == 1:
                return "atan2(" + first + ".Y, " + first + ".X)";
            case "norm" when arguments.Count == 1:
                return "sqrt(" + first + ".X^2 + " + first + ".Y^2)";
            case "log" when arguments.Count == 1:
                return "ln(" + first + ")";
            case "log" when arguments.Count == 2:
                return "(ln(" + first + ") / ln(" + arguments[1] + "))";
            case "lg":
            case "log10":
                return "log(" + all + ")";
            case "log2":
                return "log_2(" + all + ")";
            case "sqr":
                return "sqrt(" + all + ")";
            case "int":
                return "floor(" + all + ")";
            case "sgn":
                return "sign(" + all + ")";
            case "ceiling":
                return "ceil(" + all + ")";
            case "truncate" when arguments.Count == 1:
                return "(sign(" + first + ") * floor(abs(" + first + ")))";
            case "pow" when arguments.Count == 2:
                return "(" + first + ")^(" + arguments[1] + ")";
            case "deg":
            case "todeg":
                // deg(ang(A, B, C)) is the angle in degrees, as the language says it now
                return arguments.Count == 1 && Unwrap(first, "rad") is string degrees ? degrees : "deg(" + all + ")";
            case "torad":
                return "rad(" + all + ")";
            case "asin":
            case "acos":
            case "atan":
            case "asinh":
            case "acosh":
            case "atanh":
                return "arc" + name.Substring(1) + "(" + all + ")";
            case "tg":
                return "tan(" + all + ")";
            case "ctg":
                return "cot(" + all + ")";
            case "atn":
            case "arctg":
                return "arctan(" + all + ")";
            default:
                return null;
        }
    }

    /// <summary>What f(...) says inside when the text is that call whole, else null</summary>
    static string Unwrap(string text, string function)
    {
        if (!text.StartsWith(function + "(", StringComparison.Ordinal) || !text.EndsWith(')'))
        {
            return null;
        }

        var tokens = Scanner.Scan(text).Tokens;
        return tokens.Count >= 3 && ClosingParen(tokens, 1) == tokens.Count - 1
            ? text.Substring(function.Length + 1, text.Length - function.Length - 2)
            : null;
    }

    /// <summary>The index of the ) that closes the ( at the index; -1 when none does</summary>
    static int ClosingParen(IReadOnlyList<Token> tokens, int open)
    {
        int depth = 0;
        for (int i = open; i < tokens.Count; i++)
        {
            if (tokens[i].Kind == TokenType.OpenParen)
            {
                depth++;
            }
            else if (tokens[i].Kind == TokenType.CloseParen && --depth == 0)
            {
                return i;
            }
        }

        return -1;
    }

    /// <summary>The texts between the parentheses, split at the commas outside inner ones, each trimmed</summary>
    static IEnumerable<string> Arguments(string text, IReadOnlyList<Token> tokens, int open, int close)
    {
        int depth = 0;
        int start = tokens[open].Start + 1;
        for (int i = open + 1; i < close; i++)
        {
            switch (tokens[i].Kind)
            {
                case TokenType.OpenParen:
                    depth++;
                    break;
                case TokenType.CloseParen:
                    depth--;
                    break;
                case TokenType.Comma when depth == 0:
                    yield return text.Substring(start, tokens[i].Start - start).Trim();
                    start = tokens[i].Start + 1;
                    break;
            }
        }

        var last = text.Substring(start, tokens[close].Start - start).Trim();
        if (last.Length > 0 || start > tokens[open].Start + 1)
        {
            yield return last;
        }
    }
}
