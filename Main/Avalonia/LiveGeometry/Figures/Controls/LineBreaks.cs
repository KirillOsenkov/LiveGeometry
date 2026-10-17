using System.Collections.Generic;
using System.Text;

namespace LiveGeometry;

/// <summary>
/// Where the lines of a wrapped label may break. A line breaks at any space, and the math
/// in a caption came apart in the middle of an expression: "AP = 8 −" at the end of one
/// line and "r." on the next. A space next to an operator (a word that is nothing but
/// operators: = + − × ≤ ...) is shown as a no-break space, so a short expression stays on
/// one line, and a long chain (longer than <see cref="LongChain"/> characters) may still
/// break before its second and later relations: "x² = AP² − AN²" and "= (8 − r)² − r²". For
/// what is shown only: the text keeps the spaces it was typed with.
/// </summary>
public static class LineBreaks
{
    public const char NoBreakSpace = (char)0xA0;

    /// <summary>The characters of an operator, a word of which a space next to it doesn't break</summary>
    public const string Operators = "=≠≈<>≤≥+−-±×·÷*/";

    /// <summary>The characters of a relation, before which a long chain may break</summary>
    public const string Relations = "=≠≈<>≤≥";

    /// <summary>How many characters a chain of words joined by operators may have and still never break</summary>
    public const int LongChain = 24;

    public static string KeepExpressionsTogether(string text)
    {
        if (string.IsNullOrEmpty(text) || text.IndexOf(' ') < 0)
        {
            return text;
        }

        var words = FindWords(text);
        if (words.Count < 2)
        {
            return text;
        }

        StringBuilder sb = null;
        int chainStart = 0;
        while (chainStart < words.Count)
        {
            int chainEnd = chainStart;
            while (chainEnd + 1 < words.Count && Joins(text, words[chainEnd], words[chainEnd + 1]))
            {
                chainEnd++;
            }

            bool isLong = words[chainEnd].End - words[chainStart].Start > LongChain;
            bool hadRelation = IsRelation(text, words[chainStart]);
            for (int i = chainStart + 1; i <= chainEnd; i++)
            {
                bool isRelation = IsRelation(text, words[i]);
                if (!(isLong && isRelation && hadRelation))
                {
                    sb ??= new StringBuilder(text);
                    for (int j = words[i - 1].End; j < words[i].Start; j++)
                    {
                        sb[j] = NoBreakSpace;
                    }
                }

                hadRelation |= isRelation;
            }

            chainStart = chainEnd + 1;
        }

        return sb?.ToString() ?? text;
    }

    /// <summary>The runs of characters between spaces and line breaks</summary>
    static List<(int Start, int End)> FindWords(string text)
    {
        var words = new List<(int Start, int End)>();
        int start = -1;
        for (int i = 0; i <= text.Length; i++)
        {
            bool isSeparator = i == text.Length || char.IsWhiteSpace(text[i]) && text[i] != NoBreakSpace;
            if (!isSeparator && start < 0)
            {
                start = i;
            }
            else if (isSeparator && start >= 0)
            {
                words.Add((start, i));
                start = -1;
            }
        }

        return words;
    }

    /// <summary>Whether the spaces between two words don't break: only spaces between them, and one of them an operator</summary>
    static bool Joins(string text, (int Start, int End) left, (int Start, int End) right)
    {
        for (int i = left.End; i < right.Start; i++)
        {
            if (text[i] != ' ')
            {
                return false;
            }
        }

        return IsMadeOf(text, left, Operators) || IsMadeOf(text, right, Operators);
    }

    static bool IsRelation(string text, (int Start, int End) word)
    {
        if (!IsMadeOf(text, word, Operators))
        {
            return false;
        }

        for (int i = word.Start; i < word.End; i++)
        {
            if (Relations.IndexOf(text[i]) >= 0)
            {
                return true;
            }
        }

        return false;
    }

    static bool IsMadeOf(string text, (int Start, int End) word, string characters)
    {
        for (int i = word.Start; i < word.End; i++)
        {
            if (characters.IndexOf(text[i]) < 0)
            {
                return false;
            }
        }

        return true;
    }
}
