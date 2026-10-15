using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace DynamicGeometry;

/// <param name="Text">The emoji as it is written, with its FE0F if it has one</param>
/// <param name="Name">Its CLDR name: "cherries", "red heart"</param>
/// <param name="Subgroup">Its kind: "food-fruit", "heart"</param>
public record Emoji(string Text, string Name, string Subgroup);

/// <summary>
/// The emoji a point style can pick by name: the single-character ones the emoji font has
/// (Emoji.txt, made by tools/emoji.cs), in CLDR order, which keeps kinds together.
/// </summary>
public static class EmojiList
{
    static List<Emoji> all;

    public static IReadOnlyList<Emoji> All
    {
        get
        {
            if (all == null)
            {
                all = Load();
            }

            return all;
        }
    }

    static List<Emoji> Load()
    {
        var result = new List<Emoji>();
        using var stream = typeof(EmojiList).Assembly.GetManifestResourceStream("DynamicGeometry.Emoji.Emoji.txt");
        using var reader = new StreamReader(stream);
        string line;
        while ((line = reader.ReadLine()) != null)
        {
            var parts = line.Split('\t');
            if (parts.Length == 3)
            {
                result.Add(new Emoji(parts[0], parts[1], parts[2]));
            }
        }

        return result;
    }

    /// <summary>What a point would want first: shapes in colors, stars, pins, a few things</summary>
    static readonly string[] suggestions =
    {
        "🔴", "🟠", "🟡", "🟢", "🔵", "🟣", "🟤", "⚫", "⚪",
        "🟥", "🟧", "🟨", "🟩", "🟦", "🟪", "🟫", "⬛", "⬜",
        "🔺", "🔻", "🔷", "🔶", "🔹", "🔸", "💠", "🔘", "⭕",
        "⭐", "🌟", "✨", "💥", "❤️", "💎", "🎯", "❌", "✅",
        "📍", "📌", "🚩", "🏁", "🏠", "🏰", "🌳", "⛰️", "🌋",
        "🍎", "🍒", "🌸", "🌻", "🍀", "🐞", "🦋", "🐱", "🐶",
        "🐸", "🐢", "🚀", "✈️", "🚗", "⚽", "🏀", "🎈", "☀️"
    };

    public static IEnumerable<Emoji> Suggestions
    {
        get
        {
            foreach (var text in suggestions)
            {
                var emoji = Find(text);
                if (emoji != null)
                {
                    yield return emoji;
                }
            }
        }
    }

    /// <summary>The emoji that is this text, with or without its FE0F; null if none is</summary>
    public static Emoji Find(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return null;
        }

        string bare = text.Replace("️", "");
        return All.FirstOrDefault(e => e.Text == text || e.Text.Replace("️", "") == bare);
    }

    /// <summary>
    /// Every word of the query starts a word of the name ("red hea" finds red heart); the
    /// subgroup counts too, after the names ("fruit" finds the fruits, even "grapes").
    /// </summary>
    public static IEnumerable<Emoji> Search(string query)
    {
        var words = SplitWords(query);
        if (words.Length == 0)
        {
            return Enumerable.Empty<Emoji>();
        }

        var byName = All.Where(e => Matches(words, SplitWords(e.Name)));
        var bySubgroup = All.Where(e => Matches(words, SplitWords(e.Name + " " + e.Subgroup)));
        return byName.Concat(bySubgroup).Distinct();
    }

    static bool Matches(string[] queryWords, string[] words)
    {
        return queryWords.All(q => words.Any(w => w.StartsWith(q, StringComparison.OrdinalIgnoreCase)));
    }

    static string[] SplitWords(string text)
    {
        return (text ?? "").Split(new[] { ' ', '-', ':', ',', '&' }, StringSplitOptions.RemoveEmptyEntries);
    }

    /// <summary>
    /// The text is one character a user could want as it is (pasted ★ or Ω): a single text
    /// element, and not a plain letter or digit, which would just be the start of a search
    /// </summary>
    public static bool IsSingleCharacter(string text)
    {
        if (string.IsNullOrWhiteSpace(text) || new StringInfo(text).LengthInTextElements != 1)
        {
            return false;
        }

        return !(text.Length == 1 && text[0] < 128 && char.IsLetterOrDigit(text[0]));
    }

    /// <summary>"U+1F352", or "U+2764 U+FE0F"</summary>
    public static string CodePoints(string text)
    {
        var result = new List<string>();
        for (int i = 0; i < text.Length; i++)
        {
            int codePoint = text[i];
            if (char.IsSurrogatePair(text, i))
            {
                codePoint = char.ConvertToUtf32(text, i);
                i++;
            }

            result.Add("U+" + codePoint.ToString("X4", CultureInfo.InvariantCulture));
        }

        return string.Join(" ", result);
    }
}
