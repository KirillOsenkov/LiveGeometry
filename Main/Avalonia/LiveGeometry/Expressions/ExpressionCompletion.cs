using System;
using System.Collections.Generic;
using System.Linq;

namespace LiveGeometry;

/// <summary>What an item of the completion list is: the color and the icon of its row</summary>
public enum CompletionItemKind
{
    Figure,
    Member,
    Function,
    Reserved
}

/// <summary>A row of the completion list of an expression box (<see cref="ExpressionCompletion"/>)</summary>
public class CompletionItem
{
    readonly Func<string> help;

    public CompletionItem(
        string text,
        string insertText,
        CompletionItemKind kind,
        string kindText,
        IFigure figure,
        string heading,
        Func<string> help,
        IReadOnlyList<string> otherNames = null)
    {
        Text = text;
        InsertText = insertText;
        Kind = kind;
        KindText = kindText;
        Figure = figure;
        Heading = heading;
        this.help = help;
        OtherNames = otherNames ?? [];
    }

    /// <summary>What the row shows, and what the word typed is matched against: "A", "sin", "Radius", "my point"</summary>
    public string Text { get; }

    /// <summary>Other names the word typed is matched against, which a commit replaces by the text: asin for arcsin, ceiling for ceil</summary>
    public IReadOnlyList<string> OtherNames { get; }

    /// <summary>Whether a name of the row, its own or another, passes the test</summary>
    public bool AnyName(Func<string, bool> test)
    {
        return test(Text) || OtherNames.Any(test);
    }

    /// <summary>What a commit puts in place of the word: the text, in backticks where a name needs them</summary>
    public string InsertText { get; }

    public CompletionItemKind Kind { get; }

    /// <summary>A word for what it is, faint at the row's right ("point", "slider", "function"); null for none</summary>
    public string KindText { get; }

    /// <summary>The figure it is, or whose quantity it is, for the row's icon; null for a function or a constant</summary>
    public IFigure Figure { get; }

    /// <summary>The first line of the help under the list: "sin(x)", "Circle c", "c.Radius"</summary>
    public string Heading { get; }

    /// <summary>The rest of the help, worked out when the item is selected (a value as it is now); null for none</summary>
    public string Help => help?.Invoke();
}

/// <summary>
/// The word at the caret as completion sees it: what a commit replaces, from
/// <see cref="Start"/> to <see cref="End"/>, the caret anywhere in it or at either end; and
/// whether it is a quantity of a figure, after a dot (<see cref="MemberOf"/>)
/// </summary>
public class CompletionSpot
{
    public CompletionSpot(string text, int start, int end, Token memberOf)
    {
        Word = text.Substring(start, end - start);
        Start = start;
        End = end;
        MemberOf = memberOf;
    }

    public int Start { get; }

    public int End { get; }

    public string Word { get; }

    /// <summary>The name before the dot when the word is a figure's quantity (a token: it says when the name is in backticks); null for a name</summary>
    public Token MemberOf { get; }

    /// <summary>Whether the other spot is the same word still, typed on or deleted back: the same start, the same kind of list</summary>
    public bool IsSameWord(CompletionSpot other)
    {
        return other != null && other.Start == Start && (other.MemberOf == null) == (MemberOf == null);
    }
}

/// <summary>The rows the list shows for a word, which of them is selected, and whether firmly (<see cref="ExpressionCompletion.Filter"/>)</summary>
public record CompletionFilter(IReadOnlyList<CompletionItem> Shown, int Selected, bool Hard);

/// <summary>
/// The completion list of the boxes an expression is typed into, without its UI
/// (docs/completion-plan.md; the UI is <see cref="CompletionList"/>): where the word at the
/// caret is, what can go there - the names of figures, functions and constants, or the
/// quantities of the figure before a dot - and which of them the word typed so far means.
/// </summary>
public static class ExpressionCompletion
{
    /// <summary>
    /// The word at the caret: from the caret back and on over the characters of a name, so
    /// that a commit in the middle of a word replaces all of it, the digits of a number in
    /// front of it left out (2x). Null where completion has nothing to say: in a label's
    /// text outside a [...] part, inside a name in backticks, on a number.
    /// </summary>
    public static CompletionSpot Locate(string text, int caret, InsertContext context)
    {
        text ??= "";
        caret = System.Math.Clamp(caret, 0, text.Length);
        int regionStart = 0;
        int regionEnd = text.Length;
        if (context == InsertContext.Text)
        {
            int open = OpenBracketBefore(text, caret);
            if (open < 0)
            {
                return null;
            }

            regionStart = open + 1;
            int close = text.IndexOf(']', caret);
            regionEnd = close < 0 ? text.Length : close;
        }

        // inside a name in backticks: after an odd number of them
        if (text.Skip(regionStart).Take(caret - regionStart).Count(c => c == '`') % 2 == 1)
        {
            return null;
        }

        int start = caret;
        while (start > regionStart && Scanner.IsNamePart(text[start - 1]))
        {
            start--;
        }

        int end = caret;
        while (end < regionEnd && Scanner.IsNamePart(text[end]))
        {
            end++;
        }

        // a number right before a name multiplies it (2x, 3sin(x)): the word is the name
        if (start < end && char.IsAsciiDigit(text[start]))
        {
            while (start < end && char.IsAsciiDigit(text[start]))
            {
                start++;
            }

            if (start == end || start > caret)
            {
                return null;
            }
        }

        if (start < end && !Scanner.IsNameStart(text[start]))
        {
            return null;
        }

        if (start > regionStart && text[start - 1] == '.')
        {
            var owner = NameEndingAt(text.Substring(regionStart, start - 1 - regionStart));
            return owner == null ? null : new CompletionSpot(text, start, end, owner);
        }

        return new CompletionSpot(text, start, end, memberOf: null);
    }

    /// <summary>Where the [ that the caret is after opens, in a label's text: one that no ] has closed before the caret; -1 when there is none</summary>
    public static int OpenBracketBefore(string text, int caret)
    {
        caret = System.Math.Clamp(caret, 0, text.Length);
        int open = -1;
        for (int i = 0; i < caret; i++)
        {
            if (text[i] == '[')
            {
                open = i;
            }
            else if (text[i] == ']')
            {
                open = -1;
            }
        }

        return open;
    }

    /// <summary>The name the text ends with, as the scanner reads it; null when it ends with anything else (a number, an operator)</summary>
    static Token NameEndingAt(string text)
    {
        var tokens = Scanner.Scan(text).Tokens;
        var last = tokens.LastOrDefault();
        return last != null && last.Kind == TokenType.Identifier && last.Start + last.Length == text.Length ? last : null;
    }

    /// <summary>
    /// What can go at the spot: the quantities of the figure before the dot, or the names -
    /// figures an expression can say, functions, pi, e, and x in a function graph's
    /// function - sorted by name. Nothing the box's figure is built on may be offered: a
    /// cycle. Worked out once, when the list opens.
    /// </summary>
    public static IReadOnlyList<CompletionItem> Items(CompletionSpot spot, Drawing drawing, IFigure owner, InsertContext context)
    {
        if (spot == null || drawing == null)
        {
            return [];
        }

        Predicate<IFigure> isAllowed = figure => owner == null || (figure != owner && !figure.DependsOn(owner));
        return spot.MemberOf != null
            ? Members(spot.MemberOf, drawing, isAllowed, context)
            : Names(drawing, isAllowed, context);
    }

    /// <summary>The quantities a click on the figure would offer (<see cref="ReferenceText"/>), so that typing and clicking agree</summary>
    static IReadOnlyList<CompletionItem> Members(Token name, Drawing drawing, Predicate<IFigure> isAllowed, InsertContext context)
    {
        var binder = new Binder() { Drawing = drawing };
        var figure = name.IsQuoted ? binder.ResolveFigureExactly(name.Text) : binder.ResolveFigure(name.Text);
        if (figure == null || !isAllowed(figure))
        {
            return [];
        }

        string prefix = Scanner.NameInExpression(figure.Name) + ".";
        var items = new List<CompletionItem>();
        void Add(string member, string description)
        {
            string text = prefix + member;
            items.Add(new CompletionItem(
                member,
                member,
                CompletionItemKind.Member,
                kindText: null,
                figure,
                text,
                () => Capitalize(description) + ValueOf(drawing, text)));
        }

        foreach (var option in ReferenceText.OptionsFor(figure, context))
        {
            string member = option.Text.StartsWith(prefix, StringComparison.Ordinal) ? option.Text.Substring(prefix.Length) : null;
            if (member != null && Scanner.IsName(member))
            {
                Add(member, option.Description);
            }
        }

        // A click puts a segment's length in as its two points (AB); after a dot it is
        // AB.Length, which drawings say all the time
        if (figure is Segment)
        {
            Add(nameof(Segment.Length), "the length of " + ConstructionText.Of(figure));
        }

        return Sorted(items);
    }

    static IReadOnlyList<CompletionItem> Names(Drawing drawing, Predicate<IFigure> isAllowed, InsertContext context)
    {
        var items = new List<CompletionItem>();
        foreach (var figure in drawing.Figures)
        {
            if (string.IsNullOrEmpty(figure.Name) || figure.Auxiliary || !isAllowed(figure))
            {
                continue;
            }

            // what an expression can say by the name: a point, a number, a text that says
            // one, a figure with quantities (c for c.Radius)
            string name = Scanner.NameInExpression(figure.Name);
            var options = ReferenceText.OptionsFor(figure, context);
            if (!options.Any(option => option.Text == name || option.Text.StartsWith(name + ".", StringComparison.Ordinal)))
            {
                continue;
            }

            var titled = figure as FigureBase;
            items.Add(new CompletionItem(
                figure.Name,
                name,
                CompletionItemKind.Figure,
                titled?.TitleParts.Kind?.ToLowerInvariant(),
                figure,
                figure.Title,
                () => figure.Construction));
        }

        foreach (var function in FunctionCatalog.All)
        {
            items.Add(new CompletionItem(
                function.Name,
                function.Name,
                CompletionItemKind.Function,
                "function",
                figure: null,
                function.Signature,
                () => function.Help,
                function.OtherNames));
        }

        items.Add(Reserved("pi", "π, half a turn in radians: 3.14159... Typed as pi or π."));
        items.Add(Reserved("e", "The base of the natural logarithm: 2.71828..."));
        if (context == InsertContext.Function)
        {
            items.Add(Reserved("x", "The variable of the function: y is what it gives for x."));
        }

        return Sorted(items);
    }

    static CompletionItem Reserved(string name, string help)
    {
        return new CompletionItem(name, name, CompletionItemKind.Reserved, "constant", figure: null, name, () => help);
    }

    static IReadOnlyList<CompletionItem> Sorted(List<CompletionItem> items)
    {
        return items
            .OrderBy(item => item.Text, StringComparer.OrdinalIgnoreCase)
            .ThenBy(item => item.Kind)
            .ToList();
    }

    /// <summary>
    /// The rows for the word typed so far: those whose name starts with it, ignoring case,
    /// else those that contain it - the row's own name or another (asin shows arcsin); the
    /// one selected is the exact match (case first), else the first. Firmly selected (a
    /// commit character commits) once the word has a letter; an empty word shows every row,
    /// its first only outlined. A word that already means something no row is - two points
    /// run together, AC - gets a row of its own, first, or a commit character would put
    /// another name in its place (polygon ACDE).
    /// </summary>
    public static CompletionFilter Filter(IReadOnlyList<CompletionItem> items, CompletionSpot spot, Drawing drawing, IFigure owner)
    {
        string word = spot.Word;
        if (word.Length == 0)
        {
            return new CompletionFilter(items, Selected: 0, Hard: false);
        }

        var shown = items.Where(item => item.AnyName(name => name.StartsWith(word, StringComparison.OrdinalIgnoreCase))).ToList();
        if (shown.Count == 0)
        {
            shown = items.Where(item => item.AnyName(name => name.Contains(word, StringComparison.OrdinalIgnoreCase))).ToList();
        }

        if (spot.MemberOf == null && !items.Any(item => item.AnyName(name => name.Equals(word, StringComparison.OrdinalIgnoreCase))))
        {
            var reading = TwoPointsReading(word, drawing, owner);
            if (reading != null)
            {
                shown.Insert(0, reading);
            }
        }

        if (shown.Count == 0)
        {
            return new CompletionFilter(shown, Selected: -1, Hard: false);
        }

        int selected = shown.FindIndex(item => item.Text == word);
        if (selected < 0)
        {
            selected = shown.FindIndex(item => item.AnyName(name => name.Equals(word, StringComparison.OrdinalIgnoreCase)));
        }

        return new CompletionFilter(shown, System.Math.Max(selected, 0), Hard: word.Any(char.IsLetter));
    }

    /// <summary>The distance between two points the word names run together (AC), as the language reads it; null when it doesn't read so</summary>
    static CompletionItem TwoPointsReading(string word, Drawing drawing, IFigure owner)
    {
        if (drawing == null)
        {
            return null;
        }

        var result = Compiler.Instance.CompileExpression(drawing, word, figure => owner == null || !figure.DependsOn(owner));
        if (!result.IsSuccess || result.Dependencies.Count != 2 || result.Dependencies[0] is not IPoint first || result.Dependencies[1] is not IPoint second)
        {
            return null;
        }

        return new CompletionItem(
            word,
            word,
            CompletionItemKind.Figure,
            "distance",
            first,
            word,
            () => "The distance " + ConstructionText.Points(first, second) + ValueOf(drawing, word));
    }

    /// <summary>" = 2.83": what the text is worth now, as the property grid shows a number; nothing when it has no value</summary>
    static string ValueOf(Drawing drawing, string expression)
    {
        var result = Compiler.Instance.CompileExpression(drawing, expression, isFigureAllowed: null);
        if (!result.IsSuccess)
        {
            return "";
        }

        double value = result.Expression();
        return value.IsValidValue() ? " = " + ConstructionText.Number(value) : "";
    }

    static string Capitalize(string text)
    {
        return string.IsNullOrEmpty(text) ? text : char.ToUpperInvariant(text[0]) + text.Substring(1);
    }
}
