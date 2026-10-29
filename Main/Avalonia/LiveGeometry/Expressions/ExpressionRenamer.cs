using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LiveGeometry;

/// <summary>
/// A figure whose text names other figures in expressions (a label's [A.X], a function
/// graph's f(x)): when figures are renamed, the names in its text follow.
/// </summary>
public interface IRenamableExpressions
{
    /// <summary>
    /// Only the text: what the expressions compiled to already holds the figures themselves,
    /// not their names
    /// </summary>
    void RenameInExpressions(ExpressionRenamer renamer);

    /// <summary>
    /// Compiles the expressions again from their text: a figure they name was replaced by
    /// another under the same name (<see cref="Actions.ReplacePoint"/>), and what they
    /// compiled to still holds the one that left
    /// </summary>
    void RebindExpressions();

    /// <summary>
    /// The texts of the expressions, to keep and to put back as they were: undo of a join
    /// (<see cref="PointSnapping.Join"/>), where renaming back could not tell the names that
    /// were the joined point's from the target's own. Setting them compiles nothing
    /// (<see cref="RebindExpressions"/> does).
    /// </summary>
    IReadOnlyList<string> ExpressionTexts { get; set; }
}

/// <summary>
/// Rewrites expressions after a wave of renames (a point, the segments named after it, a
/// figure that had to give up its name): every name that meant a renamed figure under the old
/// names is replaced by its new name, all at once, so that A and B trading names come out
/// right. Names are read the way <see cref="ExpressionTreeBuilder"/> binds them - A.X, AB for
/// the distance between two points, the points of angle(A, B, C) and area(...), the name of a
/// Number, a name in backticks - and the rest of the text is left exactly as it was. A new
/// name goes in as an expression says it (<see cref="Scanner.NameInExpression"/>: `Drag me!`,
/// `pi`). Two points run together whose new names would read differently (PB next to a point
/// named PB), or one of which needs backticks, become distance(P, B).
/// </summary>
public class ExpressionRenamer
{
    readonly Drawing drawing;
    readonly IReadOnlyDictionary<IFigure, string> oldNames;
    readonly IReadOnlyList<IFigure> preferred;

    /// <param name="oldNames">Each renamed figure with the name it had</param>
    /// <param name="preferred">
    /// Figures a name means before any other figure of the drawing called that: pasted
    /// copies, whose texts name them by the names the originals still have
    /// </param>
    public ExpressionRenamer(Drawing drawing, IReadOnlyDictionary<IFigure, string> oldNames, IReadOnlyList<IFigure> preferred = null)
    {
        this.drawing = drawing;
        this.oldNames = oldNames;
        this.preferred = preferred ?? new IFigure[0];
    }

    // the figures names are looked up among, in the order they are tried
    IEnumerable<IFigure> TopLevel => preferred.Concat(drawing.Figures);

    IEnumerable<IFigure> AllFigures => preferred.Concat(drawing.Figures.GetAllFiguresRecursive());

    /// <param name="isFunction">A function of x (a graph), where x is the variable and not a figure</param>
    public string Rewrite(string expression, bool isFunction)
    {
        if (string.IsNullOrEmpty(expression))
        {
            return expression;
        }

        var parsed = Parser.Parse(expression);
        if (parsed.Root == null || !parsed.Errors.IsEmpty())
        {
            return expression;
        }

        var replacements = new List<(int Start, int Length, string Text)>();
        Visit(parsed.Root, replacements, isFunction);
        if (replacements.Count == 0)
        {
            return expression;
        }

        var sb = new StringBuilder(expression);
        foreach (var replacement in replacements.OrderByDescending(r => r.Start))
        {
            sb.Remove(replacement.Start, replacement.Length);
            sb.Insert(replacement.Start, replacement.Text);
        }

        return sb.ToString();
    }

    /// <summary>The name a figure had before the wave</summary>
    string OldName(IFigure figure)
    {
        return oldNames.TryGetValue(figure, out var oldName) ? oldName : figure.Name;
    }

    void Visit(Node node, List<(int Start, int Length, string Text)> replacements, bool isFunction)
    {
        if (node == null)
        {
            return;
        }

        switch (node.Kind)
        {
            case NodeType.Variable:
                RenameVariable(node.Token, replacements, isFunction);
                return;
            case NodeType.PropertyAccess:
                // A.X: the name before the dot; X is a property
                RenameFigure(node.Children[0].Token, replacements);
                return;
            case NodeType.FunctionCall:
                int joinedCount = ExpressionTreeBuilder.JoinedPointCount(node);
                if (joinedCount > 0)
                {
                    RenameJoinedPoints(node, joinedCount, replacements);
                    return;
                }

                if (!ExpressionReflection.TakesPoints(node.Token.Text, node.Children.Count))
                {
                    foreach (var argument in node.Children)
                    {
                        Visit(argument, replacements, isFunction);
                    }

                    return;
                }

                // angle(A, B, C), distance(A, B), area(A, B, C, D): the names of points
                foreach (var argument in node.Children)
                {
                    if (argument != null && argument.Kind == NodeType.Variable)
                    {
                        RenameFigure(argument.Token, replacements);
                    }
                }

                return;
        }

        foreach (var child in node.Children)
        {
            Visit(child, replacements, isFunction);
        }
    }

    /// <summary>
    /// A bare name, in the order the compiler tries it: two points (AB, the distance), then
    /// pi, e and a function's x, then a figure (a Number)
    /// </summary>
    void RenameVariable(Token token, List<(int Start, int Length, string Text)> replacements, bool isFunction)
    {
        var text = token.Text;
        if (token.IsQuoted)
        {
            RenameFigure(token, replacements);
            return;
        }

        // a number called exactly this, as the compiler takes it (Binder.ResolveExactNumber)
        bool isConstant = text.Equals("pi", StringComparison.InvariantCultureIgnoreCase)
            || text == "π"
            || text.Equals("e", StringComparison.InvariantCultureIgnoreCase)
            || isFunction && text == "x";
        var exact = isConstant ? null : TopLevel.FirstOrDefault(f => f is INumber && OldName(f) == text);
        if (exact != null)
        {
            if (IsRenamed(exact))
            {
                replacements.Add((token.Start, token.Length, Scanner.NameInExpression(exact.Name)));
            }

            return;
        }

        // pi and e in lowercase are the constants, also next to points P and I
        if (text == "pi" || text == "π" || text == "e")
        {
            return;
        }

        var twoPoints = SplitTwoPoints(text, OldName);
        if (twoPoints != null)
        {
            var (first, second, firstLength) = twoPoints.Value;
            if (IsRenamed(first) || IsRenamed(second))
            {
                var firstText = IsRenamed(first) ? first.Name : text.Substring(0, firstLength);
                var secondText = IsRenamed(second) ? second.Name : text.Substring(firstLength);
                var joined = firstText + secondText;

                // the new names run together may split differently (P and B next to a point
                // named PB), or one of them may need backticks: then the distance says its
                // points apart
                var check = SplitTwoPoints(joined, figure => figure.Name);
                if (check == null
                    || check.Value.First != first
                    || check.Value.Second != second
                    || !Scanner.IsName(first.Name)
                    || !Scanner.IsName(second.Name))
                {
                    joined = "distance(" + Scanner.NameInExpression(first.Name) + ", " + Scanner.NameInExpression(second.Name) + ")";
                }

                replacements.Add((token.Start, token.Length, joined));
            }

            return;
        }

        if (text.Equals("pi", StringComparison.InvariantCultureIgnoreCase)
            || text.Equals("e", StringComparison.InvariantCultureIgnoreCase)
            || isFunction && text == "x")
        {
            return;
        }

        RenameFigure(token, replacements);
    }

    /// <summary>
    /// ∠ABC, angle(ABC): the points in one name, the ones renamed given their new names. Where
    /// the new names run together would read differently, or one needs backticks, the names
    /// go apart: ∠(P, B, C), angle(P, B, C).
    /// </summary>
    void RenameJoinedPoints(Node call, int count, List<(int Start, int Length, string Text)> replacements)
    {
        var token = call.Children[0].Token;
        var points = ExpressionTreeBuilder.SplitNames(token.Text, count, TopLevel.OfType<PointBase>().ToArray(), OldName);
        if (points == null || !points.Any(IsRenamed))
        {
            return;
        }

        var texts = new List<string>();
        int start = 0;
        foreach (var point in points)
        {
            int length = OldName(point).Length;
            texts.Add(IsRenamed(point) ? point.Name : token.Text.Substring(start, length));
            start += length;
        }

        var joined = string.Concat(texts);
        var check = ExpressionTreeBuilder.SplitNames(joined, count, TopLevel.OfType<PointBase>().ToArray(), figure => figure.Name);
        if (check == null || !check.SequenceEqual(points) || !points.All(point => Scanner.IsName(point.Name)))
        {
            var apart = string.Join(", ", points.Select(point => Scanner.NameInExpression(point.Name)));
            joined = call.Token.Kind == TokenType.Angle ? "(" + apart + ")" : apart;
        }

        replacements.Add((token.Start, token.Length, joined));
    }

    void RenameFigure(Token token, List<(int Start, int Length, string Text)> replacements)
    {
        if (token == null || token.Kind != TokenType.Identifier)
        {
            return;
        }

        // a name in backticks is the figure called exactly that (Binder.ResolveFigureExactly)
        var figure = token.IsQuoted
            ? TopLevel.FirstOrDefault(f => f != null && OldName(f) == token.Text)
            : ResolveFigure(token.Text);
        if (figure != null && IsRenamed(figure))
        {
            replacements.Add((token.Start, token.Length, Scanner.NameInExpression(figure.Name)));
        }
    }

    bool IsRenamed(IFigure figure)
    {
        return oldNames.TryGetValue(figure, out var oldName) && oldName != figure.Name;
    }

    /// <summary><see cref="Binder.ResolveFigure"/> under the old names</summary>
    IFigure ResolveFigure(string name)
    {
        return AllFigures.FirstOrDefault(f => OldName(f) == name)
            ?? TopLevel.FirstOrDefault(f => f != null && OldName(f) == name)
            ?? TopLevel.FirstOrDefault(f => f != null
                && !OldName(f).IsEmpty()
                && OldName(f).Equals(name, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// <see cref="ExpressionTreeBuilder.ResolveTwoPoints"/> under the old names or the new: the
    /// longest point name the text starts with and the longest it ends with, together the whole text
    /// </summary>
    (PointBase First, PointBase Second, int FirstLength)? SplitTwoPoints(string text, Func<IFigure, string> nameOf)
    {
        var points = TopLevel.OfType<PointBase>().ToArray();
        string longestPrefix = "";
        string longestSuffix = "";
        foreach (var point in points)
        {
            var name = nameOf(point);
            if (string.IsNullOrEmpty(name))
            {
                continue;
            }

            if (text.StartsWith(name, StringComparison.OrdinalIgnoreCase) && name.Length > longestPrefix.Length)
            {
                longestPrefix = name;
            }

            if (text.EndsWith(name, StringComparison.OrdinalIgnoreCase) && name.Length > longestSuffix.Length)
            {
                longestSuffix = name;
            }
        }

        if (longestPrefix.Length == 0 || longestPrefix.Length + longestSuffix.Length != text.Length)
        {
            return null;
        }

        var first = points.FirstOrDefault(p => nameOf(p) == longestPrefix);
        var second = points.FirstOrDefault(p => nameOf(p) == longestSuffix);
        if (first == null || second == null)
        {
            return null;
        }

        return (first, second, longestPrefix.Length);
    }
}
