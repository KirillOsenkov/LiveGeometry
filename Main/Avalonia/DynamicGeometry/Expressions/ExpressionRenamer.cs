using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DynamicGeometry;

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
}

/// <summary>
/// Rewrites expressions after a wave of renames (a point, the segments named after it, a
/// figure that had to give up its name): every name that meant a renamed figure under the old
/// names is replaced by its new name, all at once, so that A and B trading names come out
/// right. Names are read the way <see cref="ExpressionTreeBuilder"/> binds them - A.X, AB for
/// the distance between two points, the points of ang(A, B, C) and area(...), the name of a
/// Number - and the rest of the text is left exactly as it was. Two points run together whose
/// new names would read differently (PB next to a point named PB) become dist(P, B).
/// </summary>
public class ExpressionRenamer
{
    readonly Drawing drawing;
    readonly IReadOnlyDictionary<IFigure, string> oldNames;

    /// <param name="oldNames">Each renamed figure with the name it had</param>
    public ExpressionRenamer(Drawing drawing, IReadOnlyDictionary<IFigure, string> oldNames)
    {
        this.drawing = drawing;
        this.oldNames = oldNames;
    }

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
                if (TakesNumber(node))
                {
                    Visit(node.Children[0], replacements, isFunction);
                    return;
                }

                // ang(A, B, C), dist(A, B), area(A, B, C, D): the names of points
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

    /// <summary>sin(...), sqrt(...): a function of a number and not of points, as the compiler tells them apart</summary>
    static bool TakesNumber(Node call)
    {
        var method = new Binder().ResolveMethod(call.Token.Text);
        if (method == null || call.Children.Count != 1)
        {
            return false;
        }

        var parameters = method.GetParameters();
        return parameters.Length == 1 && parameters[0].ParameterType == typeof(double);
    }

    /// <summary>
    /// A bare name, in the order the compiler tries it: two points (AB, the distance), then
    /// pi, e and a function's x, then a figure (a Number)
    /// </summary>
    void RenameVariable(Token token, List<(int Start, int Length, string Text)> replacements, bool isFunction)
    {
        var text = token.Text;
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
                // named PB): then the distance says its points apart
                var check = SplitTwoPoints(joined, figure => figure.Name);
                if (check == null || check.Value.First != first || check.Value.Second != second)
                {
                    joined = "dist(" + first.Name + ", " + second.Name + ")";
                }

                replacements.Add((token.Start, text.Length, joined));
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

    void RenameFigure(Token token, List<(int Start, int Length, string Text)> replacements)
    {
        if (token == null || token.Kind != TokenType.Identifier)
        {
            return;
        }

        var figure = ResolveFigure(token.Text);
        if (figure != null && IsRenamed(figure))
        {
            replacements.Add((token.Start, token.Text.Length, figure.Name));
        }
    }

    bool IsRenamed(IFigure figure)
    {
        return oldNames.TryGetValue(figure, out var oldName) && oldName != figure.Name;
    }

    /// <summary><see cref="Binder.ResolveFigure"/> under the old names</summary>
    IFigure ResolveFigure(string name)
    {
        return drawing.Figures.GetAllFiguresRecursive().FirstOrDefault(f => OldName(f) == name)
            ?? drawing.Figures.FirstOrDefault(f => f != null && OldName(f) == name)
            ?? drawing.Figures.FirstOrDefault(f => f != null
                && !OldName(f).IsEmpty()
                && OldName(f).Equals(name, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// <see cref="ExpressionTreeBuilder.ResolveTwoPoints"/> under the old names or the new: the
    /// longest point name the text starts with and the longest it ends with, together the whole text
    /// </summary>
    (PointBase First, PointBase Second, int FirstLength)? SplitTwoPoints(string text, Func<IFigure, string> nameOf)
    {
        var points = drawing.Figures.OfType<PointBase>().ToArray();
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
