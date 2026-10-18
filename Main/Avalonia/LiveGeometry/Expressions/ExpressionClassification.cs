using System;
using System.Collections.Generic;

namespace LiveGeometry;

/// <summary>What a piece of an expression's text is, for its color in the box it is typed into</summary>
public enum ExpressionClass
{
    /// <summary>
    /// An expression in the text of a label, from its [ to its ]: a plate under it, where
    /// the other pieces lie (a plate's span holds others)
    /// </summary>
    Part,

    /// <summary>An expression in the text of a label that doesn't compile: a plate of its own, and nothing in it colored</summary>
    ErrorPart,

    /// <summary>The [ and ] around an expression in the text of a label</summary>
    Bracket,

    /// <summary>A figure by its name: A in A.X, the points of distance(A, B) and of AB, a slider, a number</summary>
    Figure,

    /// <summary>What is read off a figure: the X of A.X, the Length of AB.Length</summary>
    Property,

    /// <summary>A function: sin, sqrt, distance, angle, log_2</summary>
    Function,

    /// <summary>A number as typed: 3, 0.5</summary>
    Number,

    /// <summary>What the language gives a meaning of its own: pi, π, e, the x of a function, the degree sign</summary>
    Reserved
}

/// <summary>A piece of an expression's text and what it is</summary>
public readonly record struct ClassifiedSpan(int Start, int Length, ExpressionClass Class)
{
    public int End => Start + Length;
}

/// <summary>
/// What the pieces of a text are, as the compiler reads it (<see cref="CompileResult.Classification"/>):
/// the names as they were resolved against the drawing, not as a second reading of the text
/// might guess them (AB is two points or a slider ab, PI the constant or points P and I).
/// Only a text that compiles has pieces; what is wrong is said under the box on commit (or,
/// in a label's text, by a plate). None without a drawing, whose names they would be.
/// </summary>
public static class ExpressionClassifier
{
    public static IReadOnlyList<ClassifiedSpan> ClassifyExpression(Drawing drawing, string text, Predicate<IFigure> isFigureAllowed)
    {
        return drawing == null ? [] : Pieces(Compiler.Instance.CompileExpression(drawing, text, isFigureAllowed));
    }

    /// <summary>An expression in x: the x of a function is its variable</summary>
    public static IReadOnlyList<ClassifiedSpan> ClassifyFunction(Drawing drawing, string text, Predicate<IFigure> isFigureAllowed)
    {
        return drawing == null ? [] : Pieces(Compiler.Instance.CompileFunction(drawing, text, isFigureAllowed));
    }

    /// <summary>
    /// The text of a label: each [...] part, whole, then the brackets and the pieces of one
    /// that compiles. What is wrong with one that doesn't, the label says on the canvas; its
    /// plate says here that something is. A [ with no ] after it is text like the rest.
    /// </summary>
    public static IReadOnlyList<ClassifiedSpan> ClassifyLabelText(Drawing drawing, string text, Predicate<IFigure> isFigureAllowed)
    {
        var spans = new List<ClassifiedSpan>();
        if (drawing == null)
        {
            return spans;
        }

        foreach (var part in LabelBase.FindExpressionParts(text))
        {
            var result = Compiler.Instance.CompileExpression(drawing, text.Substring(part.Start + 1, part.Length - 2), isFigureAllowed);
            if (!result.IsSuccess)
            {
                spans.Add(new ClassifiedSpan(part.Start, part.Length, ExpressionClass.ErrorPart));
                continue;
            }

            spans.Add(new ClassifiedSpan(part.Start, part.Length, ExpressionClass.Part));
            spans.Add(new ClassifiedSpan(part.Start, 1, ExpressionClass.Bracket));
            foreach (var piece in result.Classification)
            {
                spans.Add(piece with { Start = piece.Start + part.Start + 1 });
            }

            spans.Add(new ClassifiedSpan(part.Start + part.Length - 1, 1, ExpressionClass.Bracket));
        }

        return spans;
    }

    static IReadOnlyList<ClassifiedSpan> Pieces(CompileResult result)
    {
        return result.IsSuccess ? result.Classification : [];
    }
}
