using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.VisualTree;

namespace LiveGeometry;

/// <summary>
/// Who an <see cref="ExpressionTextBox"/> types for (the editor around it): which drawing the
/// names are of, which figure's expression it is, what kind of text it holds
/// </summary>
public interface IExpressionHost
{
    /// <summary>The drawing the expression names figures of; null to find it from where the box is</summary>
    Drawing Drawing { get; }

    /// <summary>The figure whose expression is typed, which nothing it depends on may name; null for a tool's panel</summary>
    IFigure Owner { get; }

    InsertContext Context { get; }
}

/// <summary>
/// The box every expression is typed into: a formula, a coordinate, a function, the text of
/// a label (2026-10-15). While it has the keyboard it is the drawing's
/// <see cref="Drawing.InsertTarget"/>, so a click on a figure or on its row in the Figure
/// List puts a reference to the figure in at the caret (<see cref="ReferencePicker"/>); in
/// the text of a label the reference goes in as a [...] part unless the caret is in one
/// already. One class for all of them, so that completion, highlighting or a structured
/// editor can come to every box at once.
/// </summary>
public class ExpressionTextBox : TextBox, IExpressionInsertTarget
{
    // (a subclass of a templated control renders invisible without this)
    protected override Type StyleKeyOverride => typeof(TextBox);

    public ExpressionTextBox()
    {
        GotFocus += (s, e) => BecomeTarget();

        // A press on the canvas or on the Figure List takes the keyboard before its
        // PointerPressed is raised: the box stays the target until that press has been
        // handled, and takes the keyboard back when the press put a reference in (Insert)
        LostFocus += (s, e) => Dispatcher.UIThread.Post(ClearIfUnfocused, DispatcherPriority.Input);
    }

    /// <summary>
    /// The box has the keyboard: clicks on figures put references in, and a pencil at the
    /// box's right says so for as long as that holds
    /// </summary>
    void BecomeTarget()
    {
        Drawing?.SetInsertTarget(this);
        if (InnerRightContent == null)
        {
            var mark = new Border()
            {
                Margin = new Thickness(0, 0, 6, 0),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                Child = PropertyGridIcons.Create(PropertyGridIcon.Pencil)
            };
            ToolTip.SetTip(mark, ReferencePicker.Hint);
            InnerRightContent = mark;
        }
    }

    void ClearIfUnfocused()
    {
        if (!IsFocused)
        {
            Drawing?.ClearInsertTarget(this);
            InnerRightContent = null;
        }
    }

    /// <summary>
    /// Done with the box (Enter in a figure's row): the keyboard goes to the canvas, so that
    /// a click on a figure is the tool's again
    /// </summary>
    public void EndEditing()
    {
        this.FindAncestorOfType<DrawingHost>()?.DrawingControl.Focus();
    }

    public IExpressionHost Host { get; set; }

    /// <summary>The host's drawing, or the one of the editor the box is in</summary>
    public Drawing Drawing => Host?.Drawing ?? this.FindAncestorOfType<DrawingHost>()?.CurrentDrawing;

    public IFigure Owner => Host?.Owner;

    public InsertContext Context => Host?.Context ?? InsertContext.Expression;

    // a panel rebuilt around the box takes it out of the tree without a LostFocus
    protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
    {
        Drawing?.ClearInsertTarget(this);
        base.OnDetachedFromVisualTree(e);
    }

    /// <summary>
    /// Puts the reference in at the caret, over the selection if there is one, as typing it
    /// would (the editor applies it as it applies a keystroke)
    /// </summary>
    public void Insert(string reference)
    {
        if (IsReadOnly || !IsEnabled)
        {
            return;
        }

        string text = Context == InsertContext.Text && !IsInBrackets(Text ?? "", CaretIndex)
            ? "[" + reference + "]"
            : reference;
        SelectedText = text;
        Focus();
    }

    /// <summary>
    /// Whether the caret is inside a [...] part of a label's text: after a [ that no ] has
    /// closed before the caret (one typed just now counts, so that a reference clicked after
    /// it doesn't bring brackets of its own)
    /// </summary>
    public static bool IsInBrackets(string text, int caret)
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

        return open >= 0;
    }

    /// <summary>
    /// The value of a text as it stands, for a plate under the box: "= 18.85", "= undefined"
    /// for one without a value right now; null for one that doesn't compile, or is empty
    /// (what is wrong is said on commit)
    /// </summary>
    public string LiveValueText(string text)
    {
        var drawing = Drawing;
        if (drawing == null || string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        var owner = Owner;
        var result = Compiler.Instance.CompileExpression(drawing, text, figure => owner == null || !figure.DependsOn(owner));
        if (!result.IsSuccess || result.Expression == null)
        {
            return null;
        }

        double value;
        try
        {
            value = result.Expression();
        }
        catch (Exception)
        {
            return null;
        }

        return "= " + (value.IsValidValue() ? Math.Round(value, Settings.DisplayDecimals).ToString() : LabelBase.UndefinedText);
    }
}

/// <summary>What kind of text a row holds, from its caption: the X and Y rows take a point as its X or Y</summary>
public static class InsertContexts
{
    public static InsertContext FromRowName(string rowName)
    {
        var name = (rowName ?? "").Trim().TrimEnd('=').TrimEnd();
        if (name.EndsWith("X", StringComparison.Ordinal))
        {
            return InsertContext.PointX;
        }

        if (name.EndsWith("Y", StringComparison.Ordinal))
        {
            return InsertContext.PointY;
        }

        return InsertContext.Expression;
    }
}
