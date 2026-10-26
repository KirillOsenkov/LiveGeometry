using System;
using System.Collections.Generic;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
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
/// already. Its pieces are in colors and it has a completion list, as every
/// <see cref="ColoredTextBox"/> does. One class for all of them, so that a structured editor
/// can come to every box at once.
/// </summary>
public class ExpressionTextBox : ColoredTextBox, IExpressionInsertTarget
{
    public ExpressionTextBox()
    {
        GotFocus += (s, e) => BecomeTarget();

        // A press on the canvas or on the Figure List takes the keyboard before its
        // PointerPressed is raised: the box stays the target until that press has been
        // handled, and takes the keyboard back when the press put a reference in (Insert)
        LostFocus += (s, e) => Dispatcher.UIThread.Post(ClearIfUnfocused, DispatcherPriority.Input);

        // the pencil: the picking off and on while the keyboard stays in the box
        // (TogglePicking). Off, a click on a figure selects it, as it does without a box.
        // The way out of a label's text without leaving it, since Enter there starts a line.
        brackets = CreateBracketsButton();
        pencil = CreateButton(PropertyGridIcons.Create(PropertyGridIcon.Pencil), TogglePicking);
        buttons = new StackPanel()
        {
            Orientation = Avalonia.Layout.Orientation.Horizontal,
            Children = { brackets, CreateDegreeButton(), pencil }
        };
        ShowButtons(false);
        InnerRightContent = buttons;
    }

    StackPanel buttons;
    Border brackets;
    Border pencil;

    // The drawing the box made itself the target of. Taken out of the tree (a tool's panel
    // going away with the tool) it can't find the drawing through its ancestors any more,
    // and stayed the target: every click on a figure went into a box no longer on screen,
    // not to the tool (the Midpoint tool after the Coordinates tool's panel).
    Drawing targetOf;

    /// <summary>
    /// The box has the keyboard: clicks on figures put references in, and a pencil at the
    /// box's right says so for as long as that holds. In a label's text a [ ] button comes
    /// before it, which also says that the text takes formulas at all.
    /// </summary>
    void BecomeTarget()
    {
        SetTarget(Drawing);
        brackets.IsVisible = Context == InsertContext.Text;
        ShowButtons(true);
        ShowPicking(true);
    }

    /// <summary>
    /// The buttons at the box's right while it has the keyboard. Without it they keep their
    /// height and give up their width, unseen: the box doesn't grow taller on a click into
    /// it, and its text has the whole width while it isn't edited.
    /// </summary>
    void ShowButtons(bool show)
    {
        buttons.Width = show ? double.NaN : 0;
        buttons.Opacity = show ? 1 : 0;
        buttons.IsHitTestVisible = show;
    }

    /// <summary>The ° button: the degree sign, which no keyboard has a key for (sin(30°))</summary>
    Border CreateDegreeButton()
    {
        var degree = new TextBlock()
        {
            Text = "°",
            FontSize = PropertyGridIcons.Size + 2,
            FontWeight = FontWeight.SemiBold,
            Margin = new Thickness(2, 0),
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center
        };
        degree.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.ExpressionReserved));
        var button = CreateButton(degree, () =>
        {
            if (!IsReadOnly && IsEnabled)
            {
                SelectedText = "°";
            }
        });
        ToolTip.SetTip(button, "Insert the degree sign: sin(30°) is the sine of 30 degrees");
        return button;
    }

    /// <summary>
    /// A button at the box's right end that leaves the keyboard and the caret where they are:
    /// the box's own press would move the caret and take the pointer, and the release would
    /// never come here
    /// </summary>
    Border CreateButton(Control content, Action click)
    {
        var button = new Border()
        {
            Margin = new Thickness(0, 0, 3, 0),
            Padding = new Thickness(3),
            CornerRadius = new CornerRadius(4),
            Background = Brushes.Transparent, // hit-testable around the icon too
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
            Cursor = new Cursor(StandardCursorType.Hand),
            Child = content
        };
        button.PointerEntered += (s, e) => button.BindTheme(Border.BackgroundProperty, nameof(AppTheme.ButtonHover));
        button.PointerExited += (s, e) => button.BindTheme(Border.BackgroundProperty, key: null, whenNone: Brushes.Transparent);
        button.PointerPressed += (s, e) =>
        {
            if (e.GetCurrentPoint(button).Properties.IsLeftButtonPressed)
            {
                e.Pointer.Capture(button);
                e.Handled = true;
            }
        };
        button.PointerReleased += (s, e) =>
        {
            if (e.InitialPressMouseButton != MouseButton.Left || e.Pointer.Captured != button)
            {
                return;
            }

            e.Pointer.Capture(null);
            e.Handled = true;
            if (new Rect(button.Bounds.Size).Contains(e.GetPosition(button)))
            {
                click();
            }
        };
        return button;
    }

    /// <summary>The [ ] button of a label's text, in the color the brackets of a formula take in it</summary>
    Border CreateBracketsButton()
    {
        var brackets = new TextBlock()
        {
            Text = "[ ]",
            FontSize = PropertyGridIcons.Size,
            FontWeight = FontWeight.SemiBold,
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center
        };
        brackets.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.ExpressionBracket));
        var button = CreateButton(brackets, InsertBrackets);
        ToolTip.SetTip(button, "Insert formulas in [] such as [AB] or [A.X+B.X]");
        return button;
    }

    /// <summary>
    /// The [ ] button's click: what is selected goes in brackets, the caret after them; with
    /// nothing selected a pair goes in at the caret, the caret between them for the formula
    /// </summary>
    void InsertBrackets()
    {
        if (IsReadOnly || !IsEnabled)
        {
            return;
        }

        int start = System.Math.Min(SelectionStart, SelectionEnd);
        string selected = SelectedText ?? "";
        SelectedText = "[" + selected + "]";
        CaretIndex = selected.Length == 0 ? start + 1 : start + selected.Length + 2;
    }

    /// <summary>The pencil's click: clicks on figures stop or start putting references in, and the box keeps the keyboard</summary>
    void TogglePicking()
    {
        var drawing = Drawing;
        if (drawing == null)
        {
            return;
        }

        bool picking = drawing.InsertTarget != this;
        if (picking)
        {
            SetTarget(drawing);
        }
        else
        {
            ClearTarget();
        }

        ShowPicking(picking);
    }

    void SetTarget(Drawing drawing)
    {
        if (targetOf != drawing)
        {
            ClearTarget();
        }

        targetOf = drawing;
        drawing?.SetInsertTarget(this);
    }

    void ClearTarget()
    {
        targetOf?.ClearInsertTarget(this);
        targetOf = null;
    }

    /// <summary>The pencil as ink while clicks on figures come here, faint while they don't</summary>
    void ShowPicking(bool picking)
    {
        if (pencil == null)
        {
            return;
        }

        pencil.Child.Opacity = picking ? 1 : 0.35;
        ToolTip.SetTip(pencil, picking
            ? "Formula picker enabled: click a figure to insert it into the formula."
            : "Click this button to enable the formula picker: click a figure to insert it into the formula.");
    }

    void ClearIfUnfocused()
    {
        if (!IsFocused)
        {
            ClearTarget();
            ShowButtons(false);
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

    /// <summary>The host's drawing, or the one of the editor the box is in</summary>
    public override Drawing Drawing => base.Drawing ?? this.FindAncestorOfType<DrawingHost>()?.CurrentDrawing;

    protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnAttachedToVisualTree(e);
        Watch(Drawing);
    }

    // a panel rebuilt around the box takes it out of the tree without a LostFocus
    protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
    {
        ClearTarget();
        Watch(null);
        base.OnDetachedFromVisualTree(e);
    }

    Drawing watched;

    /// <summary>
    /// The colors follow the drawing while the box shows: what a name stands for changes
    /// with the figures (one renamed, made, deleted), each such change an entry of the
    /// history. At most every 300 ms, as the Figure List follows it: a drag records a move
    /// at every step.
    /// </summary>
    void Watch(Drawing drawing)
    {
        if (watched == drawing)
        {
            return;
        }

        if (watched != null)
        {
            watched.ActionManager.CollectionChanged -= ActionManager_CollectionChanged;
        }

        watched = drawing;
        if (watched != null)
        {
            watched.ActionManager.CollectionChanged += ActionManager_CollectionChanged;
        }
    }

    void ActionManager_CollectionChanged(object sender, EventArgs e)
    {
        Throttle.Schedule(
            this,
            box => Dispatcher.UIThread.Post(() =>
            {
                // (gone from the screen while the post waited)
                if (box.watched != null)
                {
                    box.Recolor();
                }
            }),
            TimeSpan.FromMilliseconds(300),
            ThrottleOptions.RunOnceImmediatelyIfFree);
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
        return ExpressionCompletion.OpenBracketBefore(text, caret) >= 0;
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
