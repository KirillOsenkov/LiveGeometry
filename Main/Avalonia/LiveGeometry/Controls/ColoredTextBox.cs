using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Presenters;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Templates;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Media.Immutable;
using Avalonia.Media.TextFormatting;
using Avalonia.Styling;
using Avalonia.Utilities;
using AvaloniaSetter = Avalonia.Styling.Setter;
using AvaloniaStyle = Avalonia.Styling.Style;

namespace LiveGeometry;

/// <summary>
/// A box an expression is typed into, its text drawn in colors - each piece in the theme's
/// color for what the compiler reads it as (<see cref="Classify"/>) - and with a completion
/// list (<see cref="Completion"/>): <see cref="ExpressionTextBox"/>, which also takes
/// references by a click, the text of a label edited where it is (<see cref="InPlaceTextEditor"/>),
/// the theme page's sample. Whose names it says and what kind of text it holds is its
/// <see cref="Host"/>'s to say; without one it is a plain box. A TextBox draws its text in
/// one brush; the template here is Fluent's (TextBox.xaml of Avalonia 12.1) with a presenter
/// that draws it in several (<see cref="ColoredTextPresenter"/>) and the completion's popup,
/// and without what no box of the app uses: the placeholder texts and the data validation
/// errors. Fluent's styles reach it as they reach a TextBox, by the names of the parts.
/// </summary>
public class ColoredTextBox : TextBox
{
    // (a subclass of a templated control renders invisible without this)
    protected override Type StyleKeyOverride => typeof(TextBox);

    public ColoredTextBox()
    {
        Template = template;
        Completion = new CompletionList(this);
    }

    IExpressionHost host;

    /// <summary>Who the box types for: the drawing whose names it says, the figure whose expression it is, the kind of text; null for plain text</summary>
    public IExpressionHost Host
    {
        get => host;
        set
        {
            host = value;
            classifiedText = null;
            presenter?.Recolor();
        }
    }

    /// <summary>The drawing whose names the text says: the host's</summary>
    public virtual Drawing Drawing => Host?.Drawing;

    /// <summary>The figure whose expression is typed, which nothing it depends on may name; null for a tool's panel</summary>
    public IFigure Owner => Host?.Owner;

    public InsertContext Context => Host?.Context ?? InsertContext.Expression;

    /// <summary>The list of what can go at the caret (docs/completion-plan.md)</summary>
    public CompletionList Completion { get; }

    string classifiedText;
    IReadOnlyList<ClassifiedSpan> pieces = [];
    ColoredTextPresenter presenter;

    /// <summary>The pieces of the text, worked out once per text: the presenter asks at every layout</summary>
    public IReadOnlyList<ClassifiedSpan> PiecesOf(string text)
    {
        text ??= "";
        if (text != classifiedText)
        {
            classifiedText = text;
            pieces = Classify(text);
        }

        return pieces;
    }

    /// <summary>
    /// The text may mean something else now that the drawing has changed (a name it says
    /// was given to a figure, or taken away): its pieces are worked out again, and it is
    /// drawn again if they changed
    /// </summary>
    public void Recolor()
    {
        if (classifiedText == null)
        {
            return;
        }

        var fresh = Classify(classifiedText);
        if (!fresh.SequenceEqual(pieces))
        {
            pieces = fresh;
            presenter?.Recolor();
        }
    }

    /// <summary>
    /// The pieces of the text as the compiler reads it in this box: an expression, a function
    /// in x, or the [...] parts of a label's text. Only a text that compiles has pieces
    /// (<see cref="ExpressionClassifier"/>); none without a drawing.
    /// </summary>
    IReadOnlyList<ClassifiedSpan> Classify(string text)
    {
        var drawing = Drawing;
        var owner = Owner;
        Predicate<IFigure> isFigureAllowed = figure => owner == null || !figure.DependsOn(owner);
        return Context switch
        {
            InsertContext.Text => ExpressionClassifier.ClassifyLabelText(drawing, text, isFigureAllowed),
            InsertContext.Function => ExpressionClassifier.ClassifyFunction(drawing, text, isFigureAllowed),
            _ => ExpressionClassifier.ClassifyExpression(drawing, text, isFigureAllowed)
        };
    }

    protected override void OnApplyTemplate(TemplateAppliedEventArgs e)
    {
        base.OnApplyTemplate(e);
        presenter = e.NameScope.Find<ColoredTextPresenter>("PART_TextPresenter");
        Completion.Attach(e.NameScope.Find<Popup>("PART_CompletionPopup"), presenter);
    }

    static readonly IControlTemplate template = new FuncControlTemplate<TextBox>((box, scope) =>
    {
        var presenter = new ColoredTextPresenter()
        {
            Name = "PART_TextPresenter",
            [!!TextPresenter.TextProperty] = box[!!TextProperty],
            [!TextPresenter.CaretBlinkIntervalProperty] = box[!CaretBlinkIntervalProperty],
            [!TextPresenter.CaretIndexProperty] = box[!CaretIndexProperty],
            [!TextPresenter.SelectionStartProperty] = box[!SelectionStartProperty],
            [!TextPresenter.SelectionEndProperty] = box[!SelectionEndProperty],
            [!TextPresenter.TextAlignmentProperty] = box[!TextAlignmentProperty],
            [!TextPresenter.TextWrappingProperty] = box[!TextWrappingProperty],
            [!TextPresenter.LineHeightProperty] = box[!LineHeightProperty],
            [!TextPresenter.LetterSpacingProperty] = box[!LetterSpacingProperty],
            [!TextPresenter.PasswordCharProperty] = box[!PasswordCharProperty],
            [!TextPresenter.RevealPasswordProperty] = box[!RevealPasswordProperty],
            [!TextPresenter.SelectionBrushProperty] = box[!SelectionBrushProperty],
            [!TextPresenter.SelectionForegroundBrushProperty] = box[!SelectionForegroundBrushProperty],
            [!TextPresenter.CaretBrushProperty] = box[!CaretBrushProperty],
            [!HorizontalAlignmentProperty] = box[!HorizontalContentAlignmentProperty],
            [!VerticalAlignmentProperty] = box[!VerticalContentAlignmentProperty]
        }.RegisterInNameScope(scope);

        var scrollViewer = new ScrollViewer()
        {
            Name = "PART_ScrollViewer",
            [!ScrollViewer.HorizontalScrollBarVisibilityProperty] = box[!ScrollViewer.HorizontalScrollBarVisibilityProperty],
            [!ScrollViewer.VerticalScrollBarVisibilityProperty] = box[!ScrollViewer.VerticalScrollBarVisibilityProperty],
            [!ScrollViewer.IsScrollChainingEnabledProperty] = box[!ScrollViewer.IsScrollChainingEnabledProperty],
            [!ScrollViewer.AllowAutoHideProperty] = box[!ScrollViewer.AllowAutoHideProperty],
            [!ScrollViewer.BringIntoViewOnFocusChangeProperty] = box[!ScrollViewer.BringIntoViewOnFocusChangeProperty],
            Content = new Panel()
            {
                Children = { presenter.PlateLayer, presenter }
            }
        }.RegisterInNameScope(scope);
        scrollViewer.Styles.Add(new AvaloniaStyle(x => x.OfType<ScrollContentPresenter>().Name("PART_ContentPresenter"))
        {
            Setters =
            {
                new AvaloniaSetter(CursorProperty, new Cursor(StandardCursorType.Ibeam))
            }
        });

        var text = new DockPanel()
        {
            Name = "PART_InnerDockPanel",
            [!MarginProperty] = box[!PaddingProperty],
            Children = { scrollViewer }
        }.RegisterInNameScope(scope);
        Grid.SetColumn(text, 1);

        var right = new ContentPresenter()
        {
            [!ContentPresenter.VerticalContentAlignmentProperty] = box[!VerticalContentAlignmentProperty],
            [!ContentPresenter.ContentProperty] = box[!InnerRightContentProperty]
        };
        Grid.SetColumn(right, 2);

        var border = new Border()
        {
            Name = "PART_BorderElement",
            [!Border.BackgroundProperty] = box[!BackgroundProperty],
            [!Border.BorderBrushProperty] = box[!BorderBrushProperty],
            [!Border.BorderThicknessProperty] = box[!BorderThicknessProperty],
            [!Border.CornerRadiusProperty] = box[!CornerRadiusProperty],
            [!MinWidthProperty] = box[!MinWidthProperty],
            [!MinHeightProperty] = box[!MinHeightProperty]
        }.RegisterInNameScope(scope);

        return new Panel()
        {
            Children =
            {
                border,
                new Border()
                {
                    [!MarginProperty] = box[!BorderThicknessProperty],
                    Child = new Grid()
                    {
                        ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto"),
                        Children =
                        {
                            new ContentPresenter()
                            {
                                [!ContentPresenter.VerticalContentAlignmentProperty] = box[!VerticalContentAlignmentProperty],
                                [!ContentPresenter.ContentProperty] = box[!InnerLeftContentProperty]
                            },
                            text,
                            right
                        }
                    }
                },
                new Popup() { Name = "PART_CompletionPopup" }.RegisterInNameScope(scope)
            }
        };
    });
}

/// <summary>
/// The text presenter of a <see cref="ColoredTextBox"/>: the text in the colors of its
/// pieces, the [...] parts of a label's text on plates (<see cref="PlateLayer"/>).
/// TextPresenter makes its layout through a virtual method, and a layout takes runs in other
/// brushes (Avalonia's own presenter underlines the text an input method is composing that
/// way), so the caret, the selection and a click go by the very layout the colors are drawn
/// in. While an input method composes, and for a password, it is the plain presenter.
/// </summary>
public class ColoredTextPresenter : TextPresenter
{
    readonly Dictionary<ExpressionClass, IBrush> brushes = new Dictionary<ExpressionClass, IBrush>();

    /// <summary>
    /// What draws the plates: a layer of the template under the presenter, so that the
    /// selection and the text go over them. The presenter can't draw under its text itself:
    /// Avalonia sealed the Render of every control of its own (2023, before 11.0), and a
    /// run's background, the other way, is the box of its glyphs, drawn over the selection.
    /// </summary>
    public Control PlateLayer { get; }

    public ColoredTextPresenter()
    {
        PlateLayer = new PartPlates(this);
        foreach (var kind in Enum.GetValues<ExpressionClass>())
        {
            this.ObserveTheme(ThemeKey(kind), color =>
            {
                brushes[kind] = new ImmutableSolidColorBrush(color);
                Recolor();
            });
        }
    }

    static string ThemeKey(ExpressionClass kind)
    {
        return kind switch
        {
            ExpressionClass.Part => nameof(AppTheme.ExpressionBackground),
            ExpressionClass.ErrorPart => nameof(AppTheme.ExpressionErrorBackground),
            ExpressionClass.Bracket => nameof(AppTheme.ExpressionBracket),
            ExpressionClass.Figure => nameof(AppTheme.ExpressionFigure),
            ExpressionClass.Property => nameof(AppTheme.ExpressionProperty),
            ExpressionClass.Function => nameof(AppTheme.ExpressionFunction),
            ExpressionClass.Number => nameof(AppTheme.ExpressionNumber),
            _ => nameof(AppTheme.ExpressionReserved)
        };
    }

    /// <summary>The text is laid out again, in the colors its box gives now</summary>
    public void Recolor()
    {
        InvalidateTextLayout();
    }

    /// <summary>
    /// The pieces of the text as laid out; none while an input method composes (the layout
    /// then holds text that isn't the box's) or for a password
    /// </summary>
    IReadOnlyList<ClassifiedSpan> Pieces
    {
        get
        {
            if (!string.IsNullOrEmpty(PreeditText) || (PasswordChar != default && !RevealPassword))
            {
                return [];
            }

            return (TemplatedParent as ColoredTextBox)?.PiecesOf(Text) ?? [];
        }
    }

    // The plates go where the text is: drawn again whenever the text is laid out anew or
    // moves. Not when its layout is made: the plate layer makes it, drawing, and a visual
    // invalidated during the render pass throws there, which stops the renderer for good.
    // (Null while the base's constructor runs.)
    protected override void InvalidateTextLayout()
    {
        base.InvalidateTextLayout();
        PlateLayer?.InvalidateVisual();
    }

    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (change.Property == BoundsProperty)
        {
            PlateLayer?.InvalidateVisual();
        }
    }

    protected override TextLayout CreateTextLayout()
    {
        var plain = base.CreateTextLayout();
        var pieces = Pieces;
        if (pieces.Count == 0)
        {
            return plain;
        }

        // The same layout with the runs. The room it is laid out in is the base's to say,
        // and kept to itself: the size the text was measured in, or the width it was
        // arranged at (its MeasureOverride and ArrangeOverride).
        double maxWidth = plain.MaxWidth;
        double maxHeight = plain.MaxHeight;
        plain.Dispose();
        var typeface = new Typeface(FontFamily, FontStyle, FontWeight, FontStretch);
        return new TextLayout(
            Text,
            typeface,
            FontSize,
            Foreground,
            TextAlignment,
            TextWrapping,
            textTrimming: null,
            textDecorations: null,
            FlowDirection,
            maxWidth,
            maxHeight,
            LineHeight,
            LetterSpacing,
            maxLines: 0,
            FontFeatures,
            Runs(pieces, typeface));
    }

    /// <summary>
    /// The runs: each character in the brush of its piece, and a selected one in the
    /// selection's own, as the base draws it. The layout takes one run at a time, so a run
    /// is a stretch of characters drawn alike.
    /// </summary>
    List<ValueSpan<TextRunProperties>> Runs(IReadOnlyList<ClassifiedSpan> pieces, Typeface typeface)
    {
        int length = Text?.Length ?? 0;
        var foregrounds = new IBrush[length];
        foreach (var piece in pieces.Where(piece => !IsPlate(piece.Class)))
        {
            var brush = brushes.GetValueOrDefault(piece.Class);
            for (int i = System.Math.Max(piece.Start, 0); i < System.Math.Min(piece.End, length); i++)
            {
                foregrounds[i] = brush;
            }
        }

        if (ShowSelectionHighlight && SelectionForegroundBrush != null)
        {
            int selectionEnd = System.Math.Min(System.Math.Max(SelectionStart, SelectionEnd), length);
            for (int i = System.Math.Max(System.Math.Min(SelectionStart, SelectionEnd), 0); i < selectionEnd; i++)
            {
                foregrounds[i] = SelectionForegroundBrush;
            }
        }

        var runs = new List<ValueSpan<TextRunProperties>>();
        int start = 0;
        for (int i = 1; i <= length; i++)
        {
            if (i < length && foregrounds[i] == foregrounds[start])
            {
                continue;
            }

            if (foregrounds[start] != null)
            {
                runs.Add(new ValueSpan<TextRunProperties>(
                    start,
                    i - start,
                    new GenericTextRunProperties(typeface, FontSize, foregroundBrush: foregrounds[start], fontFeatures: FontFeatures)));
            }

            start = i;
        }

        return runs;
    }

    /// <summary>Whether the class is drawn under its span (a part of a label's text) rather than in its characters</summary>
    static bool IsPlate(ExpressionClass kind)
    {
        return kind is ExpressionClass.Part or ExpressionClass.ErrorPart;
    }

    /// <summary>
    /// The plates of the parts, in the presenter's own coordinates: a rectangle per line a
    /// part is on, as the layout has the line
    /// </summary>
    public IEnumerable<(Rect Rect, IBrush Brush)> Plates()
    {
        foreach (var piece in Pieces.Where(piece => IsPlate(piece.Class)))
        {
            var brush = brushes.GetValueOrDefault(piece.Class);
            if (brush == null)
            {
                continue;
            }

            foreach (var rect in TextLayout.HitTestTextRange(piece.Start, piece.Length))
            {
                yield return (rect, brush);
            }
        }
    }
}

/// <summary>
/// The plates under the [...] parts of a label's text in a <see cref="ColoredTextBox"/>
/// (<see cref="ColoredTextPresenter.PlateLayer"/>), each a little wider than its part,
/// with round corners: beside its presenter in one panel, so the text and the selection
/// go over them and the scroll viewer moves them along
/// </summary>
public class PartPlates : Control
{
    const double Overhang = 1.5;
    const double CornerRadius = 3;

    readonly ColoredTextPresenter presenter;

    public PartPlates(ColoredTextPresenter presenter)
    {
        this.presenter = presenter;
        IsHitTestVisible = false;
    }

    public override void Render(DrawingContext context)
    {
        var offset = presenter.Bounds.Position - Bounds.Position;
        foreach (var (rect, brush) in presenter.Plates())
        {
            var plate = rect.Translate(offset).Inflate(new Thickness(Overhang, 0));
            context.DrawRectangle(brush, pen: null, plate, radiusX: CornerRadius, radiusY: CornerRadius);
        }
    }
}
