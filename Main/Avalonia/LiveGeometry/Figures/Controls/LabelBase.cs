using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Threading;

namespace LiveGeometry;

public abstract partial class LabelBase : ControlBase, IRenamableExpressions
{
    protected Border selection = new Border();
    public TextBlock TextBlock { get; set; }
    Brush selectionBrush;

    public LabelBase()
    {
        // SystemColors.HighlightColor has no Avalonia equivalent; classic Windows selection blue.
        Color back = Color.FromArgb(255, 51, 153, 255);
        selectionBrush = new SolidColorBrush(
            Color.FromArgb(50, back.R, back.G, back.B));
    }

    protected override ZOrder DefaultLayer()
    {
        return ZOrder.Labels;
    }

    /// <summary>The tint behind a selected label's text</summary>
    protected Color SelectionColor => ((SolidColorBrush)selectionBrush).Color;

    protected override string Kind => "Text";

    /// <summary>The start of the text, in quotes: Text “Drag the point…”</summary>
    public override string Construction => ConstructionText.Quote(Text);

    protected override FrameworkElement CreateShape()
    {
        TextBlock = Factory.CreateLabelShape();
        selection.Child = TextBlock;
        return selection;
    }

    /// <summary>
    /// The size the text takes right now - laid out here and now, because Bounds are stale
    /// right after the text changed or the label was created. The shape is a Border around
    /// the TextBlock, and its own measure stays valid when the text or the width inside it
    /// changed, so it is invalidated first: without that it answered with the old size.
    /// A hidden label is measured as it will be when shown (Avalonia measures a hidden
    /// control as nothing): a gallery caption leaves room for the hint a box brings up.
    /// </summary>
    public Size MeasureSize()
    {
        bool wasVisible = Shape.IsVisible;
        Shape.IsVisible = true;
        Shape.InvalidateMeasure();
        Shape.Measure(Size.Infinity);
        var size = Shape.DesiredSize;
        Shape.IsVisible = wasVisible;
        return size;
    }

    public override void ApplyStyle()
    {
        if (this.Style == null)
        {
            return;
        }

        this.Apply(TextBlock, Style);
        selection.Background = Selected ? selectionBrush : null;
        if (Drawing != null)
        {
            UpdateVisual();
        }
    }

    protected string text;
    // not [PropertyGridFocus]: only a new label takes the keyboard (LabelCreator); one that
    // is merely selected (on the canvas, in the Figure List) would steal the keys from there
    [PropertyGridVisible]
    [PropertyGridMultiline]
    [PropertyGridDisallowMultiEdit]
    [PropertyGridPreferredEditor("LabelText")]
    public virtual string Text
    {
        get => text;
        set
        {
            bool showedDecimals = ShowsDecimals();
            string old = text;
            text = value ?? "";

            // the Decimals row comes and goes with the text's [...] parts, as they are typed
            if (ShowsDecimals() != showedDecimals)
            {
                RaisePropertyChanged(nameof(DecimalsToShow));
            }

            if (ShouldProcessText)
            {
                ProcessText();

                // What is built on the label's value follows it (a circle with it for its
                // radius, a point at its value): edited, or undone, the label showed the
                // new number while they stayed where they were until the next
                // recalculation. Not while a file is read, nor for a label not in the
                // drawing yet; loading and re-binding don't come through here.
                if (Drawing != null && !Drawing.IsReading && Drawing.Figures.Contains(this))
                {
                    this.RecalculateAllDependents();
                }
            }
            else
            {
                ProcessedText = text;
            }

            // A text typed follows itself where it is quoted: the side panel's title, as it
            // is typed (it stayed as it was until something else changed). Not the text of a
            // measurement or a point's name, which is worked out again at every move.
            if (ShouldProcessText && text != old)
            {
                RaisePropertyChanged(nameof(Text));
            }
        }
    }

    protected bool ShouldProcessText = false;

    /// <summary>What an expression that has no value shows (the square root of a negative number, 0 / 0)</summary>
    public const string UndefinedText = "undefined";

    /// <summary>
    /// The [...] parts of a text, each with its brackets: from a [ to the next ], over a
    /// name in backticks whole (`a]b`, whose ] ends nothing). A [ before the ] starts the
    /// part over ("[[A.X]" is a [ of the text, then [A.X]). A backtick that is never
    /// closed is an ordinary character, for the expression to say what is wrong.
    /// </summary>
    public static List<(int Start, int Length)> FindExpressionParts(string text)
    {
        var parts = new List<(int Start, int Length)>();
        int start = -1;
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (c == '[')
            {
                start = i;
            }
            else if (c == ']' && start >= 0)
            {
                parts.Add((start, i - start + 1));
                start = -1;
            }
            else if (c == '`' && start >= 0)
            {
                int close = ClosingBacktick(text, i);
                if (close >= 0)
                {
                    i = close;
                }
            }
        }

        return parts;
    }

    /// <summary>Where the name in backticks opened at <paramref name="open"/> closes, a doubled backtick being one of its characters (as <see cref="Scanner"/> reads it); -1 if it doesn't</summary>
    static int ClosingBacktick(string text, int open)
    {
        for (int i = open + 1; i < text.Length; i++)
        {
            if (text[i] != '`')
            {
                continue;
            }

            if (i + 1 < text.Length && text[i + 1] == '`')
            {
                i++;
                continue;
            }

            return i;
        }

        return -1;
    }

    protected List<string> textChunks = null;
    protected List<CompileResult> embeddedExpressions = null;

    protected void ProcessText()
    {
        if (!ShouldProcessText)
        {
            return;
        }

        string text = this.text;
        this.UnregisterFromDependencies();

        Dependencies.Clear();
        embeddedExpressions = new List<CompileResult>();
        textChunks = new List<string>();

        int chunkStart = 0;
        int chunkEnd = 0;
        string chunk;
        foreach (var part in FindExpressionParts(text))
        {
            chunkEnd = part.Start;
            chunk = "";
            if (chunkEnd > chunkStart)
            {
                chunk = text.Substring(chunkStart, chunkEnd - chunkStart);
            }
            textChunks.Add(chunk);
            ProcessPart(text.Substring(part.Start, part.Length));
            chunkStart = part.Start + part.Length;
        }
        chunkEnd = text.Length;
        chunk = "";
        if (chunkEnd > chunkStart)
        {
            chunk = text.Substring(chunkStart, chunkEnd - chunkStart);
        }
        textChunks.Add(chunk);

        this.RegisterWithDependencies();

        Recalculate();
    }

    public override void Recalculate()
    {
        if (Settings.ScaleTextWithDrawing)
        {
            var s = Drawing.CoordinateSystem.Scale;
            ScaleTransform scale = new ScaleTransform();
            scale.ScaleX = s;
            scale.ScaleY = s;
            Shape.RenderTransform = scale;
        }

        if (!ShouldProcessText)
        {
            return;
        }

        if (text.IsEmpty())
        {
            ProcessedText = "";
            return;
        }

        if (textChunks == null || embeddedExpressions == null)
        {
            ProcessText();
        }

        StringBuilder sb = new StringBuilder();

        for (int i = 0; i < textChunks.Count; i++)
        {
            if (i != 0)
            {
                var compileResult = embeddedExpressions[i - 1];
                if (compileResult.IsSuccess)
                {
                    // (the square root of a negative number, 0 / 0: it said "NaN"; 1 / 0
                    // said "Infinity")
                    double value = compileResult.Expression();
                    sb.Append(!value.IsValidValue() ? UndefinedText : FormatNumber(value));
                }
                else
                {
                    sb.Append(compileResult.ToString());
                }
            }
            sb.Append(textChunks[i]);
        }

        ProcessedText = sb.ToString();
    }

    /// <summary>
    /// A number of a text label's [...] part, with every decimal it shows, zeros at the end
    /// included (3.70, 9.00): a number that lost its last zero now and then got shorter,
    /// and the text after it jumped to and fro as a figure was dragged
    /// </summary>
    string FormatNumber(double value)
    {
        return Math.Round(value, DecimalsToShow).ToString("F" + DecimalsToShow, CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// The names in the [...] parts follow renamed figures. Only the text changes: the
    /// compiled parts already hold the figures, and recompiling here, in the middle of a
    /// rename, would re-register the dependencies being walked.
    /// </summary>
    public void RenameInExpressions(ExpressionRenamer renamer)
    {
        if (!ShouldProcessText || text.IsEmpty())
        {
            return;
        }

        var sb = new StringBuilder();
        int chunkStart = 0;
        foreach (var part in FindExpressionParts(text))
        {
            sb.Append(text, chunkStart, part.Start - chunkStart);
            sb.Append(part.Length < 3
                ? text.Substring(part.Start, part.Length)
                : "[" + renamer.Rewrite(text.Substring(part.Start + 1, part.Length - 2), isFunction: false) + "]");
            chunkStart = part.Start + part.Length;
        }

        sb.Append(text, chunkStart, text.Length - chunkStart);
        var renamed = sb.ToString();
        if (renamed != text)
        {
            text = renamed;
            RaisePropertyChanged("Text");
        }
    }

    public void RebindExpressions()
    {
        if (ShouldProcessText && !text.IsEmpty())
        {
            ProcessText();
        }
    }

    /// <summary>The expressions of the [...] parts, without the brackets: none in a text that is shown as it is</summary>
    public IEnumerable<string> Expressions
    {
        get
        {
            if (!ShouldProcessText || text.IsEmpty())
            {
                return [];
            }

            return FindExpressionParts(text)
                .Where(part => part.Length >= 3)
                .Select(part => text.Substring(part.Start + 1, part.Length - 2));
        }
    }

    public IReadOnlyList<string> ExpressionTexts
    {
        get => new[] { text };
        set
        {
            if (value[0] != text)
            {
                text = value[0];
                RaisePropertyChanged("Text");
            }
        }
    }

    void ProcessPart(string result)
    {
        if (result.Length < 3)
        {
            CompileResult error = new CompileResult();
            error.AddError("Empty expression");
            embeddedExpressions.Add(error);
            return;
        }

        var expression = result.Substring(1, result.Length - 2);

        var compileResult = Compiler.Instance.CompileExpression(Drawing, expression, figure => !figure.DependsOn(this));
        embeddedExpressions.Add(compileResult);
        if (compileResult.IsSuccess)
        {
            // once each: a figure named in two [...] parts listed twice would be swapped
            // only in its first place when it is replaced (ReplaceDependency)
            Dependencies.Merge(compileResult.Dependencies);
        }
    }

    /// <summary>
    /// The value of a label that is one expression and nothing else ("[AB / 3]"), as
    /// it is and not as the label shows it: rounded to the label's decimals, a radius
    /// taken from it was 0.33 instead of a third, and changed with Decimals. Null for
    /// any other label.
    /// </summary>
    protected double? ExactValue
    {
        get
        {
            if (!ShouldProcessText
                || textChunks == null
                || embeddedExpressions == null
                || embeddedExpressions.Count != 1
                || textChunks.Count != 2
                || !string.IsNullOrWhiteSpace(textChunks[0])
                || !string.IsNullOrWhiteSpace(textChunks[1])
                || !embeddedExpressions[0].IsSuccess)
            {
                return null;
            }

            double value = embeddedExpressions[0].Expression();
            return value.IsValidValue() ? value : double.NaN;
        }
    }

    /// <summary>
    /// How often a label shows a new text while figures move. A number that changed at
    /// every move of a drag changed its width with it, and the lines after it jumped to and
    /// fro; and each change laid the whole text out again.
    /// </summary>
    public static readonly TimeSpan TextInterval = TimeSpan.FromMilliseconds(100);

    // on screen: a pin is measured from the canvas, and a text shown later is laid out there
    protected bool HasCanvas => Drawing != null && Drawing.Canvas != null;

    string processedText;
    long textShownAt;
    IDisposable pendingText;

    /// <summary>
    /// The text as worked out from the expressions or the measure, right now. The label
    /// shows it at once, except while figures move (<see cref="Drawing.IsMoving"/>): then
    /// at once only if the text on screen has been there for <see cref="TextInterval"/>,
    /// otherwise when that time is up - once an interval while the numbers change, and the
    /// last text in the end.
    /// </summary>
    public virtual string ProcessedText
    {
        get => processedText;
        set
        {
            processedText = value;
            if (Drawing != null && Drawing.IsMoving && HasCanvas)
            {
                ShowTextSoon();
            }
            else
            {
                ShowText();
            }
        }
    }

    /// <summary>
    /// What the label shows for a text: the text itself here; a wrapped <see cref="Label"/>
    /// keeps its expressions on one line (<see cref="LineBreaks"/>)
    /// </summary>
    protected virtual string ShownText(string text) => text;

    /// <summary>The text shown again, after something <see cref="ShownText"/> goes by has changed</summary>
    protected void ShowTextAgain()
    {
        if (processedText != null)
        {
            ShowText();
        }
    }

    void ShowText()
    {
        pendingText?.Dispose();
        pendingText = null;
        var shown = ShownText(processedText);
        if (TextBlock.Text != shown)
        {
            TextBlock.Text = shown;
            textShownAt = Stopwatch.GetTimestamp();
        }
    }

    void ShowTextSoon()
    {
        var shownFor = Stopwatch.GetElapsedTime(textShownAt);
        if (TextBlock.Text == ShownText(processedText) || shownFor >= TextInterval)
        {
            ShowText();
            return;
        }

        // (above Input: a drag that keeps the dispatcher busy with moves still shows a
        // number once an interval)
        pendingText ??= DispatcherTimer.RunOnce(ShowPendingText, TextInterval - shownFor, DispatcherPriority.Normal);
    }

    void ShowPendingText()
    {
        pendingText = null;
        var shown = ShownText(processedText);
        if (TextBlock.Text == shown)
        {
            return;
        }

        TextBlock.Text = shown;
        textShownAt = Stopwatch.GetTimestamp();

        // a label placed by its size (a pinned caption, a point's name kept clear of its
        // point) goes where the new size puts it
        if (Drawing != null && HasCanvas)
        {
            UpdateVisual();
        }
    }

    private int mDecimalsToShow = Settings.DisplayDecimals;
    [PropertyGridName("Decimals")]
    [Domain(0, 10)]
    [PropertyGridVisible]
    [PropertyGridCondition(nameof(ShowsDecimals))]
    public virtual int DecimalsToShow
    {
        get => mDecimalsToShow;
        set
        {
            if (value >= 0 && value <= 10)
            {
                mDecimalsToShow = value;
                UpdateVisual();
            }
        }
    }

    /// <summary>Whether the label shows a number, for which Decimals means something: a measurement does</summary>
    public virtual bool ShowsDecimals()
    {
        return true;
    }

    public override void ReadXml(XElement element)
    {
        base.ReadXml(element);
        // files from before the attribute existed: the default, not 0
        if (element.Attribute("DecimalsToShow") != null)
        {
            DecimalsToShow = (int)element.ReadDouble("DecimalsToShow");
        }
    }

    public override void WriteXml(XmlWriter writer)
    {
        base.WriteXml(writer);
        if (DecimalsToShow != Settings.DisplayDecimals)
        {
            writer.WriteAttributeDouble("DecimalsToShow", (double)DecimalsToShow);
        }
    }

}
