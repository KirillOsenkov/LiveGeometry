using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;

namespace DynamicGeometry
{
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

        protected override int DefaultZOrder()
        {
            return (int)ZOrder.Labels;
        }

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
        public virtual string Text
        {
            get
            {
                return text;
            }
            set
            {
                text = value ?? "";
                if (ShouldProcessText)
                {
                    ProcessText();
                }
                else
                {
                    ProcessedText = text;
                }
            }
        }

        protected bool ShouldProcessText = false;

        /// <summary>What an expression that has no value shows (the square root of a negative number, 0 / 0)</summary>
        public const string UndefinedText = "undefined";

        const string squareBracketsRegexString = @"[\[][^\[\]]*[\]]";
        static Regex squareBrackets = new Regex(squareBracketsRegexString, RegexOptions.None);

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

            var matches = squareBrackets.Matches(text);
            int chunkStart = 0;
            int chunkEnd = 0;
            string chunk;
            foreach (Match match in matches)
            {
                chunkEnd = match.Index;
                chunk = "";
                if (chunkEnd > chunkStart)
                {
                    chunk = text.Substring(chunkStart, chunkEnd - chunkStart);
                }
                textChunks.Add(chunk);
                ProcessMatch(match);
                chunkStart = match.Index + match.Length;
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
                        sb.Append(!value.IsValidValue() ? UndefinedText : Math.Round(value, DecimalsToShow).ToString());
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

            var renamed = squareBrackets.Replace(text, match => match.Value.Length < 3
                ? match.Value
                : "[" + renamer.Rewrite(match.Value.Substring(1, match.Value.Length - 2), isFunction: false) + "]");
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

        public IReadOnlyList<string> ExpressionTexts
        {
            get
            {
                return new[] { text };
            }
            set
            {
                if (value[0] != text)
                {
                    text = value[0];
                    RaisePropertyChanged("Text");
                }
            }
        }

        void ProcessMatch(Match match)
        {
            var result = match.Value;
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

        public virtual string ProcessedText
        {
            get
            {
                return TextBlock.Text;
            }
            set
            {
                TextBlock.Text = value;
            }
        }

        private int mDecimalsToShow = Settings.DisplayDecimals;
        [PropertyGridName("Decimals")]
        [Domain(0, 10)]
        [PropertyGridVisible]
        public virtual int DecimalsToShow
        {
            get
            {
                return mDecimalsToShow;
            }
            set
            {
                if (value >= 0 && value <= 10)
                {
                    mDecimalsToShow = value;
                    UpdateVisual();
                }
            }
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
}
