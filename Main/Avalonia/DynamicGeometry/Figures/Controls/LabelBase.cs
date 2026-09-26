using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using System.Xml.Linq;
using System.Xml;

namespace DynamicGeometry
{
    public abstract partial class LabelBase : ControlBase
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
        /// </summary>
        public Size MeasureSize()
        {
            Shape.InvalidateMeasure();
            Shape.Measure(Size.Infinity);
            return Shape.DesiredSize;
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
        [PropertyGridVisible]
        [PropertyGridFocus]
        public virtual string Text
        {
            get
            {
                return text;
            }
            set
            {
                text = value;
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

            if (!ShouldProcessText || text.IsEmpty())
            {
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
                        sb.Append(Math.Round(compileResult.Expression(), DecimalsToShow).ToString());
                        //sb.Append(compileResult.Expression().ToString("F01"));
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

            var compileResult = Drawing.CompileExpression(expression);
            embeddedExpressions.Add(compileResult);
            if (compileResult.IsSuccess)
            {
                Dependencies.AddRange(compileResult.Dependencies);
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
            writer.WriteAttributeDouble("DecimalsToShow", (double)DecimalsToShow);
        }

    }
}
