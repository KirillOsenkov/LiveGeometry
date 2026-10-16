using System.ComponentModel;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;

namespace LiveGeometry;

/// <summary>
/// One click: where the number goes. The number is in the drawing from then on (adding it
/// is the undo step), and the tool's own panel opens with its Formula box and OK
/// (<see cref="NumberPanel"/>): a value (3) or a formula (d.ArcLength + f.ArcLength) typed
/// there goes into the number as it is typed; a click on a figure, or on its row in the
/// Figure List, puts a reference into the box while it has the keyboard
/// (<see cref="ReferencePicker"/>). OK, Enter, Escape or a right click are "done": the tool
/// goes down and the panel with it. A click on the paper meanwhile places the next number.
/// </summary>
[Category(BehaviorCategories.Measure)]
[Order(4)]
public class NumberCreator : Behavior
{
    // the number placed by the last click, whose panel is up
    Number editing;
    NumberPanel panel;

    public override void MouseDown(object sender, MouseButtonEventArgs e)
    {
        var drawing = Drawing;
        var number = new Number() { Drawing = drawing };
        number.Position = Coordinates(e);
        number.Visible = true;
        Actions.Add(drawing, number);
        editing = number;
        panel = new NumberPanel(this, number);
        drawing.RaiseDisplayProperties(panel);
    }

    /// <summary>The panel of the number just placed, while it is in the drawing (undo takes it out, and the panel with it)</summary>
    public override object PropertyBag
    {
        get
        {
            return editing != null && Drawing != null && Drawing.Figures.Contains(editing) ? panel : null;
        }
    }

    public override void Stopping()
    {
        editing = null;
        panel = null;
    }

    /// <summary>Done with the number: the tool goes down, and its panel closes with it</summary>
    public void FinishEditing()
    {
        AbortAndSetDefaultTool();
    }

    /// <summary>A cross: a new number appears here</summary>
    protected override Cursor GetCursor(Point coordinates)
    {
        return CrossCursor;
    }

    public override string Name => "Number";

    public override string HintText => "Click where the number goes, then type a value or a formula. Click a figure, or its row in the Figure List, to put it into the formula.";

    public override FrameworkElement CreateIcon()
    {
        var text = new TextBlock()
        {
            // as the Text tool's "Abc": a caption as the number shows it, upright, in ink
            Text = "a=2",
            FontSize = 13,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
        text.BindTheme(TextBlock.ForegroundProperty, nameof(AppTheme.Ink));
        var grid = new Grid()
        {
            MinWidth = IconBuilder.IconSize,
            MinHeight = IconBuilder.IconSize,
        };
        grid.Children.Add(text);
        return grid;
    }

    /// <summary>
    /// The tool's panel for the number just placed: its Formula box, which the keyboard
    /// goes into, and OK. What is typed goes into the number as soon as it compiles (the
    /// caption follows), and an empty box makes it a typed number again; what doesn't
    /// compile is said under the box when OK is pressed. Not an undo step of its own: the
    /// number with its formula is one, its placing. The panel says whose expression the box
    /// holds (<see cref="IExpressionHost"/>), so a reference to the number itself is refused.
    /// </summary>
    public class NumberPanel : ToolPanel, IExpressionHost
    {
        readonly NumberCreator tool;
        readonly Number number;
        string formula;

        public NumberPanel(NumberCreator tool, Number number)
        {
            this.tool = tool;
            this.number = number;
            formula = number.Expression.Text ?? "";
        }

        [PropertyGridVisible]
        [PropertyGridFocus]
        [PropertyGridEvent("KeyDown", "Formula_KeyDown")]
        [PropertyGridName(Number.FormulaName)]
        [PropertyGridPreferredEditor("Expression")]
        public string Formula
        {
            get => formula;
            set
            {
                formula = value ?? "";
                Apply();
            }
        }

        /// <summary>The number follows the box: a formula that compiles, or none for an empty box</summary>
        void Apply()
        {
            string text = formula.Trim();
            if (text.Length == 0)
            {
                if (!number.IsTyped)
                {
                    number.SetFormula("");
                }

                return;
            }

            var result = Compiler.Instance.CompileExpression(number.Drawing, text, figure => !figure.DependsOn(number));
            if (result.IsSuccess && text != number.Expression.Text)
            {
                number.SetFormula(text);
            }
        }

        internal void Formula_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                OK();
                e.Handled = true;
            }
        }

        /// <summary>Done: the tool goes down with the panel; a formula that doesn't compile keeps the panel and is said under the box</summary>
        [PropertyGridVisible]
        [PropertyGridIcon(PropertyGridIcon.Check)]
        public void OK()
        {
            string text = formula.Trim();
            if (text.Length > 0)
            {
                var result = Compiler.Instance.CompileExpression(number.Drawing, text, figure => !figure.DependsOn(number));
                ReportError(nameof(Formula), result.GetErrorText());
                if (!result.IsSuccess)
                {
                    return;
                }
            }

            tool.FinishEditing();
        }

        public Drawing Drawing => number.Drawing;

        public IFigure Owner => number;

        public InsertContext Context => InsertContext.Expression;

        /// <summary>The panel's title: "Number a"</summary>
        public override string ToString()
        {
            return number.Title;
        }
    }
}
