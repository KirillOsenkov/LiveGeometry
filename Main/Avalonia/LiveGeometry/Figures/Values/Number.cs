using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Linq;
using Avalonia;

namespace LiveGeometry;

/// <summary>
/// A number in the drawing: a figure other figures depend on the way they depend on a
/// segment's length, usable in expressions by name - a, b, c as a slider is named (n1, n2
/// for the helpers tools make). Its value is typed, or worked out by a formula over the
/// drawing (<see cref="Expression"/>: the sum of two arc lengths), and on the paper it is a
/// caption "a = 18.85" like a slider's, which a drag moves (the Number tool,
/// <see cref="NumberCreator"/>). Hidden by default (<see cref="VisibleByDefault"/>): the
/// helpers stay off the paper - a typed distance of a translation, a fixed length's number -
/// and the tool shows the ones it makes; hidden, it is a variable of the Figure List.
/// Provides itself as a length (in units) and as an angle (in degrees). One created on
/// demand for a single figure is <see cref="FigureBase.Auxiliary"/> and leaves the drawing
/// with its last user.
/// </summary>
public class Number : CompositeFigure, INumber, ILengthProvider, IAngleProvider, IMovableParts, IRenamableExpressions, IExpressionOwner
{
    readonly WholeHandle wholeHandle;

    public Number()
    {
        Expression = new DrawingExpression(this) { Name = FormulaName };
        Expression.ValueChanged += Expression_ValueChanged;
        Caption = new NumberCaption(this) { Name = "number caption" };
        AddChild(Caption);
        wholeHandle = new WholeHandle(this);

        // the caption is a label: among the labels, where Bring to front and Send to back apply
        Layer = ZOrder.Labels;
        Visible = false;
    }

    /// <summary>The name of the Formula row, which the Number tool puts the keyboard into</summary>
    public const string FormulaName = "Formula";

    /// <summary>"a = 18.85" on the paper</summary>
    public NumberCaption Caption { get; }

    public static Number CreateAuxiliary(Drawing drawing, double value)
    {
        return new Number() { Drawing = drawing, Auxiliary = true, Value = value };
    }

    #region Value and formula

    double value;
    bool placed;

    /// <summary>
    /// The formula the number is worked out by, over the drawing ("d.ArcLength + f.ArcLength",
    /// "2 * a"); empty for a typed number. Compiled as a point by coordinates' X is.
    /// </summary>
    [PropertyGridVisible]
    [PropertyGridName(FormulaName)]
    public DrawingExpression Expression { get; }

    /// <summary>Whether the value is typed rather than worked out by a formula</summary>
    public bool IsTyped => string.IsNullOrEmpty(Expression.Text);

    /// <summary>
    /// The number. Typed into a number that has a formula, it takes the formula's place:
    /// the row is a plain number again (undo brings the formula back: <see cref="ValueState"/>).
    /// </summary>
    [PropertyGridVisible]
    [PropertyGridPreferredEditor("UpDown")]
    [PropertyGridCustomValueProvider(typeof(ValueState))]
    public double Value
    {
        get => value;
        set
        {
            this.value = value;
            if (!IsTyped)
            {
                Expression.Text = "";
                FormulaChanged();
                return;
            }

            RaisePropertyChanged(nameof(Value));
            RefreshCaption();
            if (Drawing != null && !Dependents.IsEmpty())
            {
                this.RecalculateAllDependents();
            }
        }
    }

    /// <summary>The caption says the number anew, where it is on the paper; a hidden one waits until it shows</summary>
    void RefreshCaption()
    {
        if (Drawing != null && Visible)
        {
            Caption.UpdateVisual();
        }
    }

    /// <summary>
    /// The Value row as the grid sets it: undo puts back the formula a typed value replaced,
    /// not only the number it gave
    /// </summary>
    public class ValueState : PropertyValue, IRestorableValue
    {
        record SavedValue(double Value, string Formula);

        Number Owner => (Number)Parent;

        public object CaptureState()
        {
            return new SavedValue(Owner.value, Owner.Expression.Text);
        }

        public void RestoreState(object state)
        {
            var saved = (SavedValue)state;
            var owner = Owner;
            if (string.IsNullOrEmpty(saved.Formula))
            {
                owner.Value = saved.Value;
                return;
            }

            owner.Expression.Text = saved.Formula;
            owner.Expression.Rebind();
            owner.FormulaChanged();
        }
    }

    /// <summary>The Formula row was set (the grid, undo): the number follows it, or is typed again when it was cleared</summary>
    void Expression_ValueChanged()
    {
        FormulaChanged();
    }

    void FormulaChanged()
    {
        if (IsTyped && !Dependencies.IsEmpty())
        {
            // a typed number again, at the value the formula last gave: built on nothing
            this.UnregisterFromDependencies();
            Dependencies.SetItems(new List<IFigure>());
        }

        if (Drawing != null && Drawing.Figures.Contains(this))
        {
            this.RecalculateAllDependents();
        }

        RefreshCaption();
        RaisePropertyChanged(nameof(Value));
        RaiseConstructionChanged();
    }

    public IEnumerable<DrawingExpression> Expressions => new[] { Expression };

    public void RenameInExpressions(ExpressionRenamer renamer)
    {
        Expression.RenameInExpression(renamer);
        RaiseConstructionChanged();
    }

    public void RebindExpressions()
    {
        Expression.Rebind();
    }

    public IReadOnlyList<string> ExpressionTexts
    {
        get => new[] { Expression.Text };
        set
        {
            Expression.Text = value[0];
            RaiseConstructionChanged();
        }
    }

    public override void Recalculate()
    {
        if (!IsTyped && Drawing != null)
        {
            // a formula that doesn't compile (it names a figure that isn't there) gives no
            // number: the figure doesn't exist, nor does anything built on it; one that
            // has no value right now (the root of a negative number) says "undefined"
            var worked = Expression.Value;
            if (worked == null)
            {
                Exists = false;
                UpdateCaptionExistence();
                return;
            }

            value = worked();
        }

        Exists = Dependencies.Exists();
        UpdateCaptionExistence();
        base.Recalculate();
    }

    public override void UpdateExistence()
    {
        base.UpdateExistence();
        if (Exists && !IsTyped && Expression.Value == null)
        {
            Exists = false;
        }

        UpdateCaptionExistence();
    }

    void UpdateCaptionExistence()
    {
        Caption.Exists = Exists;
    }

    public override void OnAddingToDrawing(Drawing drawing)
    {
        // compiles the formula, so that the number and its dependencies are right at once
        Recalculate();
        base.OnAddingToDrawing(drawing);
    }

    public double Length => Value;

    /// <summary>Angle providers speak radians; the number itself is in degrees</summary>
    public double Angle => Value.ToRadians();

    [PropertyGridVisible]
    [Domain(0, 10)]
    public int Decimals
    {
        get => Caption.DecimalsToShow; set => Caption.DecimalsToShow = value;
    }

    bool showName = true;

    /// <summary>Whether the caption says the name: "a = 18.85", or "18.85" alone</summary>
    [PropertyGridVisible]
    [PropertyGridName("Show name")]
    public bool ShowName
    {
        get => showName;
        set
        {
            showName = value;
            RefreshCaption();
        }
    }

    /// <summary>What the caption says: the name and the number, or "undefined" for a formula without a value right now</summary>
    public string CaptionText
    {
        get
        {
            string number = value.IsValidValue() ? Math.Round(value, Decimals).ToString() : LabelBase.UndefinedText;
            return showName ? NameDisplay.Format(Name) + " = " + number : number;
        }
    }

    #endregion

    #region Place

    Point position;

    /// <summary>Where the caption's top-left corner sits in the plane</summary>
    public Point Position
    {
        get => position;
        set
        {
            position = value;
            placed = true;
            RefreshCaption();
        }
    }

    public override Point Center => Position;

    /// <summary>A number is a hidden helper until it is shown</summary>
    protected override bool VisibleByDefault => false;

    /// <summary>Shown for the first time without a place of its own, a number goes to the middle of the window</summary>
    public override bool Visible
    {
        get => base.Visible;
        set
        {
            if (value && !placed && Drawing != null && Drawing.Canvas != null)
            {
                Position = Drawing.CoordinateSystem.ViewCenter;
            }

            base.Visible = value;
            if (value && Drawing != null)
            {
                UpdateVisual();
            }
        }
    }

    public IMovable FindMovablePart(Point point)
    {
        return wholeHandle;
    }

    public IMovable WholePart => wholeHandle;

    /// <summary>
    /// Dragging the caption moves the number by as much as the cursor, the text keeping
    /// its grip under it; undo puts it back where it was
    /// </summary>
    class WholeHandle : IMovable, IRestorablePlace
    {
        readonly Number number;

        public WholeHandle(Number number)
        {
            this.number = number;
        }

        public Point Coordinates => number.Position;

        public bool AllowMove()
        {
            return !number.Locked;
        }

        public void MoveTo(Point position)
        {
            number.Position = position;
        }

        public object CapturePlace()
        {
            return number.Position;
        }

        public void RestorePlace(object place)
        {
            number.Position = (Point)place;
        }
    }

    /// <summary>A hit on the caption is a hit on the number: the caption is never handed out</summary>
    public override IFigure HitTest(Point point, Predicate<IFigure> filter)
    {
        return Visible && Caption.HitTest(point) != null && filter(this) ? this : null;
    }

    #endregion

    #region Name and words

    // a, b, c: what an expression says, next to points A, B, C, from the pool sliders are
    // named from; the helpers tools make stay n1, n2, n3 - short, since they are what a
    // "tied to" row shows
    public override string GenerateFigureName(List<string> blacklist)
    {
        if (!Auxiliary)
        {
            return GenerateLetterName(this, "a", blacklist);
        }

        for (int i = 1; ; i++)
        {
            var candidate = "n" + i;
            if (this.NameAvailable(candidate) && (blacklist == null || !blacklist.Contains(candidate)))
            {
                return candidate;
            }
        }
    }

    /// <summary>The caption shows the name</summary>
    [PropertyGridVisible]
    [PropertyGridDisallowMultiEdit]
    public override string Name
    {
        get => base.Name;
        set
        {
            base.Name = value;
            RefreshCaption();
        }
    }

    /// <summary>A composite dumps its parts by default</summary>
    public override string ToString()
    {
        return Name;
    }

    protected override string Kind => "Number";

    /// <summary>A number goes by its name: "radius n1"</summary>
    public override string Noun => null;

    /// <summary>"= 3", or "= d.ArcLength + f.ArcLength" for a formula</summary>
    public override string Construction => "= " + (IsTyped ? ConstructionText.Number(Value) : Expression.Text);

    #endregion

    #region Style

    /// <summary>The caption's text style stands for the number's</summary>
    public override IFigureStyle Style
    {
        get => Caption.Style; set => Caption.Style = value;
    }

    #endregion

    #region The caption

    /// <summary>
    /// "a = 18.85" with its top-left corner at the number's place. Not draggable on its own:
    /// a drag moves the number.
    /// </summary>
    public class NumberCaption : LabelWithOffset
    {
        readonly Number number;

        public NumberCaption(Number number)
        {
            this.number = number;
        }

        public override Point Anchor => number.Position;

        protected override bool HasAnchor => true;

        public override bool AllowMove()
        {
            return false;
        }

        public override void UpdateVisual()
        {
            if (Drawing == null)
            {
                return;
            }

            Text = number.CaptionText;
            base.UpdateVisual();
        }
    }

    #endregion

    #region Serialization

    public override void WriteXml(XmlWriter writer)
    {
        base.WriteXml(writer);
        if (IsTyped)
        {
            writer.WriteAttributeDouble("Value", Value);
        }
        else
        {
            writer.WriteAttributeString(FormulaName, Expression.Text);
        }

        if (placed)
        {
            writer.WriteAttributeDouble("X", position.X);
            writer.WriteAttributeDouble("Y", position.Y);
        }

        if (!showName)
        {
            writer.WriteAttributeBool("ShowName", false);
        }

        if (Decimals != Settings.DisplayDecimals)
        {
            writer.WriteAttributeDouble("Decimals", Decimals);
        }
    }

    public override void ReadXml(XElement element)
    {
        base.ReadXml(element);
        if (element.Attribute("Decimals") != null)
        {
            Decimals = (int)element.ReadDouble("Decimals");
        }

        showName = element.ReadBool("ShowName", true);
        if (element.Attribute("X") != null)
        {
            Position = new Point(element.ReadDouble("X"), element.ReadDouble("Y"));
        }

        var formula = element.ReadString(FormulaName);
        if (!string.IsNullOrEmpty(formula))
        {
            Expression.Text = formula;
        }
        else
        {
            Value = element.ReadDouble("Value");
        }
    }

    #endregion
}
