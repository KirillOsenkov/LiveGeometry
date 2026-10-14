using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace DynamicGeometry;

/// <summary>
/// The side panel right after a figure with a length or a direction is made: its length with
/// Fix length or Free length (a segment, a vector, a circle), and its direction with Fix
/// direction, Horizontal and Vertical (a segment, a vector, a ray, a line through two
/// points), forwarded to the figure - so that "a segment of length 2" is draw, type 2, click
/// Fix, and a horizontal one is draw, click Horizontal, type 2, click Fix length, without a
/// box on every tool. Captions come from the figure (Radius, Fix radius); the title is the
/// figure's ("Segment AB"). The figure's own grid has the same rows. A vector's panel ends
/// with its X and Y. (A regular polygon has a panel of its own, <see cref="RegularPolygonPanel"/>.)
/// </summary>
public class DimensionsPanel :
    IConditionalProperties,
    ICustomPropertyProvider,
    ICustomMethodProvider,
    IPropertyGridHost,
    INotifyPropertyChanged
{
    readonly IFigure figure;
    readonly IConditionalProperties conditions;
    readonly IFixableLength length;
    readonly IFixableDirection direction;
    PropertyGrid propertyGrid;

    /// <param name="figure">An <see cref="IFixableLength"/>, an <see cref="IFixableDirection"/>, or both</param>
    public DimensionsPanel(IFigure figure)
    {
        this.figure = figure;
        conditions = (IConditionalProperties)figure;
        length = figure as IFixableLength;
        direction = figure as IFixableDirection;
    }

    public event PropertyChangedEventHandler PropertyChanged;

    /// <summary>
    /// Set by the grid while it shows the panel: the figure's own changes are passed on as
    /// the panel's for that long.
    /// </summary>
    public PropertyGrid PropertyGrid
    {
        get
        {
            return propertyGrid;
        }
        set
        {
            if (propertyGrid == value)
            {
                return;
            }

            if (figure is INotifyPropertyChanged notifying)
            {
                if (value != null)
                {
                    notifying.PropertyChanged += Figure_PropertyChanged;
                }
                else
                {
                    notifying.PropertyChanged -= Figure_PropertyChanged;
                }
            }

            propertyGrid = value;
        }
    }

    void Figure_PropertyChanged(object sender, PropertyChangedEventArgs e)
    {
        PropertyChanged?.Invoke(this, e);
    }

    /// <summary>Whether the panel has the direction: where the figure has one that can be set</summary>
    bool ShowsDirection => direction != null && direction.CanEdit(nameof(Direction));

    /// <summary>
    /// The rows: the length, then Show where there is a figure to hang a measurement on, then
    /// the direction, then a vector's X and Y (its own rows: "a vector (3, 4)" is draw, type
    /// 3 and 4).
    /// </summary>
    public IEnumerable<IValueProvider> GetProperties()
    {
        if (length != null)
        {
            yield return PropertyDiscoveryStrategy.CreateValueProvider(this, nameof(Length));
            // (a figure with nothing in the drawing to measure has no Show)
            if (length.MeasuredFigures != null)
            {
                yield return PropertyDiscoveryStrategy.CreateValueProvider(this, nameof(Show));
            }
        }

        if (ShowsDirection)
        {
            yield return PropertyDiscoveryStrategy.CreateValueProvider(this, nameof(Direction));
        }

        if (figure is Vector vector && vector.CanEdit(nameof(Vector.X)))
        {
            yield return PropertyDiscoveryStrategy.CreateValueProvider(vector, nameof(Vector.X));
            yield return PropertyDiscoveryStrategy.CreateValueProvider(vector, nameof(Vector.Y));
        }
    }

    [PropertyGridVisible]
    [PropertyGridGroup("Length")]
    [PropertyGridPreferredEditor("UpDown")]
    [PropertyGridCustomValueProvider(typeof(PanelLengthValue))]
    public double Length
    {
        get { return length.Length; }
        set { length.Length = value; }
    }

    public class PanelLengthValue : LengthPropertyValue
    {
        public override IFixableLength Figure => ((DimensionsPanel)Parent).length;
    }

    /// <summary>The figure's direction, in degrees (named as the figure's row is, so that its CanEdit answers for both)</summary>
    [PropertyGridVisible]
    [PropertyGridGroup("Direction")]
    [PropertyGridPreferredEditor("UpDown")]
    [PropertyGridCustomValueProvider(typeof(PanelDirectionValue))]
    public double Direction
    {
        get { return direction.Direction; }
        set { direction.Direction = value; }
    }

    public class PanelDirectionValue : DirectionPropertyValue
    {
        public override IFixableDirection Figure => ((DimensionsPanel)Parent).direction;
    }

    /// <summary>
    /// A distance measurement on the figure, added or removed. The grid records the
    /// property set as the undo step, and this setter runs inside it - so the drawing is
    /// changed directly (the undo library refuses an action recorded from within another);
    /// undo sets the property back, which removes or re-adds the measurement. The same one:
    /// the measurement taken away is kept for its figure, so that it comes back where it
    /// was dragged to and a drag of it in the undo history still finds it.
    /// </summary>
    [PropertyGridVisible]
    [PropertyGridGroup("Length")]
    [PropertyGridCustomValueProvider(typeof(ConditionalPropertyValue))]
    public bool Show
    {
        get
        {
            return LengthConstraint.FindMeasurement(length) != null;
        }
        set
        {
            var existing = LengthConstraint.FindMeasurement(length);
            if (value && existing == null)
            {
                if (!retiredMeasurements.TryGetValue(figure, out var measurement)
                    || !measurement.Dependencies.SequenceEqual(length.MeasuredFigures))
                {
                    measurement = LengthConstraint.CreateMeasurement(length);
                }

                retiredMeasurements.Remove(figure);
                figure.Drawing.Figures.Return(measurement, owner: figure);
            }
            else if (!value && existing != null)
            {
                figure.Drawing.Figures.Retire(existing);
                retiredMeasurements.AddOrUpdate(figure, existing);
            }
        }
    }

    // by figure, not in the panel: the panel is made anew every time it is shown
    static readonly ConditionalWeakTable<IFigure, DistanceMeasurement> retiredMeasurements = new ConditionalWeakTable<IFigure, DistanceMeasurement>();

    [PropertyGridName("Fix length")]
    [PropertyGridGroup("Length")]
    [PropertyGridIcon(PropertyGridIcon.Lock)]
    public void FixLength()
    {
        length.FixLength();
        ShowAgain();
    }

    [PropertyGridName("Free length")]
    [PropertyGridGroup("Length")]
    [PropertyGridIcon(PropertyGridIcon.Unlock)]
    public void FreeLength()
    {
        length.FreeLength();
        ShowAgain();
    }

    [PropertyGridName("Fix direction")]
    [PropertyGridGroup("Direction")]
    [PropertyGridIcon(PropertyGridIcon.Lock)]
    public void FixDirection()
    {
        direction.FixDirection();
        ShowAgain();
    }

    [PropertyGridName("Free direction")]
    [PropertyGridGroup("Direction")]
    [PropertyGridIcon(PropertyGridIcon.Unlock)]
    public void FreeDirection()
    {
        direction.FreeDirection();
        ShowAgain();
    }

    /// <summary>The direction fixed at 0° or 180°, whichever is nearer</summary>
    [PropertyGridName("Horizontal")]
    [PropertyGridGroup("Direction")]
    [PropertyGridIcon(PropertyGridIcon.Horizontal)]
    public void MakeHorizontal()
    {
        direction.MakeHorizontal();
        ShowAgain();
    }

    /// <summary>The direction fixed at 90° or 270°, whichever is nearer</summary>
    [PropertyGridName("Vertical")]
    [PropertyGridGroup("Direction")]
    [PropertyGridIcon(PropertyGridIcon.Vertical)]
    public void MakeVertical()
    {
        direction.MakeVertical();
        ShowAgain();
    }

    /// <summary>Closes the panel; the figure stays as it is</summary>
    [PropertyGridIcon(PropertyGridIcon.Check)]
    public void OK()
    {
        figure.Drawing.RaiseDisplayProperties(null);
        figure.Drawing.ShowBehaviorHint();
    }

    // the figure shows its own grid after the verb; this panel comes back on top of that
    void ShowAgain()
    {
        figure.Drawing.RaiseDisplayProperties(this);
    }

    /// <summary>
    /// The verbs that apply, captioned by the figure ("Fix radius"): the length's, the
    /// direction's (only with its row: none for a direction nothing can set), then OK
    /// </summary>
    public IEnumerable<IOperationDescription> GetMethods()
    {
        var names = new List<string>();
        if (length != null)
        {
            names.AddRange(new[] { nameof(FixLength), nameof(FreeLength) });
        }

        if (ShowsDirection)
        {
            names.AddRange(new[] { nameof(FixDirection), nameof(FreeDirection), nameof(MakeHorizontal), nameof(MakeVertical) });
        }

        foreach (var name in names)
        {
            if (conditions.CanEdit(name))
            {
                yield return new CaptionedMethod(MethodDescription.Get<DimensionsPanel>(name), conditions);
            }
        }

        yield return MethodDescription.Get<DimensionsPanel>("OK");
    }

    public bool CanEdit(string propertyName)
    {
        return conditions.CanEdit(propertyName);
    }

    public string Caption(string propertyName, string defaultCaption)
    {
        return conditions.Caption(propertyName, defaultCaption);
    }

    public override string ToString()
    {
        return figure.Title;
    }
}
