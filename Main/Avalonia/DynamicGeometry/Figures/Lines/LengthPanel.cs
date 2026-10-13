using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace DynamicGeometry;

/// <summary>
/// The side panel right after a figure with a length is made (a segment, a vector, a
/// regular polygon, a circle): just its length and Fix length or Free length, forwarded to
/// the figure, so that "a segment of length 2" is draw, type 2, click Fix - without a
/// length box on every tool. A segment or a vector has its direction under that, with Fix
/// direction, Horizontal and Vertical: a horizontal segment of length 2 is draw, click
/// Horizontal, type 2, click Fix length. Captions come from the figure (Side, Radius, Fix
/// radius); the title is the figure's ("Segment AB"). The figure's own grid has the same
/// rows. A regular polygon's panel starts with its number of sides, the polygon's own row.
/// </summary>
public class LengthPanel :
    IConditionalProperties,
    ICustomPropertyProvider,
    ICustomMethodProvider,
    IPropertyGridHost,
    INotifyPropertyChanged
{
    readonly IFixableLength figure;
    readonly IFixableDirection direction;
    PropertyGrid propertyGrid;

    public LengthPanel(IFixableLength figure)
    {
        this.figure = figure;
        direction = figure as IFixableDirection;
    }

    public event PropertyChangedEventHandler PropertyChanged;

    /// <summary>
    /// Set by the grid while it shows the panel: the figure's own changes (a new number of
    /// sides, which also changes the side and the title) are passed on as the panel's for
    /// that long.
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
        // (a regular polygon says that its side changed with the number of sides)
        PropertyChanged?.Invoke(this, e);
    }

    /// <summary>
    /// The rows: a regular polygon's number of sides first (the polygon's own property, its
    /// up/down within the polygon's limits, an undo step on the polygon), then the length,
    /// then Show where there is a figure to hang a measurement on, then the direction where
    /// the figure has one that can be set.
    /// </summary>
    public IEnumerable<IValueProvider> GetProperties()
    {
        if (figure is RegularPolygon polygon)
        {
            yield return PropertyDiscoveryStrategy.CreateValueProvider(polygon, nameof(RegularPolygon.NumberOfSides));
        }

        yield return PropertyDiscoveryStrategy.CreateValueProvider(this, nameof(Length));
        // a regular polygon's sides are its own children, nothing in the drawing to measure
        if (figure.MeasuredFigures != null)
        {
            yield return PropertyDiscoveryStrategy.CreateValueProvider(this, nameof(Show));
        }

        if (direction != null && direction.CanEdit(nameof(Angle)))
        {
            yield return PropertyDiscoveryStrategy.CreateValueProvider(this, nameof(Angle));
        }
    }

    [PropertyGridVisible]
    [PropertyGridGroup("Length")]
    [PropertyGridPreferredEditor("UpDown")]
    [PropertyGridCustomValueProvider(typeof(PanelLengthValue))]
    public double Length
    {
        get { return figure.Length; }
        set { figure.Length = value; }
    }

    public class PanelLengthValue : LengthPropertyValue
    {
        public override IFixableLength Figure => ((LengthPanel)Parent).figure;
    }

    /// <summary>The figure's direction, in degrees (named as the figure's row is, so that its CanEdit answers for both)</summary>
    [PropertyGridVisible]
    [PropertyGridName("Direction")]
    [PropertyGridGroup("Direction")]
    [PropertyGridPreferredEditor("UpDown")]
    [PropertyGridCustomValueProvider(typeof(PanelDirectionValue))]
    public double Angle
    {
        get { return direction.Angle; }
        set { direction.Angle = value; }
    }

    public class PanelDirectionValue : DirectionPropertyValue
    {
        public override IFixableDirection Figure => ((LengthPanel)Parent).direction;
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
            return LengthConstraint.FindMeasurement(figure) != null;
        }
        set
        {
            var existing = LengthConstraint.FindMeasurement(figure);
            if (value && existing == null)
            {
                if (!retiredMeasurements.TryGetValue(figure, out var measurement)
                    || !measurement.Dependencies.SequenceEqual(figure.MeasuredFigures))
                {
                    measurement = LengthConstraint.CreateMeasurement(figure);
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
        figure.FixLength();
        ShowAgain();
    }

    [PropertyGridName("Free length")]
    [PropertyGridGroup("Length")]
    [PropertyGridIcon(PropertyGridIcon.Unlock)]
    public void FreeLength()
    {
        figure.FreeLength();
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
        var names = new List<string>() { nameof(FixLength), nameof(FreeLength) };
        if (direction != null && direction.CanEdit(nameof(Angle)))
        {
            names.AddRange(new[] { nameof(FixDirection), nameof(FreeDirection), nameof(MakeHorizontal), nameof(MakeVertical) });
        }

        foreach (var name in names)
        {
            if (figure.CanEdit(name))
            {
                yield return new CaptionedMethod(MethodDescription.Get<LengthPanel>(name), figure);
            }
        }

        yield return MethodDescription.Get<LengthPanel>("OK");
    }

    public bool CanEdit(string propertyName)
    {
        return figure.CanEdit(propertyName);
    }

    public string Caption(string propertyName, string defaultCaption)
    {
        return figure.Caption(propertyName, defaultCaption);
    }

    public override string ToString()
    {
        return figure.Title;
    }
}
