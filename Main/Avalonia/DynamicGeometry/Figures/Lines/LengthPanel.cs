using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace DynamicGeometry;

/// <summary>
/// The side panel right after a figure with a length is made (a segment, a vector, a
/// regular polygon, a circle): just its length and Fix length or Free length, forwarded to
/// the figure, so that "a segment of length 2" is draw, type 2, click Fix - without a
/// length box on every tool. Captions come from the figure (Side, Radius, Fix radius);
/// the title is the figure's ("Segment AB"). The figure's own grid has the same rows. A
/// regular polygon's panel starts with its number of sides, the polygon's own row.
/// </summary>
public class LengthPanel :
    IConditionalProperties,
    ICustomPropertyProvider,
    ICustomMethodProvider,
    IPropertyGridHost,
    INotifyPropertyChanged
{
    readonly IFixableLength figure;
    PropertyGrid propertyGrid;

    public LengthPanel(IFixableLength figure)
    {
        this.figure = figure;
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
    /// then Show where there is a figure to hang a measurement on.
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
    }

    [PropertyGridVisible]
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

    /// <summary>
    /// A distance measurement on the figure, added or removed. The grid records the
    /// property set as the undo step, and this setter runs inside it - so the drawing is
    /// changed directly (the undo library refuses an action recorded from within another);
    /// undo sets the property back, which removes or re-adds the measurement. The same one:
    /// the measurement taken away is kept for its figure, so that it comes back where it
    /// was dragged to and a drag of it in the undo history still finds it.
    /// </summary>
    [PropertyGridVisible]
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
    [PropertyGridIcon(PropertyGridIcon.Lock)]
    public void FixLength()
    {
        figure.FixLength();
        ShowAgain();
    }

    [PropertyGridName("Free length")]
    [PropertyGridIcon(PropertyGridIcon.Unlock)]
    public void FreeLength()
    {
        figure.FreeLength();
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

    /// <summary>The verb that applies, captioned by the figure ("Fix radius"), then OK</summary>
    public IEnumerable<IOperationDescription> GetMethods()
    {
        foreach (var name in new[] { "FixLength", "FreeLength" })
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
