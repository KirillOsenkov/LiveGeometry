using System.Collections.Generic;
using System.Linq;

namespace LiveGeometry;

/// <summary>
/// The side panel right after the Length tool measured a figure with more than one length
/// (a circle, an arc, a sector): the measurement's own Measure row, which says the radius,
/// the diameter, the circumference... (an undo step on the measurement), and OK. As the
/// arc's panel does for its sweep (<see cref="ArcPanel"/>): the click measures, and one more
/// picks which length.
/// </summary>
public class LengthMeasurementPanel : ICustomPropertyProvider, ICustomMethodProvider
{
    readonly LengthMeasurement measurement;

    public LengthMeasurementPanel(LengthMeasurement measurement)
    {
        this.measurement = measurement;
    }

    /// <summary>Puts the panel for the measurement up, and says so in the status bar</summary>
    public static void Show(LengthMeasurement measurement)
    {
        var drawing = measurement.Drawing;
        drawing.RaiseDisplayProperties(new LengthMeasurementPanel(measurement));
        var lengths = measurement.Quantities.Select(quantity => LengthMeasurement.KindOf(quantity).ToLowerInvariant()).ToList();
        drawing.RaiseStatusNotification(
            "Pick its " + string.Join(", ", lengths.Take(lengths.Count - 1)) + " or " + lengths.Last() + " in the panel, or go on.");
    }

    /// <summary>"Measure circle c": the same whichever length is picked</summary>
    string Title
    {
        get { return measurement.Dependencies.Count == 1 ? "Measure " + ConstructionText.Of(measurement.Dependencies[0]) : measurement.Title; }
    }

    public IEnumerable<IValueProvider> GetProperties()
    {
        yield return PropertyDiscoveryStrategy.CreateValueProvider(measurement, nameof(LengthMeasurement.Quantity));
    }

    public IEnumerable<IOperationDescription> GetMethods()
    {
        yield return MethodDescription.Get<LengthMeasurementPanel>(nameof(OK));
    }

    /// <summary>Closes the panel; the measurement stays as it is</summary>
    [PropertyGridIcon(PropertyGridIcon.Check)]
    public void OK()
    {
        measurement.Drawing.RaiseDisplayProperties(null);
        measurement.Drawing.ShowBehaviorHint();
    }

    // the title of the panel
    public override string ToString()
    {
        return Title;
    }
}
