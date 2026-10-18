using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using GuiLabs.Undo;

namespace LiveGeometry;

/// <summary>
/// The side panel right after a regular polygon is made (<see cref="RegularPolygonCreator"/>):
/// its number of sides, its side with Fix length and Free length, and OK, which brings the
/// tool's panel back for the next polygon. The polygon is plain figures
/// (<see cref="RegularPolygonBuilder"/>), and its count is no property of any of them: a new
/// count builds it again - the figures made for it go, new ones come - one undo step, for as
/// long as this panel is up. Undo of it closes the panel (only a tool's own panel comes back
/// after undo), so the panel never holds figures that have left. On a side that is a segment
/// or a vector, what shows is that side's panel with the number of sides on top
/// (<see cref="Shown"/>): its length and its direction, as after the Square.
/// </summary>
public class RegularPolygonPanel : IConditionalProperties
{
    /// <param name="first">The center, or the first of two adjacent vertices: what the side is measured from</param>
    /// <param name="made">What the tool made, the polygon last</param>
    public RegularPolygonPanel(IPoint first, int sides, bool onSide, IList<IFigure> made)
    {
        center = first;
        this.sides = sides;
        this.onSide = onSide;
        this.made = made.ToList();
    }

    // the center is never replaced while the panel is up; the two points on a side may be
    // (Fix length, Horizontal: a side the drawing had may run the other way, and then it is the
    // first click that the side's verbs replace), and so may the vertex by the center, so they
    // are read off the polygon
    readonly IPoint center;
    readonly bool onSide;
    int sides;
    List<IFigure> made;

    Polygon Polygon => (Polygon)made[made.Count - 1];

    /// <summary>What the side is measured from: the center, or on a side the polygon's first vertex</summary>
    IPoint Pivot => onSide ? (IPoint)Polygon.Dependencies[0] : center;

    /// <summary>What the status bar says under the panel</summary>
    public string Hint => Polygon.Title + (BaseSide != null
        ? ": set its number of sides, or its side's length and direction, in the panel, or fix them."
        : ": set its number of sides or its side in the panel, or fix the side.");

    IPoint Vertex => (IPoint)Polygon.Dependencies[onSide ? 1 : 0];

    /// <summary>The side the polygon stands on, the two points clicked, when a segment or a vector draws it: on a side only</summary>
    IFixableLength BaseSide => onSide ? Polygon.Drawing.Figures.FindLine(Pivot, Vertex) as IFixableLength : null;

    /// <summary>
    /// What the side panel shows for the polygon: the panel of the side it stands on with the
    /// number of sides on top, where there is such a side (<see cref="BaseSide"/>), so that
    /// its length and direction are set and fixed as a square's are; else this panel. Made
    /// anew each time: a new number of sides may make the side anew.
    /// </summary>
    public object Shown
    {
        get
        {
            var side = BaseSide;
            return side != null ? new DimensionsPanel(side) { RegularPolygon = this } : this;
        }
    }

    /// <summary>The side over the distance from the pivot to the vertex: 1 on a side, the chord of the central angle by the center</summary>
    double SideFactor => onSide ? 1 : 2 * System.Math.Sin(System.Math.PI / sides);

    [PropertyGridVisible]
    [PropertyGridName("Number of sides")]
    [Domain(RegularPolygonBuilder.MinimumSides, RegularPolygonBuilder.MaximumSides)]
    [PropertyGridCustomValueProvider(typeof(SidesValue))]
    public int NumberOfSides
    {
        get => sides; set => Rebuild(Polygon.Drawing.ActionManager, value);
    }

    /// <summary>A new count is the rebuild's own undo step, not a property set: what it adds and removes are actions of their own</summary>
    public class SidesValue : PropertyValue, ISelfRecordingValue
    {
        public bool TryRecord(ActionManager actionManager, object value)
        {
            ((RegularPolygonPanel)Parent).Rebuild(actionManager, System.Convert.ToInt32(value, CultureInfo.InvariantCulture));
            return true;
        }
    }

    [PropertyGridVisible]
    [PropertyGridName("Side")]
    [PropertyGridGroup("Side")]
    [PropertyGridPreferredEditor("UpDown")]
    [PropertyGridCustomValueProvider(typeof(SideValue))]
    public double Side
    {
        get => Vertex.Coordinates.Distance(Pivot.Coordinates) * SideFactor; set => LengthConstraint.SetDistance(Vertex, Pivot, value / SideFactor);
    }

    /// <summary>Undo of a side puts the vertex back where it was (<see cref="LengthPropertyValue"/>)</summary>
    public class SideValue : LengthPropertyValue
    {
        RegularPolygonPanel Panel => (RegularPolygonPanel)Parent;

        protected override IPoint Endpoint => Panel.Vertex;

        protected override Drawing Drawing => Panel.Polygon.Drawing;
    }

    /// <summary>The polygon keeps its size: the vertex stays at its distance from the pivot</summary>
    [PropertyGridVisible]
    [PropertyGridName("Fix length")]
    [PropertyGridGroup("Side")]
    [PropertyGridIcon(PropertyGridIcon.Lock)]
    public void FixLength()
    {
        LengthConstraint.Fix(Vertex, Pivot);
        Polygon.Drawing.RaiseDisplayProperties(this);
    }

    [PropertyGridVisible]
    [PropertyGridName("Free length")]
    [PropertyGridGroup("Side")]
    [PropertyGridIcon(PropertyGridIcon.Unlock)]
    public void FreeLength()
    {
        var fixedEnd = LengthConstraint.FixedEnd(Vertex, Pivot);
        if (fixedEnd != null)
        {
            LengthConstraint.Free(fixedEnd);
            Polygon.Drawing.RaiseDisplayProperties(this);
        }
    }

    /// <summary>Back to the tool's panel, for the next polygon; this one stays as it is</summary>
    [PropertyGridVisible]
    [PropertyGridIcon(PropertyGridIcon.Check)]
    public void OK()
    {
        var drawing = Polygon.Drawing;
        drawing.RaiseDisplayProperties(null);
        drawing.RaiseDisplayProperties(drawing.Behavior?.PropertyBag);
        drawing.ShowBehaviorHint();
    }

    /// <summary>
    /// The polygon made again with this many sides on the same two points: what was made
    /// for it removed, last first, and the new figures added, one transaction. Not while
    /// something else is built on what was made (only its names: a point's label goes and
    /// comes with its point). The next polygon of the tools gets the count too.
    /// </summary>
    void Rebuild(ActionManager actionManager, int count)
    {
        count = System.Math.Clamp(count, RegularPolygonBuilder.MinimumSides, RegularPolygonBuilder.MaximumSides);
        var drawing = Polygon.Drawing;
        if (count == sides)
        {
            return;
        }

        var builtOn = made
            .SelectMany(figure => figure.Dependents)
            .FirstOrDefault(dependent => !made.Contains(dependent) && !(dependent is NameLabel));
        if (builtOn != null)
        {
            drawing.RaiseStatusNotification(builtOn.Title + " is built on " + Polygon.Title + ", which keeps its " + sides + " sides.");
            return;
        }

        var pivot = Pivot;
        var vertex = Vertex;
        List<IFigure> rebuilt;
        Polygon polygon;
        using (Transaction.Create(actionManager, false))
        {
            foreach (var figure in Enumerable.Reverse(made))
            {
                if (drawing.Figures.Contains(figure))
                {
                    Actions.Remove(figure);
                }
            }

            rebuilt = RegularPolygonBuilder.Create(drawing, pivot, vertex, count, onSide, out polygon);
            foreach (var figure in rebuilt)
            {
                Actions.Add(drawing, figure);
            }
        }

        RegularPolygonCreator.NextNumberOfSides = count;

        // a polygon the drawing had on these vertices already is not this panel's to change
        if (!rebuilt.Contains(polygon))
        {
            drawing.RaiseDisplayProperties(null);
            drawing.RaiseStatusNotification(polygon.Title + " is there already.");
            return;
        }

        made = rebuilt;
        sides = count;
        drawing.RaiseDisplayProperties(Shown);
        drawing.RaiseStatusNotification(Hint);
    }

    public bool CanEdit(string propertyName)
    {
        bool isFixed = LengthConstraint.FixedEnd(Vertex, Pivot) != null;
        switch (propertyName)
        {
            case nameof(Side):
                return isFixed || LengthConstraint.CanStretch(Vertex, Pivot);
            case nameof(FixLength):
                return !isFixed && LengthConstraint.CanFix(Vertex, Pivot);
            case nameof(FreeLength):
                return isFixed && LengthConstraint.CanFree(LengthConstraint.FixedEnd(Vertex, Pivot));
            default:
                return true;
        }
    }

    public string Caption(string propertyName, string defaultCaption)
    {
        return defaultCaption;
    }

    // the title of the panel: the polygon ("Hexagon ABCDEF")
    public override string ToString()
    {
        return Polygon.Title;
    }
}
