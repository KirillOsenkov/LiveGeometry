using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Avalonia;

namespace DynamicGeometry;

public enum AxisDirection
{
    X,
    Y
}

/// <summary>
/// An axis of the coordinate grid as a line to build on: a point on the x-axis, where a
/// circle crosses the y-axis, a reflection in either. The grid draws the axes; this line is
/// never drawn itself (it is there to be selected in the Figure List). A drawing has one of
/// each for its whole life (<see cref="Drawing.GetAxisLine"/>), so that there is never a
/// second x-axis, and they are in its list only while something is built on them: a figure
/// that comes to depend on one brings both in (<see cref="AddMissing"/>: one is soon wanted
/// after the other), and they leave together once nothing depends on either
/// (<see cref="RemoveFigureAction.FindOrphanedAuxiliaries"/>). Out of the list they are
/// still hit (<see cref="RootFigureList"/>), so that a tool can take one - but only while
/// the grid shows its axes.
/// </summary>
public class AxisLine : LineBase, ILine, IConditionalProperties
{
    public AxisLine(AxisDirection direction)
    {
        Direction = direction;
        mName = direction == AxisDirection.X ? "x-axis" : "y-axis";
        Auxiliary = true;
        ApplyStyle();
    }

    public AxisDirection Direction { get; }

    protected override int DefaultZOrder()
    {
        return (int)ZOrder.Axes;
    }

    public override PointPair Coordinates
    {
        get
        {
            var direction = Direction == AxisDirection.X ? new Point(1, 0) : new Point(0, 1);
            return new PointPair(new Point(0, 0), direction);
        }
    }

    public override PointPair OnScreenCoordinates
    {
        get
        {
            return Math.GetLineFromSegment(Coordinates, CanvasLogicalBorders);
        }
    }

    /// <summary>A click takes the axis only where the grid shows it</summary>
    public override bool IsHitTestVisible
    {
        get
        {
            return Drawing != null && Drawing.CoordinateGrid != null && Drawing.CoordinateGrid.ShowsAxes;
        }
        set
        {
        }
    }

    /// <summary>
    /// The grid draws the axis: the line's own stroke is transparent, there for the halo of a
    /// selection (in the Figure List) and for hit testing, whatever style it is given
    /// </summary>
    public override void ApplyStyle()
    {
        Shape.Stroke = Avalonia.Media.Brushes.Transparent;
        Shape.StrokeThickness = 1;
        if (Drawing != null)
        {
            UpdateVisual();
        }
    }

    /// <summary>"on x-axis": the name says what it is</summary>
    public override string Noun
    {
        get
        {
            return null;
        }
    }

    public override string Construction
    {
        get
        {
            return Direction == AxisDirection.X ? "y = 0" : "x = 0";
        }
    }

    // The app names it, the grid draws it and nothing moves it: no rows for its name, its
    // look, whether it shows or is locked.

    [PropertyGridVisible(false)]
    public override string Name
    {
        get
        {
            return base.Name;
        }
        set
        {
            base.Name = value;
        }
    }

    [PropertyGridVisible(false)]
    public override IFigureStyle StyleDisplay
    {
        get
        {
            return base.StyleDisplay;
        }
        set
        {
            base.StyleDisplay = value;
        }
    }

    [PropertyGridVisible(false)]
    public override bool Visible
    {
        get
        {
            return base.Visible;
        }
        set
        {
            base.Visible = value;
        }
    }

    [PropertyGridVisible(false)]
    public override bool Locked
    {
        get
        {
            return base.Locked;
        }
        set
        {
            base.Locked = value;
        }
    }

    [PropertyGridVisible(false)]
    public override bool ShowName
    {
        get
        {
            return base.ShowName;
        }
        set
        {
            base.ShowName = value;
        }
    }

    public bool CanEdit(string propertyName)
    {
        return propertyName != nameof(EditStyleButton) && propertyName != nameof(CreateNewStyle);
    }

    public string Caption(string propertyName, string defaultCaption)
    {
        return defaultCaption;
    }

    /// <summary>
    /// Which axis is all a file says: reading one gives the drawing's own
    /// (<see cref="DrawingDeserializer"/>), never a second one
    /// </summary>
    public override void WriteXml(XmlWriter writer)
    {
        writer.WriteAttributeString("Axis", Direction.ToString());
    }

    public static AxisDirection ReadDirection(XElement element)
    {
        return element.ReadString("Axis") == nameof(AxisDirection.Y) ? AxisDirection.Y : AxisDirection.X;
    }

    /// <summary>
    /// Figures are coming into the drawing: when one of them is built on an axis line the
    /// drawing's list doesn't have (or is one), the axis lines it lacks come in first, both
    /// of them. Returns those added, for the caller to take out again on undo. Not while a
    /// file is read: it has the axes it uses, each in its place, and adds them one by one.
    /// </summary>
    public static List<AxisLine> AddMissing(Drawing drawing, ICollection<IFigure> adding)
    {
        var added = new List<AxisLine>();
        if (drawing.IsReading)
        {
            return added;
        }

        bool needed = adding
            .SelectMany(figure => figure.Dependencies.Prepend(figure))
            .OfType<AxisLine>()
            .Any(axis => !drawing.Figures.Contains(axis));
        if (!needed)
        {
            return added;
        }

        foreach (var direction in new[] { AxisDirection.X, AxisDirection.Y })
        {
            var axis = drawing.GetAxisLine(direction);
            if (!drawing.Figures.Contains(axis) && !adding.Contains(axis))
            {
                drawing.Figures.Add(axis);
                added.Add(axis);
            }
        }

        return added;
    }

    /// <summary>Takes out the axis lines <see cref="AddMissing"/> added</summary>
    public static void Remove(Drawing drawing, List<AxisLine> added)
    {
        for (int i = added.Count - 1; i >= 0; i--)
        {
            drawing.Figures.Remove(added[i]);
        }
    }
}
