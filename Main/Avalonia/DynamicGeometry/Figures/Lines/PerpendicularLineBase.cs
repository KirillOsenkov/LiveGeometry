using System.Xml;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;

namespace DynamicGeometry;

/// <summary>
/// A line that is perpendicular to another by construction, and says so with a
/// <see cref="RightAngleMark"/> where they meet.
/// </summary>
public abstract class PerpendicularLineBase : LineTwoPoints
{
    RightAngleMark rightAngleMark;
    bool cornerChosen;

    // not a field initializer: virtual members get called from the base constructor
    RightAngleMark Mark
    {
        get
        {
            if (rightAngleMark == null)
            {
                rightAngleMark = new RightAngleMark(this);
            }

            return rightAngleMark;
        }
    }

    /// <summary>
    /// Where the right angle is, or false if it isn't there to be marked
    /// (the foot of the perpendicular is beyond the end of the segment, say).
    /// </summary>
    /// <param name="vertex">Where the two lines meet</param>
    /// <param name="baseLine">The line this one is perpendicular to</param>
    /// <param name="pointAcross">A point on this line that tells which side the mark starts on</param>
    protected abstract bool TryGetRightAngle(out Point vertex, out PointPair baseLine, out Point pointAcross);

    [PropertyGridVisible]
    [PropertyGridName("Right angle mark")]
    public bool ShowRightAngle
    {
        get
        {
            return Mark.IsEnabled;
        }
        set
        {
            Mark.IsEnabled = value;
        }
    }

    public override bool Visible
    {
        get
        {
            return base.Visible;
        }
        set
        {
            base.Visible = value;
            if (!value)
            {
                Mark.Hide();
            }
            else if (Drawing != null)
            {
                UpdateVisual();
            }
        }
    }

    public override void UpdateVisual()
    {
        base.UpdateVisual();

        Point vertex;
        PointPair baseLine;
        Point pointAcross;
        if (!Exists || !Visible || !TryGetRightAngle(out vertex, out baseLine, out pointAcross))
        {
            Mark.Hide();
            return;
        }

        // once, when the line first shows up; from then on the corner only changes by a click
        if (!cornerChosen)
        {
            cornerChosen = true;
            Mark.Corner = RightAngleMark.GetRoomiestCorner(vertex, baseLine, pointAcross);
        }

        Mark.Show(Drawing, vertex, baseLine);
    }

    public override void OnAddingToCanvas(Canvas newContainer)
    {
        base.OnAddingToCanvas(newContainer);
        Mark.OnAddingToCanvas(newContainer);
    }

    public override void OnRemovingFromCanvas(Canvas leavingContainer)
    {
        base.OnRemovingFromCanvas(leavingContainer);
        Mark.OnRemovingFromCanvas(leavingContainer);
    }

    public override void ReadXml(XElement element)
    {
        base.ReadXml(element);
        if (Mark.ReadXml(element) != null)
        {
            cornerChosen = true;
        }
    }

    public override void WriteXml(XmlWriter writer)
    {
        base.WriteXml(writer);
        Mark.WriteXml(writer);
    }
}
