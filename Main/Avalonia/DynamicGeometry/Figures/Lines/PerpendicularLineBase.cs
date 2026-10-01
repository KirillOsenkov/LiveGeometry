using System.Collections.Generic;
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

    // the perpendicular bisector of AB doesn't run through A and B: not "AB"
    protected override IReadOnlyList<string> NamesFromDependencies()
    {
        return null;
    }

    protected override bool IsThroughTwoPoints
    {
        get
        {
            return false;
        }
    }

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

        if (!Exists)
        {
            Mark.Hide();
            return;
        }

        bool hasRightAngle = TryGetRightAngle(out Point vertex, out PointPair baseLine, out Point pointAcross);

        // Once, when the line is first worked out; from then on the corner only changes by
        // a click. Also for a hidden line (a square's helper) and one whose foot is beyond
        // the end of its segment: chosen only when the mark first showed, the corner
        // changed - and with it what a file saves - when the line was shown or a point was
        // dragged, and undo of that could not put it back.
        if (!cornerChosen)
        {
            cornerChosen = true;
            Mark.Corner = RightAngleMark.GetRoomiestCorner(vertex, baseLine, pointAcross);
        }

        if (!Visible || !hasRightAngle)
        {
            Mark.Hide();
            return;
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
