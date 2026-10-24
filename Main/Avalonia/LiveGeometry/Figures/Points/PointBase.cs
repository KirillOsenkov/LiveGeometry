using Avalonia;
using Avalonia.Controls.Shapes;
using System.Collections.Generic;

namespace LiveGeometry;

public partial class PointBase : CoordinatesShapeBase<Shape>, IPoint
{
    private static bool mSuppress;
    public static bool SuppressAutoLabelPoints
    {
        get => mSuppress; set => mSuppress = value;
    }

    public override string GenerateFigureName(List<string> blacklist)
    {
        // a hidden helper (the square's midpoint) mustn't take a letter from the points on
        // screen: the square would be ABCE
        if (!Visible)
        {
            return base.GenerateFigureName(blacklist);
        }

        var alphabet = Settings.Instance.PointAlphabet;
        for (int i = 0; ; i++)
        {
            string number = i.ToString();
            if (i == 0)
            {
                number = "";
            }
            foreach (var letter in alphabet)
            {
                var candidate = letter.ToString() + number;
                if (this.NameAvailable(candidate))
                {
                    if (blacklist != null)
                    {
                        if (!blacklist.Contains(candidate))
                        {
                            return candidate;
                        }
                    }
                    else
                    {
                        return candidate;
                    }
                }
            }
        }
    }

    protected override string Kind => "Point";

    /// <summary>A point goes by its name: "through E", not "through point E"</summary>
    public override string Noun => null;

    // The label (NameLabel) is added and removed with the point and by its Show name and
    // Show coordinates, outside the undo history, so it is the same object every time
    // (retiredNameLabel): the history may hold a drag of it, and its place is its own.

    // the point left the drawing with its label on (undo of adding it): it comes back with it
    bool returnsWithLabel;

    // new points are labeled once, when they are made: redo makes the point as it was
    bool wasInDrawing;

    public override void OnAddingToDrawing(Drawing drawing)
    {
        base.OnAddingToDrawing(drawing);

        // Make sure this is in the drawing's figure list before labeling.
        if (!Drawing.Figures.Contains(this))
        {
            return;
        }

        bool returning = wasInDrawing;
        bool withLabel = returnsWithLabel;
        wasInDrawing = true;
        returnsWithLabel = false;

        // while suppressed, whoever adds the point sees to its label (undo of a deletion
        // puts the label back itself)
        if (SuppressAutoLabelPoints)
        {
            return;
        }

        if (returning)
        {
            if (withLabel && NameLabel == null && retiredNameLabel != null && !Drawing.Figures.Contains(retiredNameLabel))
            {
                NameLabel = retiredNameLabel;
                retiredNameLabel = null;
                Drawing.Figures.Add(NameLabel);
            }
        }
        else if (Settings.Instance.AutoLabelPoints && IsHitTestVisible && Visible)
        {
            // a constructed point (a midpoint, a crossing) is where its figures put it only
            // once it is worked out, which the list does after this: the name's first place
            // is worked out from where the point is, and went to the origin
            UpdateExistence();
            Recalculate();
            ShowName = true;
        }
    }

    public override void OnRemovingFromDrawing(Drawing drawing)
    {
        var label = NameLabel;
        if (label != null)
        {
            retiredNameLabel = label;
            returnsWithLabel = true;
            NameLabel = null;
            Drawing.Figures.Remove(label);
        }
    }

    // a constructed point: dragged, it moves what it is built on (FreePoint and the
    // others that can be dragged by themselves say otherwise)
    protected override ZOrder DefaultLayer()
    {
        return ZOrder.FixedPoints;
    }

    protected override Shape CreateShape()
    {
        return Factory.CreatePointShape();
    }

    public virtual double X
    {
        get => Coordinates.X; set => Coordinates = Coordinates.SetX(value);
    }

    public virtual double Y
    {
        get => Coordinates.Y; set => Coordinates = Coordinates.SetY(value);
    }

    double PointSize => Shape.ActualWidth / 2;

    public override IFigure HitTest(Point point)
    {
        double tolerance = CursorTolerance + ToLogical(PointSize);
        if (point.X >= Coordinates.X - tolerance
            && point.X <= Coordinates.X + tolerance
            && point.Y >= Coordinates.Y - tolerance
            && point.Y <= Coordinates.Y + tolerance)
        {
            return this;
        }

        return null;
    }

    /// <summary>The label of the name and the coordinates, while the point shows either</summary>
    PointLabel PointLabel => NameLabel as PointLabel;

    protected override NameLabel CreateNameLabel()
    {
        return Factory.CreatePointLabel(Drawing, new[] { this });
    }

    [PropertyGridVisible]
    [ContextMenu]
    [PropertyGridName("Show name")]
    [PropertyGridSection(PropertyGridSection.Display, 10)]
    public bool ShowName
    {
        get => PointLabel?.ShowName == true;
        set
        {
            if (ShowName == value)
            {
                return;
            }

            HasNameLabel = true;
            PointLabel.ShowName = value;
            RetireLabelSayingNothing();
        }
    }

    [PropertyGridVisible]
    [ContextMenu]
    [PropertyGridName("Show coordinates")]
    [PropertyGridSection(PropertyGridSection.Display, 20)]
    public bool ShowCoordinates
    {
        get => PointLabel?.ShowCoordinates == true;
        set
        {
            if (ShowCoordinates == value)
            {
                return;
            }

            HasNameLabel = true;
            PointLabel.ShowCoordinates = value;
            RetireLabelSayingNothing();
        }
    }

    /// <summary>The label goes when it shows neither the name nor the coordinates, and is retired: it says nothing when it comes back</summary>
    void RetireLabelSayingNothing()
    {
        if (PointLabel is { ShowName: false, ShowCoordinates: false })
        {
            HasNameLabel = false;
        }
    }

    /// <summary>
    /// The point, an anchor of a Bezier path, has both its handles worked out by the path
    /// (<see cref="BezierPath.SmoothAnchor"/>). Any kind of point: a path's anchor may be a
    /// midpoint or a crossing as well as a point that can be dragged.
    /// </summary>
    [PropertyGridVisible]
    [ContextMenu]
    [PropertyGridName("Smooth automatically")]
    [PropertyGridIcon(PropertyGridIcon.Arc)]
    [PropertyGridCondition(nameof(CanSmoothAnchor))]
    public void SmoothAnchor()
    {
        BezierPath.SmoothAnchor(this);
        Drawing?.RaiseDisplayProperties(this);
    }

    /// <summary>The point, an anchor of a Bezier path, has both its handles on it: the path turns there (<see cref="BezierPath.SharpenAnchor"/>)</summary>
    [PropertyGridVisible]
    [ContextMenu]
    [PropertyGridName("Sharp corner")]
    [PropertyGridIcon(PropertyGridIcon.Angle)]
    [PropertyGridCondition(nameof(CanSharpenAnchor))]
    public void SharpenAnchor()
    {
        BezierPath.SharpenAnchor(this);
        Drawing?.RaiseDisplayProperties(this);
    }

    public bool CanSmoothAnchor()
    {
        return BezierPath.CanSmoothAnchor(this);
    }

    public bool CanSharpenAnchor()
    {
        return BezierPath.CanSharpenAnchor(this);
    }
}