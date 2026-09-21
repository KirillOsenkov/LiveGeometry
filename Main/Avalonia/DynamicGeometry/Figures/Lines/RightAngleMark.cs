using System;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using AvaloniaShapes = Avalonia.Controls.Shapes;

namespace DynamicGeometry;

/// <summary>
/// The little square in the corner that says "these two are perpendicular". A figure that is
/// perpendicular by construction (a perpendicular line, a segment bisector) owns one and shows it
/// by itself: it is not a figure, so it isn't in the figure list and needs no cleanup.
/// An angle measured at the same vertex takes over (it draws the same sign, in its own style).
/// Which of the four corners it sits in is remembered (<see cref="Corner"/>) and never derived
/// from the geometry, so it can't jump around; a click on it with the drag tool moves it on.
/// </summary>
public class RightAngleMark
{
    /// <summary>Side of the square in pixels; an angle measurement uses the same</summary>
    public static double Size = 10;

    public static IBrush Stroke = new SolidColorBrush(Color.FromArgb(0xA0, 0x70, 0x70, 0x70));

    const string CornerAttribute = "RightAngleCorner";
    const string VisibleAttribute = "RightAngleMark";

    readonly IFigure owner;

    readonly AvaloniaShapes.Polyline sign = new AvaloniaShapes.Polyline()
    {
        Stroke = Stroke,
        StrokeThickness = 1,
        StrokeJoin = PenLineJoin.Miter,
        IsHitTestVisible = false,
        ZIndex = (int)ZOrder.Figures - 1,
        IsVisible = false
    };

    // the sign is a hairline; this is what the mouse gets: the whole square
    readonly AvaloniaShapes.Polygon clickArea = new AvaloniaShapes.Polygon()
    {
        Fill = Brushes.Transparent,
        ZIndex = (int)ZOrder.Figures - 1,
        IsVisible = false
    };

    public RightAngleMark(IFigure owner)
    {
        this.owner = owner;
        clickArea.PointerPressed += ClickArea_PointerPressed;

        // a hand where a click does something: with the drag tool only
        clickArea.PointerEntered += (s, e) => clickArea.Cursor = CanClick ? handCursor : null;
    }

    static readonly Cursor handCursor = new Cursor(StandardCursorType.Hand);

    bool CanClick
    {
        get
        {
            return owner.Drawing != null && owner.Drawing.Behavior is Dragger;
        }
    }

    int corner;

    /// <summary>
    /// 0 to 3, counterclockwise. 0 is between the direction of the base line (from its first
    /// point to its second) and that direction turned left by 90 degrees.
    /// </summary>
    public int Corner
    {
        get
        {
            return corner;
        }
        set
        {
            corner = ((value % 4) + 4) % 4;
            UpdateOwner();
        }
    }

    bool isEnabled = true;

    /// <summary>The user can switch the mark off for a figure</summary>
    public bool IsEnabled
    {
        get
        {
            return isEnabled;
        }
        set
        {
            isEnabled = value;
            UpdateOwner();
        }
    }

    void UpdateOwner()
    {
        if (owner.Drawing != null)
        {
            owner.UpdateVisual();
        }
    }

    void ClickArea_PointerPressed(object sender, PointerPressedEventArgs e)
    {
        // only the drag tool: with any other, a click near the corner is that tool's business
        // (a point at the foot of the perpendicular, say)
        if (!CanClick || !e.GetCurrentPoint(clickArea).Properties.IsLeftButtonPressed)
        {
            return;
        }

        owner.Drawing.ActionManager.SetProperty(this, nameof(Corner), Corner + 1);
        e.Handled = true;
    }

    public void OnAddingToCanvas(Canvas canvas)
    {
        if (!canvas.Children.Contains(sign))
        {
            canvas.Children.Add(sign);
            canvas.Children.Add(clickArea);
        }
    }

    public void OnRemovingFromCanvas(Canvas canvas)
    {
        canvas.Children.Remove(sign);
        canvas.Children.Remove(clickArea);
    }

    public void Hide()
    {
        sign.IsVisible = false;
        clickArea.IsVisible = false;
    }

    /// <param name="vertex">Where the two lines meet, logical</param>
    /// <param name="baseLine">The line the owner is perpendicular to, logical; its direction
    /// is what <see cref="Corner"/> counts from</param>
    public void Show(Drawing drawing, Point vertex, PointPair baseLine)
    {
        var coordinateSystem = drawing.CoordinateSystem;
        var along = baseLine.P2.Minus(baseLine.P1);
        var across = new Point(-along.Y, along.X);
        switch (Corner)
        {
            case 1:
                along = along.Minus();
                break;
            case 2:
                along = along.Minus();
                across = across.Minus();
                break;
            case 3:
                across = across.Minus();
                break;
        }

        var cornerPoint = coordinateSystem.ToPhysical(vertex);
        var first = Direction(cornerPoint, coordinateSystem.ToPhysical(vertex.Plus(along)));
        var second = Direction(cornerPoint, coordinateSystem.ToPhysical(vertex.Plus(across)));
        if (!IsEnabled || first == null || second == null || HasAngleMeasuredAt(drawing, vertex))
        {
            Hide();
            return;
        }

        var points = GetPoints(cornerPoint, first.Value, second.Value, Size);
        sign.Points = points;
        clickArea.Points = new Points() { cornerPoint, points[0], points[1], points[2] };
        sign.IsVisible = true;
        clickArea.IsVisible = true;
    }

    /// <summary>
    /// The corner with the most room, to start in: towards the farther end of the base line
    /// and towards the given point across it.
    /// </summary>
    public static int GetRoomiestCorner(Point vertex, PointPair baseLine, Point pointAcross)
    {
        var along = baseLine.P2.Minus(baseLine.P1);
        var across = new Point(-along.Y, along.X);
        bool forward = vertex.Distance(baseLine.P2) >= vertex.Distance(baseLine.P1);
        var toPoint = pointAcross.Minus(vertex);
        bool left = toPoint.X * across.X + toPoint.Y * across.Y >= 0;
        if (left)
        {
            return forward ? 0 : 1;
        }

        return forward ? 3 : 2;
    }

    /// <summary>
    /// The three points of the sign: out along one side, the far corner of the square, back
    /// onto the other side. Also what an angle measurement of 90 degrees draws.
    /// </summary>
    public static Points GetPoints(Point corner, Point firstDirection, Point secondDirection, double size)
    {
        return new Points()
        {
            corner + firstDirection * size,
            corner + (firstDirection + secondDirection) * size,
            corner + secondDirection * size
        };
    }

    /// <summary>The unit vector from one physical point to another; null if they coincide</summary>
    public static Point? Direction(Point from, Point to)
    {
        var length = from.Distance(to);
        if (!(length > 1e-9) || double.IsInfinity(length))
        {
            return null;
        }

        return (to - from) / length;
    }

    static bool HasAngleMeasuredAt(Drawing drawing, Point vertex)
    {
        var tolerance = drawing.CoordinateSystem.CursorTolerance;
        return drawing.Figures
            .OfType<AngleArc>()
            .Any(a => a.Visible && a.Exists && a.Center.Distance(vertex) < tolerance);
    }

    /// <summary>Null if the file says nothing about the corner (it is older than the mark)</summary>
    public int? ReadXml(XElement element)
    {
        isEnabled = element.ReadBool(VisibleAttribute, true);
        var cornerAttribute = element.Attribute(CornerAttribute);
        int value;
        if (cornerAttribute != null && int.TryParse(cornerAttribute.Value, out value))
        {
            corner = ((value % 4) + 4) % 4;
            return corner;
        }

        return null;
    }

    public void WriteXml(XmlWriter writer)
    {
        writer.WriteAttributeString(CornerAttribute, Corner.ToString());
        if (!IsEnabled)
        {
            writer.WriteAttributeBool(VisibleAttribute, false);
        }
    }
}
