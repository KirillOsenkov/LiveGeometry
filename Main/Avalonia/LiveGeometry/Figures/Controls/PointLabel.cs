using Avalonia;

namespace LiveGeometry;

public class PointLabel : NameLabel
{
    PointBase GetPoint()
    {
        return (PointBase)Figure;
    }

    protected override string Kind => "Label";

    /// <summary>
    /// A hidden point keeps its name: the name is then a text that moves with the point, as the
    /// numbers of a clock face, the Fibonacci squares' sizes and the Platonic solids' names are
    /// </summary>
    protected override bool HidesWithFigure => false;

    /// <summary>
    /// The least room between the point's rim and the letters, in pixels: none, the name
    /// may touch the point's outline
    /// </summary>
    public const double Clearance = 0;

    /// <summary>
    /// In pixels: what the orbit and the first place of a label add to the size of the
    /// text and the point. A number of its own: it was the cursor's tolerance, which a
    /// finger has more of than a mouse, and a label made by touch would sit elsewhere.
    /// </summary>
    public const double Margin = 5;

    /// <summary>
    /// Keeps the label in orbit around its point: the center of the label may not get
    /// further from the point than the label's larger dimension plus the point radius
    /// (a long label - name and coordinates - gets a proportionally bigger orbit), and
    /// no part of the letters may come closer to the point than <see cref="Clearance"/>.
    /// </summary>
    /// <param name="newPosition">Desired top-left corner of the label, logical</param>
    public Point ClampPosition(Point newPosition)
    {
        var width = selection.Bounds.Width;
        var height = selection.Bounds.Height;
        if (width == 0 || height == 0)
        {
            // not measured yet (e.g. while loading): nothing to base the radius on
            return newPosition;
        }

        // in pixels: the orbit is about the size of the text, whatever the zoom
        var point = ToPhysical(Anchor);
        var halfSize = new Point(width / 2, height / 2);
        var pointRadius = GetPoint().Shape.Bounds.Width / 2;
        var radius = System.Math.Max(width, height) + pointRadius + Margin;
        var fromPoint = ToPhysical(newPosition).Plus(halfSize).Minus(point);
        fromPoint = fromPoint.TrimToMaxLength(radius);

        // the letters are kept clear, not the line box around them
        var ink = GetInkBounds(width, height);
        var inkShift = ink.Center.Minus(halfSize);
        var inkHalfSize = new Point(ink.Width / 2, ink.Height / 2);
        fromPoint = PushClear(fromPoint.Plus(inkShift), inkHalfSize, pointRadius + Clearance).Minus(inkShift);
        return ToLogical(point.Plus(fromPoint).Minus(halfSize));
    }

    /// <summary>
    /// Where the letters are in the label, in pixels from its top left corner: the line box
    /// has room above capitals and below the baseline, and a little at the sides. The whole
    /// label if the text doesn't say (more than one line, nothing measured yet).
    /// </summary>
    Rect GetInkBounds(double width, double height)
    {
        var whole = new Rect(0, 0, width, height);
        var lines = TextBlock.TextLayout?.TextLines;
        if (lines == null || lines.Count != 1)
        {
            return whole;
        }

        var line = lines[0];
        var bottom = line.Height + line.OverhangAfter;
        var left = line.Start + line.OverhangLeading;
        var right = line.Start + line.WidthIncludingTrailingWhitespace - line.OverhangTrailing;
        var ink = new Rect(left, bottom - line.Extent, right - left, line.Extent)
            .Translate(new Avalonia.Vector(TextBlock.Padding.Left, TextBlock.Padding.Top));
        if (!(ink.Width > 0 && ink.Height > 0) || !whole.Contains(ink.Center))
        {
            return whole;
        }

        return ink;
    }

    /// <summary>
    /// Moves the center of a box out along its own direction from the origin until no part
    /// of the box is nearer to the origin than the minimum.
    /// </summary>
    static Point PushClear(Point center, Point halfSize, double minimum)
    {
        if (center.X == 0 && center.Y == 0)
        {
            // right on the point: above it, where a new label goes
            center = new Point(0, -1);
        }

        if (Gap(center, halfSize) >= minimum)
        {
            return center;
        }

        // the gap grows with the scale: bisect between "as it is" and "far enough"
        double low = 1;
        double high = 2;
        while (Gap(center.Scale(high), halfSize) < minimum)
        {
            high *= 2;
        }

        for (int i = 0; i < 20; i++)
        {
            double middle = (low + high) / 2;
            if (Gap(center.Scale(middle), halfSize) < minimum)
            {
                low = middle;
            }
            else
            {
                high = middle;
            }
        }

        return center.Scale(high);
    }

    /// <summary>The distance from the origin to the nearest point of a box</summary>
    static double Gap(Point center, Point halfSize)
    {
        double dx = System.Math.Max(System.Math.Abs(center.X) - halfSize.X, 0);
        double dy = System.Math.Max(System.Math.Abs(center.Y) - halfSize.Y, 0);
        return System.Math.Sqrt(dx * dx + dy * dy);
    }

    /// <summary>
    /// Puts the label where the orbit rule allows, starting from where it is: for labels
    /// from files that were placed on top of their point.
    /// </summary>
    public void SpaceFromPoint()
    {
        MoveTo(Coordinates);
    }

    public override void MoveToCore(Point newPosition)
    {
        base.MoveToCore(ClampPosition(newPosition));
    }

    /// <summary>
    /// A new label: centered above its point, just clear of it (or the nearest free place to
    /// that, not on a vertical line through the point: <see cref="LabelWithOffset.FirstOffset"/>)
    /// </summary>
    protected override Point? DefaultOffset()
    {
        var size = MeasureSize();
        return new Point(-size.Width / 2, -(size.Height + GetPoint().Shape.ActualHeight / 2 + Margin));
    }

    public override void UpdateVisual()
    {
        if (Dependencies.IsEmpty())
        {
            return;
        }

        UpdateText();
        base.UpdateVisual();
    }

    private void UpdateText()
    {
        var text = "";
        if (ShowName)
        {
            // A_1 shows as A₁
            text = NameDisplay.Format(Figure.Name);
        }
        if (ShowCoordinates)
        {
            var coordinates = Point(0);
            var x = Math.Round(coordinates.X, DecimalsToShow);
            var y = Math.Round(coordinates.Y, DecimalsToShow);
            // as a textbook writes them (it was "(3;4)", the way of the countries that
            // write a decimal comma; numbers here have a point)
            var coordinatesText = string.Format("({0}, {1})", x, y);
            if (!text.IsEmpty())
            {
                text += " ";
            }
            text += coordinatesText;
        }
        Text = text;
    }

    bool showName;
    [PropertyGridVisible(false)]    // Handled in the point's property grid.
    public bool ShowName
    {
        get => showName;
        set
        {
            showName = value;
            UpdateVisual();
        }
    }

    bool showCoordinates;
    [PropertyGridVisible(false)]    // Handled in the point's property grid.
    public bool ShowCoordinates
    {
        get => showCoordinates;
        set
        {
            showCoordinates = value;
            UpdateVisual();
        }
    }

    public override void ReadXml(System.Xml.Linq.XElement element)
    {
        base.ReadXml(element);
        this.showName = element.ReadBool("ShowName", true);
        this.showCoordinates = element.ReadBool("ShowCoordinates", false);
        UpdateText();
    }

    public override void WriteXml(System.Xml.XmlWriter writer)
    {
        base.WriteXml(writer);

        // only where they differ from what ReadXml takes without them
        if (!ShowName)
        {
            writer.WriteAttributeBool("ShowName", false);
        }

        if (ShowCoordinates)
        {
            writer.WriteAttributeBool("ShowCoordinates", true);
        }
    }

    [PropertyGridVisible]
    [PropertyGridName("Show name")]
    [PropertyGridSection(PropertyGridSection.Display, 10)]
    public bool ShowNameDisplay
    {
        get => ShowName;
        set
        {
            if (Figure is PointBase point)
            {
                point.ShowName = value;
            }
        }
    }

    [PropertyGridVisible]
    [PropertyGridName("Show coordinates")]
    [PropertyGridSection(PropertyGridSection.Display, 20)]
    public bool ShowCoordinatesDisplay
    {
        get => ShowCoordinates;
        set
        {
            if (Figure is PointBase point)
            {
                point.ShowCoordinates = value;
            }
        }
    }
}
