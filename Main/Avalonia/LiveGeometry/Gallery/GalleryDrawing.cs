using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using Avalonia;
using DynamicGeometry;
using Drawing = DynamicGeometry.Drawing;
using Label = DynamicGeometry.Label;

namespace LiveGeometry;

/// <summary>
/// What the drawings of the gallery have in common: a caption - a heading and an explanation,
/// two labels named "Title" and "Description". The caption is pinned to the screen
/// (<see cref="Label.Pin"/>), on a plate, and wraps to its column, so zooming and panning
/// leave it alone; <see cref="Fit"/> decides where it goes and how much room the figure gets.
/// It can be dragged like any pinned label: on a phone, where a long explanation runs off the
/// bottom, that is how the rest of it is read.
/// </summary>
public static class GalleryDrawing
{
    public const string TitleName = "Title";
    public const string DescriptionName = "Description";

    /// <summary>The caption beside the figure is a column this wide (wider only for a heading that needs it)</summary>
    public const double CaptionWidth = 400;

    // a column narrower than this reads badly: the caption goes under the figure instead
    const double minimumCaptionWidth = 240;

    // the caption's plate from the edge of the canvas (the text is a padding further in);
    // the figure keeps CoordinateSystem.FitMarginPixels
    const double textMarginPixels = 16;

    // between the figure and the caption, and between the heading and the explanation: their
    // plates overlap by a padding, so the two texts are one padding apart
    const double gapPixels = 32;
    const double lineGapPixels = -Label.BackdropPadding;

    // the text never squeezes the figure below this share of the canvas
    const double figureShare = 0.4;

    /// <summary>
    /// For a thumbnail: text pinned to the screen is unreadable at that size and the geometry
    /// gets all the room
    /// </summary>
    public static void HideText(Drawing drawing)
    {
        foreach (var label in drawing.Figures.OfType<Label>().Where(label => label.Pin != LabelPin.None).ToArray())
        {
            label.Visible = false;
        }
    }

    public static bool HasCaption(Drawing drawing)
    {
        return FindCaption(drawing, out _, out _);
    }

    static bool FindCaption(Drawing drawing, out Label title, out Label description)
    {
        title = drawing.Figures[TitleName] as Label;
        description = drawing.Figures[DescriptionName] as Label;
        return title != null && description != null;
    }

    /// <summary>
    /// Zoom to fit, with the caption where it can't cover the figure: a column at the right in
    /// a wide window, a strip at the bottom in a tall one. The text has a size in pixels
    /// whatever the zoom, so the figure gets the canvas minus the text, and the zoom follows
    /// from that directly. In a window too small for both, the figure keeps a share of the
    /// canvas and the text runs off the bottom edge.
    /// </summary>
    /// <param name="plane">See <see cref="GetPlane"/></param>
    public static void Fit(Drawing drawing, Rect? plane)
    {
        var coordinateSystem = drawing.CoordinateSystem;
        double canvasWidth = drawing.Canvas.Bounds.Width;
        double canvasHeight = drawing.Canvas.Bounds.Height;
        bool hasScene = drawing.Scenes.Count > 0;
        if (!FindCaption(drawing, out var title, out var description))
        {
            if (hasScene)
            {
                drawing.ShowScene(drawing.ChooseScene(canvasWidth, canvasHeight).Value);
            }
            else
            {
                coordinateSystem.ZoomExtend(plane);
            }

            return;
        }

        // a drawing with scenes shows the scene, not its content (ground goes on forever)
        Rect figure = default;
        if (!hasScene)
        {
            bool hasFigure = coordinateSystem.TryGetContentBounds(out figure, include: f => f != title && f != description);
            if (plane != null)
            {
                figure = hasFigure ? figure.Union(plane.Value) : plane.Value;
            }
            else if (!hasFigure)
            {
                coordinateSystem.ZoomExtend();
                return;
            }
        }

        double margin = CoordinateSystem.FitMarginPixels;

        // beside the figure when there is room for a column, else under it
        bool isWide = canvasWidth >= canvasHeight;
        double column = 0;
        if (isWide)
        {
            title.WrapWidth = 0;
            title.Backdrop = true;
            double available = canvasWidth - textMarginPixels - gapPixels - margin - figureShare * canvasWidth;
            column = System.Math.Min(System.Math.Max(CaptionWidth, title.MeasureSize().Width), available);
            if (column < minimumCaptionWidth)
            {
                isWide = false;
            }
        }

        Rect room;
        if (isWide)
        {
            SetCaption(title, description, LabelPin.TopRight, column);
            var titleSize = title.MeasureSize();
            var descriptionSize = description.MeasureSize();
            double textHeight = titleSize.Height + lineGapPixels + descriptionSize.Height;
            double top = System.Math.Max(textMarginPixels, (canvasHeight - textHeight) / 2);
            title.PinOffset = new Point(textMarginPixels, top);
            description.PinOffset = new Point(textMarginPixels, top + titleSize.Height + lineGapPixels);
            room = new Rect(
                margin,
                margin,
                canvasWidth - margin - gapPixels - column - textMarginPixels - margin,
                canvasHeight - 2 * margin);
        }
        else
        {
            SetCaption(title, description, LabelPin.BottomLeft, canvasWidth - 2 * textMarginPixels);
            var titleSize = title.MeasureSize();
            var descriptionSize = description.MeasureSize();
            double textHeight = titleSize.Height + lineGapPixels + descriptionSize.Height;
            double roomHeight = System.Math.Max(
                canvasHeight - margin - gapPixels - textHeight - textMarginPixels,
                figureShare * canvasHeight);
            double textTop = margin + roomHeight + gapPixels;
            title.PinOffset = new Point(textMarginPixels, canvasHeight - textTop - titleSize.Height);
            description.PinOffset = new Point(textMarginPixels, canvasHeight - textTop - textHeight);
            room = new Rect(margin, margin, canvasWidth - 2 * margin, roomHeight);
        }

        // the scene nearest in shape to the room: landscape or portrait
        if (hasScene)
        {
            figure = drawing.ChooseScene(room.Width, room.Height).Value;
            drawing.ActiveScene = figure;
        }

        // the zoom that fills the room; a figure with no size keeps the zoom it has
        double unitLength = coordinateSystem.UnitLength;
        if (figure.Width > 0 || figure.Height > 0)
        {
            unitLength = System.Math.Min(
                figure.Width > 0 ? room.Width / figure.Width : double.MaxValue,
                figure.Height > 0 ? room.Height / figure.Height : double.MaxValue);
            unitLength = System.Math.Min(unitLength, CoordinateSystem.MaxFitUnitLength);
        }

        coordinateSystem.SetView(figure.Center, CoordinateSystem.ClampUnitLength(unitLength), room.Center);
    }

    static void SetCaption(Label title, Label description, LabelPin pin, double width)
    {
        title.Pin = pin;
        description.Pin = pin;
        title.WrapWidth = width;
        description.WrapWidth = width;
        title.Backdrop = true;
        description.Backdrop = true;
    }

    /// <summary>
    /// A drawing about coordinates (a graph on the grid) has no bounds to fit: what it shows
    /// is the part of the plane its file says. Null for all other drawings.
    /// </summary>
    public static Rect? GetPlane(string drawingText)
    {
        var viewport = XElement.Parse(drawingText).Element("Viewport");
        if (viewport == null || (string)viewport.Attribute("Grid") != "true")
        {
            return null;
        }

        double left = Read("Left");
        double right = Read("Right");
        double bottom = Read("Bottom");
        double top = Read("Top");
        return new Rect(left, bottom, right - left, top - bottom);

        double Read(string name) => double.Parse((string)viewport.Attribute(name), CultureInfo.InvariantCulture);
    }
}
