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
/// two labels named "Title" and "Description", and maybe a third under them, "Hint", hidden
/// until a show/hide box brings it up (its place is kept). The caption is pinned to the screen
/// (<see cref="Label.Pin"/>), on a plate, and wraps to its column, so zooming and panning
/// leave it alone; <see cref="Fit"/> decides where it goes and how much room the figure gets.
/// It can be dragged like any pinned label: on a phone, where a long explanation runs off the
/// bottom, that is how the rest of it is read.
/// </summary>
public static class GalleryDrawing
{
    public const string TitleName = "Title";
    public const string DescriptionName = "Description";
    public const string HintName = "Hint";

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

    // the text never squeezes the figure below this share of the canvas: of its width beside
    // the caption, of its height above it - a third there, or on a phone the last line of
    // an explanation (Circumscribed Circle's question) was cut off at the bottom edge. A
    // drawing may ask for more under the caption (GalleryItem.StackedFigureShare)
    const double figureShare = 0.4;
    public const double StackedFigureShare = 1.0 / 3;

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
    /// <param name="stackedFigureShare">With the caption under the figure (a phone), the figure
    /// keeps at least this share of the height: see <see cref="GalleryItem.StackedFigureShare"/></param>
    public static void Fit(Drawing drawing, Rect? plane, double stackedFigureShare = StackedFigureShare)
    {
        var coordinateSystem = drawing.CoordinateSystem;
        double canvasWidth = drawing.Canvas.Bounds.Width;
        double canvasHeight = drawing.Canvas.Bounds.Height;

        // nothing to fit into (a window shrunk to its title bar); refitted when it grows
        if (canvasWidth <= 0 || canvasHeight <= 0)
        {
            return;
        }

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

        var hint = drawing.Figures[HintName] as Label;
        bool IsCaption(IFigure figure) => figure == title || figure == description || figure == hint;

        // the geometry a show/hide box brings up gets room from the start, or on a phone a
        // hint appears under the caption. Not its text: that is sized in pixels, so it has
        // no size in the plane to make room for (a long one only shrank the figure)
        var revealable = drawing.Figures
            .OfType<ShowHideControl>()
            .SelectMany(box => box.Dependencies)
            .Where(figure => figure is not ControlBase)
            .ToHashSet();
        bool TryGetFigure(out Rect bounds, bool withText = true)
        {
            bool hasFigure = coordinateSystem.TryGetContentBounds(
                out bounds,
                include: f => !IsCaption(f) && (withText || f is not ControlBase),
                includeHidden: revealable.Contains);
            if (plane != null)
            {
                bounds = hasFigure ? bounds.Union(plane.Value) : plane.Value;
                return true;
            }

            return hasFigure;
        }

        // The text of the figure (a point's name, a measurement) keeps its size in pixels, so
        // in the plane it is bigger the further the view is zoomed out: measured at whatever
        // zoom the view had, the fit depended on where the view came from - a window shrunk
        // to a sliver and grown back fitted the figure at half the size a fresh load did. So
        // it is measured at a zoom the window alone decides: the figure without its text
        // across the whole canvas. (Not round after round at the zoom the fit comes to, as
        // ZoomExtend does: that makes room for all of the text, and on a small screen the
        // figure shrank under its names.)
        if (!hasScene && TryGetFigure(out var bare, withText: false) && (bare.Width > 0 || bare.Height > 0))
        {
            double zoom = System.Math.Min(
                bare.Width > 0 ? canvasWidth / bare.Width : double.MaxValue,
                bare.Height > 0 ? canvasHeight / bare.Height : double.MaxValue);
            coordinateSystem.SetView(bare.Center, CoordinateSystem.ClampUnitLength(System.Math.Min(zoom, CoordinateSystem.MaxFitUnitLength)));
        }

        // a drawing with scenes shows the scene, not its content (ground goes on forever)
        Rect figure = default;
        if (!hasScene && !TryGetFigure(out figure))
        {
            coordinateSystem.ZoomExtend();
            return;
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

        // the hint is set off as another paragraph of the explanation would be, by about a line
        double hintGap = hint != null ? lineGapPixels + hint.TextBlock.FontSize : 0;
        Rect room;
        if (isWide)
        {
            SetCaption(LabelPin.TopRight, column, title, description, hint);
            var titleSize = title.MeasureSize();
            var descriptionSize = description.MeasureSize();
            double hintHeight = hint != null ? hintGap + hint.MeasureSize().Height : 0;
            double textHeight = titleSize.Height + lineGapPixels + descriptionSize.Height + hintHeight;
            double top =System.Math.Max(textMarginPixels, (canvasHeight - textHeight) / 2);
            title.PinOffset = new Point(textMarginPixels, top);
            description.PinOffset = new Point(textMarginPixels, top + titleSize.Height + lineGapPixels);
            if (hint != null)
            {
                hint.PinOffset = new Point(textMarginPixels, top + titleSize.Height + lineGapPixels + descriptionSize.Height + hintGap);
            }

            room = new Rect(
                margin,
                margin,
                canvasWidth - margin - gapPixels - column - textMarginPixels - margin,
                canvasHeight - 2 * margin);
        }
        else
        {
            SetCaption(LabelPin.BottomLeft, canvasWidth - 2 * textMarginPixels, title, description, hint);
            var titleSize = title.MeasureSize();
            var descriptionSize = description.MeasureSize();
            double hintHeight = hint != null ? hintGap + hint.MeasureSize().Height : 0;
            double textHeight = titleSize.Height + lineGapPixels + descriptionSize.Height + hintHeight;
            double roomHeight = System.Math.Max(
                canvasHeight - margin - gapPixels - textHeight - textMarginPixels,
                stackedFigureShare * canvasHeight);
            double textTop = margin + roomHeight + gapPixels;

            // pinned at the bottom: an offset is from there to the label's lower edge
            title.PinOffset = new Point(textMarginPixels, canvasHeight - textTop - titleSize.Height);
            description.PinOffset = new Point(textMarginPixels, canvasHeight - textTop - titleSize.Height - lineGapPixels - descriptionSize.Height);
            if (hint != null)
            {
                hint.PinOffset = new Point(textMarginPixels, canvasHeight - textTop - textHeight);
            }

            room = new Rect(margin, margin, canvasWidth - 2 * margin, roomHeight);
        }

        room = LayOutPinnedBoxes(drawing, room, figure, hasScene, margin);
        if (room.Width <= 0 || room.Height <= 0)
        {
            return;
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
            if (!hasScene)
            {
                // a big point (an emoji) would stick out of the room by half its size
                unitLength = coordinateSystem.LimitZoomByPointReach(
                    unitLength,
                    figure.Center,
                    room.Size,
                    include: f => !IsCaption(f));
            }

            unitLength = System.Math.Min(unitLength, CoordinateSystem.MaxFitUnitLength);
        }

        coordinateSystem.SetView(figure.Center, CoordinateSystem.ClampUnitLength(unitLength), room.Center);
    }

    /// <summary>
    /// The show/hide boxes pinned in the top left corner, laid out where they leave the figure
    /// the most room: one under another beside it, or in a row over it (on a phone in
    /// portrait a column took a third of the width). Returns the room left for the figure.
    /// </summary>
    static Rect LayOutPinnedBoxes(Drawing drawing, Rect room, Rect figure, bool hasScene, double margin)
    {
        var boxes = drawing.Figures
            .OfType<ShowHideControl>()
            .Where(box => box.Pin == LabelPin.TopLeft && box.Visible)
            .ToList();
        if (boxes.Count == 0)
        {
            return room;
        }

        var sizes = boxes.Select(box => box.MeasureSize()).ToList();
        double left = System.Math.Max(room.X, textMarginPixels + sizes.Max(size => size.Width) + margin);
        double top = System.Math.Max(room.Y, textMarginPixels + sizes.Max(size => size.Height) + margin);
        double rowWidth = sizes.Sum(size => size.Width) + boxRowGapPixels * (boxes.Count - 1);
        var beside = new Rect(left, room.Y, System.Math.Max(0, room.Right - left), room.Height);
        var under = new Rect(room.X, top, room.Width, System.Math.Max(0, room.Bottom - top));
        double Zoom(Rect candidate)
        {
            if (candidate.Width <= 0 || candidate.Height <= 0)
            {
                return 0;
            }

            var shown = hasScene ? drawing.ChooseScene(candidate.Width, candidate.Height).Value : figure;
            return System.Math.Min(candidate.Width / shown.Width, candidate.Height / shown.Height);
        }

        bool inRow = textMarginPixels + rowWidth <= room.Right && Zoom(under) > Zoom(beside);
        double x = textMarginPixels;
        double y = textMarginPixels;
        for (int i = 0; i < boxes.Count; i++)
        {
            boxes[i].PinOffset = new Point(x, y);
            boxes[i].UpdateVisual();
            if (inRow)
            {
                x += sizes[i].Width + boxRowGapPixels;
            }
            else
            {
                y += sizes[i].Height;
            }
        }

        return inRow ? under : beside;
    }

    // between two show/hide boxes in a row
    const double boxRowGapPixels = 16;

    static void SetCaption(LabelPin pin, double width, params Label[] labels)
    {
        foreach (var label in labels.Where(label => label != null))
        {
            label.Pin = pin;
            label.WrapWidth = width;
            label.Backdrop = true;
        }
    }

    /// <summary>
    /// A drawing about coordinates (a graph on the grid) has no bounds to fit: what it shows
    /// is the part of the plane its file says. Null for all other drawings.
    /// </summary>
    public static Rect? GetPlane(string drawingText)
    {
        return GetPlane(XElement.Parse(drawingText));
    }

    /// <summary>The same, of a drawing parsed already: parsing the whole file again for one attribute was a good part of loading a tile</summary>
    public static Rect? GetPlane(XElement drawing)
    {
        var viewport = drawing.Element("Viewport");
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
