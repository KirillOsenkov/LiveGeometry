using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using Avalonia;
using DynamicGeometry;
using Drawing = DynamicGeometry.Drawing;
using Label = DynamicGeometry.Label;

namespace LiveGeometry;

/// <summary>
/// What the drawings of the gallery have in common: a heading and an explanation, which are
/// two labels named "Title" and "Description" that can't be clicked (so that they are never
/// in the way of dragging).
/// </summary>
public static class GalleryDrawing
{
    public const string TitleName = "Title";
    public const string DescriptionName = "Description";

    // between the figure and the text, and between the heading and the explanation
    const double gapPixels = 40;
    const double lineGapPixels = 6;

    // the text never squeezes the figure below this share of the canvas
    const double figureShare = 0.4;

    /// <summary>For a thumbnail: text is unreadable at that size and the geometry gets all the room</summary>
    public static void HideText(Drawing drawing)
    {
        foreach (var label in drawing.Figures.OfType<Label>().Where(label => !label.IsHitTestVisible).ToArray())
        {
            label.Visible = false;
        }
    }

    /// <summary>
    /// Zoom to fit, with the text put where it can't cover the figure: to the right of it in a
    /// wide window, under it in a tall one. The text has a size in pixels whatever the zoom,
    /// so the figure gets the canvas minus the text, and the zoom follows from that directly
    /// (iterating "place the text, zoom to fit" instead runs away once the text needs more
    /// than its share: every round zooms out a little more). In a window too small for both,
    /// the figure keeps a share of the canvas and the text runs off the edge.
    /// </summary>
    /// <param name="plane">See <see cref="GetPlane"/></param>
    public static void Fit(Drawing drawing, Rect? plane)
    {
        var coordinateSystem = drawing.CoordinateSystem;
        var title = drawing.Figures[TitleName] as Label;
        var description = drawing.Figures[DescriptionName] as Label;
        double canvasWidth = drawing.Canvas.Bounds.Width;
        double canvasHeight = drawing.Canvas.Bounds.Height;
        bool hasScene = drawing.Scenes.Count > 0;
        if (title == null || description == null)
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

        // everything in pixels first
        var titleSize = Measure(title);
        var descriptionSize = Measure(description);
        double textWidth = System.Math.Max(titleSize.Width, descriptionSize.Width);
        double titleHeight = titleSize.Height + lineGapPixels;
        double textHeight = titleHeight + descriptionSize.Height;
        double margin = CoordinateSystem.FitMarginPixels;

        bool isWide = canvasWidth >= canvasHeight;
        double roomWidth = 0;
        double roomHeight = 0;
        if (isWide)
        {
            roomWidth = canvasWidth - 2 * margin - gapPixels - textWidth;
            roomHeight = canvasHeight - 2 * margin;
            if (roomWidth < figureShare * canvasWidth)
            {
                // the text would leave the figure a sliver: under it instead
                isWide = false;
            }
        }

        if (!isWide)
        {
            roomWidth = canvasWidth - 2 * margin;
            roomHeight = System.Math.Max(canvasHeight - 2 * margin - gapPixels - textHeight, figureShare * canvasHeight);
        }

        // the scene nearest in shape to the room: landscape or portrait
        if (hasScene)
        {
            figure = drawing.ChooseScene(roomWidth, roomHeight).Value;
            drawing.ActiveScene = figure;
        }

        // the zoom that fills the room; a figure with no size keeps the zoom it has
        double unitLength = coordinateSystem.UnitLength;
        if (figure.Width > 0 || figure.Height > 0)
        {
            unitLength = System.Math.Min(
                figure.Width > 0 ? roomWidth / figure.Width : double.MaxValue,
                figure.Height > 0 ? roomHeight / figure.Height : double.MaxValue);
            unitLength = System.Math.Min(unitLength, CoordinateSystem.MaxFitUnitLength);
        }

        unitLength = CoordinateSystem.ClampUnitLength(unitLength);

        // now in logical units: where the text goes, and the box around figure and text that
        // is to sit in the middle of the canvas (the y axis points up: Rect.Bottom is the top)
        double gap = gapPixels / unitLength;
        Rect block;
        Point textTopLeft;
        if (isWide)
        {
            double blockHeight = System.Math.Max(figure.Height, textHeight / unitLength);
            double top = figure.Center.Y + blockHeight / 2;
            textTopLeft = new Point(figure.Right + gap, top);
            block = new Rect(figure.X, top - blockHeight, figure.Width + gap + textWidth / unitLength, blockHeight);
        }
        else
        {
            textTopLeft = new Point(figure.X, figure.Y - gap);
            double bottom = textTopLeft.Y - textHeight / unitLength;
            double blockWidth = System.Math.Max(figure.Width, textWidth / unitLength);
            block = new Rect(figure.X, bottom, blockWidth, figure.Bottom - bottom);
        }

        title.MoveTo(textTopLeft);
        description.MoveTo(new Point(textTopLeft.X, textTopLeft.Y - titleHeight / unitLength));

        // centered - unless the block is too big for the canvas: then it starts at the margin,
        // so that the figure stays in view and it is the end of the text that runs off
        var center = block.Center;
        if (block.Width * unitLength > canvasWidth - 2 * margin)
        {
            center = center.WithX(block.X + (canvasWidth / 2 - margin) / unitLength);
        }

        if (block.Height * unitLength > canvasHeight - 2 * margin)
        {
            center = center.WithY(block.Bottom - (canvasHeight / 2 - margin) / unitLength);
        }

        coordinateSystem.SetView(center, unitLength);
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

    static Size Measure(Label label)
    {
        label.Shape.Measure(Size.Infinity);
        return label.Shape.DesiredSize;
    }
}
