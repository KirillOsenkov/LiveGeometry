using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using DynamicGeometry;
using Drawing = DynamicGeometry.Drawing;
using Label = DynamicGeometry.Label;

namespace LiveGeometry;

/// <summary>
/// What the drawings of the gallery have in common: a heading and an explanation, which are
/// two labels named "Title" and "Description" that can't be clicked (so that they are never
/// in the way of dragging). tools/gallerize.cs writes them.
/// </summary>
public static class GalleryDrawing
{
    public const string TitleName = "Title";
    public const string DescriptionName = "Description";

    // between the figure and the text, and between the heading and the explanation
    const double gapPixels = 40;
    const double lineGapPixels = 6;
    const int fitRounds = 4;

    static readonly IBrush textPlate = new SolidColorBrush(Color.FromArgb(0xEB, 0xFF, 0xFF, 0xFF));

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
    /// wide window, under it in a tall one. A label has a size in pixels and a place in the
    /// coordinates of the drawing, so where "next to the figure" is depends on the zoom, which
    /// depends on how much room the text takes: a few rounds settle it.
    /// </summary>
    /// <param name="plane">See <see cref="GetPlane"/></param>
    public static void Fit(Drawing drawing, Rect? plane)
    {
        var coordinateSystem = drawing.CoordinateSystem;
        var title = drawing.Figures[TitleName] as Label;
        var description = drawing.Figures[DescriptionName] as Label;
        if (title == null || description == null)
        {
            coordinateSystem.ZoomExtend(plane);
            return;
        }

        // lines, graphs and the grid have no end: there is no place they can't reach
        foreach (var label in new[] { title, description })
        {
            if (label.Shape is Border plate)
            {
                plate.Background = textPlate;
                plate.Padding = new Thickness(10, 4, 10, 6);
                plate.CornerRadius = new CornerRadius(6);
            }
        }

        bool isWide =drawing.Canvas.Bounds.Width >= drawing.Canvas.Bounds.Height;
        for (int i = 0; i < fitRounds; i++)
        {
            bool hasFigure = coordinateSystem.TryGetContentBounds(out var figure, include: f => f != title && f != description);
            if (plane != null)
            {
                figure = hasFigure ? figure.Union(plane.Value) : plane.Value;
            }
            else if (!hasFigure)
            {
                break;
            }

            var titleSize = Measure(title);
            var descriptionSize = Measure(description);
            double gap = coordinateSystem.ToLogical(gapPixels);
            double titleHeight = coordinateSystem.ToLogical(titleSize.Height + lineGapPixels);
            double textHeight = titleHeight + coordinateSystem.ToLogical(descriptionSize.Height);

            // the y axis points up: Rect.Bottom is the top of the figure
            double x;
            double y;
            if (isWide)
            {
                x = figure.Right + gap;
                y = figure.Center.Y + System.Math.Max(figure.Height, textHeight) / 2;
            }
            else
            {
                x = figure.X;
                y = figure.Y - gap;
            }

            title.MoveTo(new Point(x, y));
            description.MoveTo(new Point(x, y - titleHeight));
            coordinateSystem.ZoomExtend(plane);
        }
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
