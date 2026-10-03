using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;

namespace LiveGeometry;

/// <summary>
/// A drawing of the gallery: an .lgf file embedded in this assembly (Gallery/Drawings).
/// </summary>
public class GalleryItem
{
    public GalleryItem(string slug, string title, string fileName, double stackedFigureShare)
    {
        Slug = slug;
        Title = title;
        FileName = fileName;
        StackedFigureShare = stackedFigureShare;
    }

    /// <summary>
    /// With the caption under the figure (a phone in portrait), the figure keeps at least this
    /// share of the canvas's height, and what doesn't fit of the caption runs off the bottom
    /// (it is dragged up to be read). More than the usual third for a picture that is the
    /// point of the drawing.
    /// </summary>
    public double StackedFigureShare { get; }

    /// <summary>The name in the url: /gallery/morley</summary>
    public string Slug { get; }

    public string Title { get; }

    public string FileName { get; }

    public string Path => GalleryCatalog.PathPrefix + Slug;

    Avalonia.Rect? plane;
    bool isPlaneRead;

    /// <summary>See <see cref="GalleryDrawing.GetPlane(string)"/>; read once, it is asked on every resize of the window</summary>
    public Avalonia.Rect? Plane
    {
        get
        {
            if (!isPlaneRead)
            {
                plane = GalleryDrawing.GetPlane(LoadText());
                isPlaneRead = true;
            }

            return plane;
        }
    }

    /// <summary>Some of its points are characters (an emoji, a ★), drawn in the emoji font (<see cref="DynamicGeometry.EmojiFont"/>)</summary>
    public bool UsesEmoji => LoadText().Contains(" Character=\"", StringComparison.Ordinal);

    string text;

    public string LoadText()
    {
        if (text == null)
        {
            var assembly = typeof(GalleryItem).Assembly;
            using var stream = assembly.GetManifestResourceStream("LiveGeometry.Gallery.Drawings." + FileName);
            using var reader = new StreamReader(stream);
            text = reader.ReadToEnd();
        }

        return text;
    }
}

/// <summary>
/// The drawings of the gallery, in the order of the tour (previous / next).
/// </summary>
public static class GalleryCatalog
{
    public const string PathPrefix = "/gallery/";

    public static IReadOnlyList<GalleryItem> Items { get; } = new[]
    {
        Item("continuous-deformations", "Continuous Deformations"),
        Item("magic-tree", "Magic Tree", stackedFigureShare: 0.5),
        Item("bubbles", "Bubbles"),
        Item("treasure-island", "Treasure Island", stackedFigureShare: 0.5),
        Item("pythagoras", "Pythagorean Theorem"),
        Item("circumscribed-circle", "Circumscribed Circle"),
        Item("aperiodic-monotile", "Aperiodic Monotile"),
        Item("inscribed-circle", "Inscribed Circle"),
        Item("circle-touching-three-lines", "Circle Touching Three Lines"),
        Item("morley", "Morley's Miracle"),
        Item("castle", "Castle"),
        Item("squares-around-rhombus", "Squares Around a Rhombus"),
        Item("carpenters-square", "Carpenter's Square"),
        Item("fibonacci-spiral", "Fibonacci Spiral"),
        Item("square-between-squares", "Square Between Squares"),
        Item("van-aubels-theorem", "Van Aubel's Theorem"),
        Item("bezier", "Bézier Curve"),
        Item("triangle-on-3-lines", "Triangle on Three Lines", "TriangleOn3Lines"),
        Item("simson-line", "Simson Line"),
        Item("steiners-problem", "Steiner's Problem"),
        Item("pentagon", "Regular Pentagon"),
        Item("catenary", "Catenary"),
        Item("circle-inversion", "Inversion in a Circle"),
        Item("hat-kites", "Eight Kites Make a Hat"),
        Item("hat-family", "The Hat Family"),
        Item("angles-in-a-circle", "Inscribed Angle"),
        Item("fireworks", "Fireworks"),
        Item("circle-tangents", "Tangents to a Circle"),
        Item("splitting-triangle", "Midsegments of a Triangle"),
        Item("parabola", "Parabola"),
        Item("ellipse-from-circle", "Ellipse from a Circle"),
        Item("wireframe-cube", "Wireframe Cube"),
        Item("sine-wave", "Sine Wave"),
        Item("spiral", "Spiral"),
        Item("pappus", "Pappus's Theorem"),
        Item("trapezoid", "A Trapezoid Surprise"),
        Item("quadrilateral-midpoints", "Varignon's Theorem"),
        Item("falling-ladder", "The Falling Ladder", "Ladder"),
        Item("square-in-square", "Square in a Square"),
        Item("composition-of-reflections", "Two Reflections"),
        Item("sierpinski", "Sierpinski Triangle"),
        Item("napoleons-theorem", "Napoleon's Theorem"),
        Item("golden-angle", "Golden Angle"),
        Item("platonic-solids", "The Five Platonic Solids", "PlatonicSolids"),
        Item("parabola-graph", "Graph of a Parabola"),
        Item("reuleaux-triangle", "Reuleaux Triangle"),
        Item("cavalieri-principle", "Cavalieri's Principle"),
        Item("rose", "A Rose"),
        Item("picks-theorem", "Pick's Theorem", "PickTheorem"),
        Item("kaleidoscope", "Kaleidoscope"),
        Item("desargues", "Desargues' Theorem"),
        Item("ellipse-evolute", "Ellipse and Its Evolute"),
        Item("line-of-best-fit", "Line of Best Fit"),
        Item("measuring-distance", "Measuring Across a Lake"),
        Item("complex-numbers", "Complex Multiplication"),
        Item("ceva", "Ceva's Theorem"),
        Item("conic-through-five-points", "Conic Through Five Points", "Pascal"),
    };

    /// <summary>
    /// Rewrites the Items block of this source file in the given order, for the gallery's
    /// arrange mode: every drawing keeps its own line, comments and blank lines between them
    /// go. Returns the path written. Desktop only: it needs the source tree.
    /// </summary>
    public static string SaveOrder(IEnumerable<GalleryItem> order)
    {
        var sourcePath = SourcePath();
        var text = File.ReadAllText(sourcePath);
        var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None).ToList();
        int start = lines.FindIndex(line => line.Contains("Items { get; } = new[]")) + 2;
        int end = lines.FindIndex(start, line => line.Trim() == "};");
        var itemLines = lines
            .Skip(start)
            .Take(end - start)
            .Select(line => (line, match: Regex.Match(line, "Item\\(\"([^\"]+)\"")))
            .Where(pair => pair.match.Success)
            .ToDictionary(pair => pair.match.Groups[1].Value, pair => pair.line);
        var block = order.Select(item => itemLines[item.Slug]).ToList();
        lines.RemoveRange(start, end - start);
        lines.InsertRange(start, block);
        File.WriteAllText(sourcePath, string.Join("\r\n", lines));
        return sourcePath;
    }

    // this file: the attribute names the file of the call, which is why it is called from here
    static string SourcePath([CallerFilePath] string path = null)
    {
        return path;
    }

    /// <param name="fileName">Without extension; by default the slug in PascalCase</param>
    /// <param name="stackedFigureShare">See <see cref="GalleryItem.StackedFigureShare"/></param>
    static GalleryItem Item(
        string slug,
        string title,
        string fileName = null,
        double stackedFigureShare = GalleryDrawing.StackedFigureShare)
    {
        fileName ??= string.Concat(slug.Split('-').Select(word => char.ToUpperInvariant(word[0]) + word.Substring(1)));
        return new GalleryItem(slug, title, fileName + ".lgf", stackedFigureShare);
    }

    public static GalleryItem FindBySlug(string slug)
    {
        return Items.FirstOrDefault(item => string.Equals(item.Slug, slug, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>The drawing a path like /gallery/morley names, or null</summary>
    public static GalleryItem FindByPath(string path)
    {
        if (path == null || !path.StartsWith(PathPrefix, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return FindBySlug(path.Substring(PathPrefix.Length).Trim('/'));
    }

    public static int IndexOf(GalleryItem item)
    {
        for (int i = 0; i < Items.Count; i++)
        {
            if (Items[i] == item)
            {
                return i;
            }
        }

        return -1;
    }
}
