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
    public GalleryItem(string slug, string title, string fileName)
    {
        Slug = slug;
        Title = title;
        FileName = fileName;
    }

    /// <summary>The name in the url: /gallery/morley</summary>
    public string Slug { get; }

    public string Title { get; }

    public string FileName { get; }

    public string Path => GalleryCatalog.PathPrefix + Slug;

    /// <summary>See <see cref="GalleryDrawing.GetPlane"/></summary>
    public Avalonia.Rect? Plane => GalleryDrawing.GetPlane(LoadText());

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
        Item("bubbles", "Bubbles"),
        Item("pythagoras", "Pythagorean Theorem"),
        Item("circumscribed-circle", "Circumscribed Circle"),
        Item("inscribed-circle", "Inscribed Circle"),
        Item("morley", "Morley's Miracle"),
        Item("castle", "Castle"),
        Item("squares-around-rhombus", "Squares Around a Rhombus"),
        Item("carpenters-square", "Carpenter's Square"),
        Item("fibonacci-spiral", "Fibonacci Spiral"),
        Item("square-between-squares", "Square Between Squares"),
        Item("van-aubels-theorem", "Van Aubel's Theorem"),
        Item("desargues", "Desargues' Theorem"),
        Item("bezier", "Bézier Curve"),
        Item("triangle-on-3-lines", "Triangle on Three Lines", "TriangleOn3Lines"),
        Item("wireframe-cube", "Wireframe Cube"),
        Item("simson-line", "Simson Line"),
        Item("pentagon", "Regular Pentagon"),
        Item("ceva", "Ceva's Theorem"),
        Item("angles-in-a-circle", "Inscribed Angle"),
        Item("spiral", "Spiral"),
        Item("splitting-triangle", "Midsegments of a Triangle"),
        Item("parabola", "Parabola"),
        Item("ellipse-from-circle", "Ellipse from a Circle"),
        Item("sine-amplitude", "Sine Wave: Amplitude", "SinScaleY"),
        Item("sine-frequency", "Sine Wave: Frequency", "SinScaleX"),
        Item("pappus", "Pappus's Theorem"),
        Item("trapezoid", "A Trapezoid Surprise"),
        Item("quadrilateral-midpoints", "Varignon's Theorem"),
        Item("falling-ladder", "The Falling Ladder", "Ladder"),
        Item("square-in-square", "Square in a Square"),
        Item("composition-of-reflections", "Two Reflections"),
        Item("sierpinski", "Sierpinski Triangle"),
        Item("napoleons-theorem", "Napoleon's Theorem"),
        Item("parabola-graph", "Graph of a Parabola"),
        Item("reuleaux-triangle", "Reuleaux Triangle"),
        Item("cavalieri-principle", "Cavalieri's Principle"),
        Item("circle-tangents", "Tangents to a Circle"),
        Item("rose", "A Rose"),
        Item("steiners-problem", "Steiner's Problem"),
        Item("picks-theorem", "Pick's Theorem", "PickTheorem"),
        Item("ellipse-evolute", "Ellipse and Its Evolute"),
        Item("icosahedron", "Icosahedron"),
        Item("tetrahedron", "Tetrahedron"),
        Item("best-fit-circle", "Best-Fit Circle"),
        Item("measuring-distance", "Measuring Across a Lake"),
        Item("complex-numbers", "Complex Multiplication"),
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
    static GalleryItem Item(string slug, string title, string fileName = null)
    {
        fileName ??= string.Concat(slug.Split('-').Select(word => char.ToUpperInvariant(word[0]) + word.Substring(1)));
        return new GalleryItem(slug, title, fileName + ".lgf");
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
