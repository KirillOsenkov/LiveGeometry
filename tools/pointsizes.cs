#:property TargetFramework=net10.0
#:property Nullable=disable
#:property PublishAot=false

// pointsizes - lists the point styles of the .lgf files in a folder that are smaller than the
// standard sizes (10 for a point that can be dragged - FreePoint, PointOnFigure - 8 for any
// other), with the kinds of point that use each, and with --apply raises them to the standard.
//
//   dotnet tools/pointsizes.cs -- <folder of .lgf> [--apply]

using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;

if (args.Length < 1)
{
    Console.WriteLine("usage: pointsizes <folder of .lgf> [--apply]");
    return 1;
}

bool apply = args.Contains("--apply");
var draggable = new HashSet<string> { "FreePoint", "PointOnFigure" };
const double draggableSize = 10;
const double dependentSize = 8;

foreach (var file in Directory.GetFiles(args[0], "*.lgf").OrderBy(f => f))
{
    var document = XDocument.Load(file, LoadOptions.PreserveWhitespace);
    var figures = document.Root.Element("Figures")?.Elements() ?? Enumerable.Empty<XElement>();
    var usersByStyle = figures
        .Where(f => f.Attribute("Style") != null)
        .GroupBy(f => (string)f.Attribute("Style"))
        .ToDictionary(g => g.Key, g => g.Select(f => f.Name.LocalName).Distinct().OrderBy(n => n).ToArray());

    bool changed = false;
    foreach (var style in document.Root.Element("Styles").Elements("PointStyle"))
    {
        string name = (string)style.Attribute("Name");
        double size = double.Parse((string)style.Attribute("Size"), CultureInfo.InvariantCulture);
        usersByStyle.TryGetValue(name, out var users);
        users ??= Array.Empty<string>();
        bool isDraggable = users.Any(draggable.Contains);
        double standard = isDraggable ? draggableSize : dependentSize;
        if (size >= standard || users.Length == 0)
        {
            continue;
        }

        Console.WriteLine($"{Path.GetFileName(file)}: style {name} size {size} -> {standard} ({string.Join(",", users)})");
        if (apply)
        {
            style.SetAttributeValue("Size", standard.ToString(CultureInfo.InvariantCulture));
            changed = true;
        }
    }

    if (changed)
    {
        var settings = new XmlWriterSettings() { Indent = true, Encoding = new UTF8Encoding(false), NewLineChars = "\r\n" };
        using var writer = XmlWriter.Create(file, settings);
        document.Save(writer);
    }
}

return 0;
