#:property TargetFramework=net10.0
#:property Nullable=disable
#:property PublishAot=false

// stylecensus - the distinct styles the .lgf files of a folder carry, by their values (the
// name left out), with how many files carry each and under what names: what the old
// numbered copies of the defaults look like, to map them onto the defaults by signature.
//
//   dotnet tools/stylecensus.cs -- <folder of .lgf> [--min 2]

using System.Xml.Linq;

if (args.Length < 1)
{
    Console.WriteLine("usage: stylecensus <folder of .lgf> [--min 2]");
    return 1;
}

int minimum = 1;
int minimumIndex = Array.IndexOf(args, "--min");
if (minimumIndex >= 0 && minimumIndex + 1 < args.Length)
{
    minimum = int.Parse(args[minimumIndex + 1]);
}

var census = new Dictionary<string, (HashSet<string> Files, HashSet<string> Names, int Uses)>();
foreach (var file in Directory.GetFiles(args[0], "*.lgf").OrderBy(f => f))
{
    var document = XDocument.Load(file);
    var figures = document.Root.Element("Figures")?.Descendants() ?? Enumerable.Empty<XElement>();
    var usesByStyle = figures
        .Where(f => f.Attribute("Style") != null)
        .GroupBy(f => (string)f.Attribute("Style"))
        .ToDictionary(g => g.Key, g => g.Count());

    foreach (var style in document.Root.Element("Styles")?.Elements() ?? Enumerable.Empty<XElement>())
    {
        string name = (string)style.Attribute("Name");
        var attributes = style.Attributes()
            .Where(a => a.Name.LocalName != "Name")
            .Select(a => a.Name.LocalName + "=" + a.Value);
        string signature = style.Name.LocalName + " " + string.Join(" ", attributes);
        if (style.HasElements)
        {
            signature += " " + string.Concat(style.Elements().Select(e => e.ToString(SaveOptions.DisableFormatting)));
        }

        if (!census.TryGetValue(signature, out var entry))
        {
            entry = (new HashSet<string>(), new HashSet<string>(), 0);
            census[signature] = entry;
        }

        entry.Files.Add(Path.GetFileNameWithoutExtension(file));
        entry.Names.Add(name);
        entry.Uses += usesByStyle.TryGetValue(name, out int uses) ? uses : 0;
        census[signature] = entry;
    }
}

foreach (var pair in census.Where(p => p.Value.Files.Count >= minimum).OrderByDescending(p => p.Value.Files.Count))
{
    var entry = pair.Value;
    Console.WriteLine($"{entry.Files.Count} files, {entry.Uses} uses, names {string.Join(",", entry.Names.OrderBy(n => n))}");
    Console.WriteLine("    " + pair.Key);
}

return 0;
