// Builds the catalog of the Catriona Agg library (CatrionaAgg/puzzles.json, the /CatrionaAgg
// pages read it) from Andrew Stacey's index of her puzzles on notes.mathforge.org: the puzzles
// by year (title, date, thumbnail), the puzzles by topic (the techniques and their categories)
// and, for every puzzle that has a drawing here or is among the --recent ones, its own page
// (the title as written, the question as Catriona posted it, the BlueSky thread).
//
//   dotnet run tools/catriona.cs -- [--recent 20] [--images <folder>] [--cache <folder>] [--refresh]
//
// Pages are kept in the cache folder (default: livegeometry-mathforge in the temp folder) and
// fetched again only with --refresh, so a rerun after adding drawings asks mathforge for
// nothing new but the pages of the puzzles drawn since. --images downloads the full pictures
// of the puzzles with details into a folder, to draw from; they are not part of the site.
// What a drawing is called: CatrionaAgg/drawings/<slug>.lgf.

#nullable disable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

const string Site = "https://notes.mathforge.org";
string repository = FindRepository();
string libraryFolder = Path.Combine(repository, "CatrionaAgg");
string drawingsFolder = Path.Combine(libraryFolder, "drawings");
string catalogFile = Path.Combine(libraryFolder, "puzzles.json");
int recent = int.TryParse(Option("--recent"), out int n) ? n : 0;
string imagesFolder = Option("--images");
string cacheFolder = Option("--cache") ?? Path.Combine(Path.GetTempPath(), "livegeometry-mathforge");
bool refresh = args.Contains("--refresh");
Directory.CreateDirectory(cacheFolder);

using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
http.DefaultRequestHeaders.UserAgent.ParseAdd("LiveGeometry-catalog/1.0 (+https://livegeometry.com/CatrionaAgg)");

// ---------- the puzzles by year ----------

var home = await Source("HomePage");
var years = Regex.Matches(home, @"\[\[Catriona Agg's puzzles (\d{4})\|?[^\]]*\]\]|\[\[Catriona Agg's puzzles (\d{4})\]\]")
    .Select(m => m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value)
    .Distinct()
    .OrderByDescending(y => y)
    .ToList();
if (years.Count == 0)
{
    // the home page names the years as links ([[Catriona Agg's puzzles 2026|2026]]); read them off the published page if the source says them otherwise
    years = Enumerable.Range(2018, DateTime.UtcNow.Year - 2017).Select(y => y.ToString()).Reverse().ToList();
}

var puzzles = new List<Puzzle>();
foreach (var year in years)
{
    string text = await Source("Catriona Agg's puzzles " + year);
    if (text == null)
    {
        continue;
    }

    foreach (Match match in Regex.Matches(text, @"1\.\s+\[\[([^\]|]+)(?:\|[^\]]*)?\]\]\s*_([^_]+)_\s*(?:\+--[^\n]*\n\s*\[\[([^\]:]+):pic\]\])?"))
    {
        var puzzle = new Puzzle
        {
            Name = match.Groups[1].Value.Trim(),
            Date = ParseDate(match.Groups[2].Value, int.Parse(year)),
            Thumbnail = match.Groups[3].Success ? match.Groups[3].Value.Trim() : null,
        };
        if (!puzzles.Any(p => Same(p.Name, puzzle.Name)))
        {
            puzzles.Add(puzzle);
        }
    }
}

// newest first; one day's puzzles keep the year page's order
puzzles = puzzles.Select((p, i) => (p, i)).OrderByDescending(x => x.p.Date).ThenBy(x => x.i).Select(x => x.p).ToList();
Console.WriteLine($"{puzzles.Count} puzzles in {years.Count} years");

// ---------- the puzzles by topic: a category (## ...), its techniques (* [[...]]) and the
// puzzles listed under each (<a href="name">), or under the category alone ("Various ...",
// "With previous categories") ----------

string topics = await Source("Puzzles by Topic");
string category = null;
string technique = null;
var categoryOfTechnique = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
foreach (var line in topics.Split('\n'))
{
    var heading = Regex.Match(line, @"^##\s+(.+?)\s*$");
    if (heading.Success)
    {
        category = heading.Groups[1].Value;
        technique = null;
        continue;
    }

    var item = Regex.Match(line, @"^\*\s+(?:\[\[([^\]|]+)(?:\|[^\]]*)?\]\]|(.+?))\s*$");
    if (item.Success)
    {
        technique = item.Groups[1].Success ? item.Groups[1].Value.Trim() : null;
        if (technique != null && category != null)
        {
            categoryOfTechnique.TryAdd(technique, category);
        }

        continue;
    }

    if (category == null)
    {
        continue;
    }

    foreach (Match link in Regex.Matches(line, "<a href=\"([^\"]+)\""))
    {
        var puzzle = puzzles.FirstOrDefault(p => Same(p.Name, link.Groups[1].Value));
        if (puzzle == null)
        {
            continue;
        }

        if (!puzzle.Categories.Contains(category))
        {
            puzzle.Categories.Add(category);
        }

        if (technique != null && !puzzle.Techniques.Contains(technique, StringComparer.OrdinalIgnoreCase))
        {
            puzzle.Techniques.Add(technique);
        }
    }
}

// ---------- the pages of the puzzles drawn here, and of the most recent ----------

var drawn = Directory.Exists(drawingsFolder)
    ? Directory.GetFiles(drawingsFolder, "*.lgf").Select(Path.GetFileNameWithoutExtension).ToHashSet(StringComparer.OrdinalIgnoreCase)
    : new HashSet<string>();
var detailed = puzzles.Where((p, i) => i < recent || drawn.Contains(p.Slug)).ToList();
foreach (var puzzle in detailed)
{
    string text = await Source(puzzle.Name);
    if (text == null)
    {
        Console.WriteLine($"warning: no page for {puzzle.Name}");
        continue;
    }

    puzzle.Title = Regex.Match(text, @"^#\s+(.+?)\s*$", RegexOptions.Multiline) is { Success: true } title ? title.Groups[1].Value : null;
    puzzle.Image = Regex.Match(text, @"\[\[([^\]:]+\.(?:png|jpe?g)):pic\]\]", RegexOptions.IgnoreCase) is { Success: true } image ? image.Groups[1].Value : null;
    var quote = Regex.Matches(text, @"^>\s?(.*)$", RegexOptions.Multiline).Select(m => m.Groups[1].Value.Trim()).Where(s => s.Length > 0);
    puzzle.Statement = quote.Any() ? string.Join(" ", quote) : null;
    puzzle.Thread = Regex.Match(text, @"\((https://bsky\.app/[^)\s]+)\)") is { Success: true } thread ? thread.Groups[1].Value : null;

    // a puzzle newer than the last update of the topic index: its own techniques page says them
    if (puzzle.Categories.Count == 0 && await Source(puzzle.Name + " techniques") is string techniques)
    {
        foreach (Match link in Regex.Matches(techniques, @"\[\[([^\]|]+)(?:\|[^\]]*)?\]\]"))
        {
            string name = link.Groups[1].Value.Trim();
            if (Same(name, puzzle.Name))
            {
                continue;
            }

            if (categoryOfTechnique.TryGetValue(name, out var found) || categoryOfTechnique.TryGetValue(Regex.Replace(name, "^regular ", ""), out found))
            {
                if (!puzzle.Categories.Contains(found))
                {
                    puzzle.Categories.Add(found);
                }

                string known = categoryOfTechnique.Keys.First(k => k.Equals(name, StringComparison.OrdinalIgnoreCase) || k.Equals(Regex.Replace(name, "^regular ", ""), StringComparison.OrdinalIgnoreCase));
                if (!puzzle.Techniques.Contains(known, StringComparer.OrdinalIgnoreCase))
                {
                    puzzle.Techniques.Add(known);
                }
            }
        }

        // the index's order of the categories
        var order = categoryOfTechnique.Values.Distinct().ToList();
        puzzle.Categories.Sort((x, y) => order.IndexOf(x).CompareTo(order.IndexOf(y)));
    }
    if (imagesFolder != null && puzzle.Image != null)
    {
        Directory.CreateDirectory(imagesFolder);
        string target = Path.Combine(imagesFolder, puzzle.Slug + Path.GetExtension(puzzle.Image));
        if (!File.Exists(target))
        {
            var bytes = await http.GetByteArrayAsync(Site + "/notes/files/" + Uri.EscapeDataString(puzzle.Image));
            await File.WriteAllBytesAsync(target, bytes);
            await Task.Delay(300);
        }
    }
}

// ---------- the catalog ----------

// what was written by hand into the catalog (a puzzle's "play" line) is kept: fields this tool
// doesn't write are carried over from the file as it was
var kept = new Dictionary<string, JsonObject>();
if (File.Exists(catalogFile) && JsonNode.Parse(await File.ReadAllTextAsync(catalogFile))?["puzzles"] is JsonArray existing)
{
    foreach (var node in existing.OfType<JsonObject>())
    {
        kept[(string)node["slug"]] = node;
    }
}

var written = new HashSet<string> { "slug", "name", "title", "date", "statement", "thread", "categories", "techniques", "drawing" };
var array = new JsonArray();
foreach (var puzzle in puzzles)
{
    var entry = new JsonObject
    {
        ["slug"] = puzzle.Slug,
        ["name"] = puzzle.Name,
        ["title"] = puzzle.Title ?? TitleCase(puzzle.Name),
        ["date"] = puzzle.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
    };
    if (puzzle.Statement != null)
    {
        entry["statement"] = puzzle.Statement;
    }

    if (puzzle.Thread != null)
    {
        entry["thread"] = puzzle.Thread;
    }

    entry["categories"] = new JsonArray(puzzle.Categories.Select(c => (JsonNode)c).ToArray());
    entry["techniques"] = new JsonArray(puzzle.Techniques.Select(t => (JsonNode)t).ToArray());
    if (drawn.Contains(puzzle.Slug))
    {
        entry["drawing"] = puzzle.Slug + ".lgf";
    }

    if (kept.TryGetValue(puzzle.Slug, out var old))
    {
        // details fetched before and not this time stay as they were
        foreach (var field in new[] { "statement", "thread" })
        {
            if (entry[field] == null && old[field] != null)
            {
                entry[field] = old[field].DeepClone();
            }
        }

        foreach (var (key, value) in old)
        {
            if (!written.Contains(key))
            {
                entry[key] = value?.DeepClone();
            }
        }
    }

    array.Add((JsonNode)entry);
}

var catalog = new JsonObject
{
    ["source"] = Site + "/notes/published/HomePage",
    ["puzzles"] = array,
};
Directory.CreateDirectory(libraryFolder);
string json = catalog.ToJsonString(new JsonSerializerOptions { WriteIndented = true, NewLine = "\r\n", Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping });
await File.WriteAllTextAsync(catalogFile, json + "\r\n", new UTF8Encoding(false));
Console.WriteLine($"{catalogFile}: {puzzles.Count} puzzles, {drawn.Count} drawn, {detailed.Count} with details");
return 0;

// The source of a page of the wiki (its Markdown, in the textarea of /notes/source/<name>), cached
async Task<string> Source(string name)
{
    string cached = Path.Combine(cacheFolder, Regex.Replace(name, @"[^A-Za-z0-9]+", "_") + ".txt");
    if (!refresh && File.Exists(cached))
    {
        return await File.ReadAllTextAsync(cached);
    }

    var response = await http.GetAsync(Site + "/notes/source/" + PageName(name));
    if (!response.IsSuccessStatusCode)
    {
        return null;
    }

    string html = await response.Content.ReadAsStringAsync();
    var match = Regex.Match(html, "<textarea[^>]*>(.*?)</textarea>", RegexOptions.Singleline);
    if (!match.Success)
    {
        return null;
    }

    string text = WebUtility.HtmlDecode(match.Groups[1].Value);
    await File.WriteAllTextAsync(cached, text);
    // a small wiki on a small server: one page at a time, and a pause between them
    await Task.Delay(300);
    return text;
}

static string PageName(string name) => Uri.EscapeDataString(name).Replace("%20", "+").Replace("'", "%27");

static bool Same(string a, string b) => string.Equals(Normalize(a), Normalize(b), StringComparison.OrdinalIgnoreCase);

static string Normalize(string name) => Regex.Replace(WebUtility.HtmlDecode(name).Replace('’', '\''), @"\s+", " ").Trim();

static DateTime ParseDate(string text, int year)
{
    text = Regex.Replace(text.Trim(), @"(\d+)(st|nd|rd|th)", "$1");
    foreach (var format in new[] { "MMMM d, yyyy", "MMMM d yyyy", "d MMMM, yyyy", "d MMMM yyyy", "MMM d, yyyy" })
    {
        if (DateTime.TryParseExact(text, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
        {
            // a date on the year's page with another year is a typo (June 28, 2005 among 2025's)
            return date.Year == year ? date : new DateTime(year, date.Month, date.Day);
        }
    }

    return new DateTime(year, 1, 1);
}

static string TitleCase(string name)
{
    var small = new HashSet<string> { "a", "an", "and", "in", "on", "of", "the", "to", "with", "by", "at", "from", "inside", "across" };
    var words = Normalize(name).Split(' ');
    for (int i = 0; i < words.Length; i++)
    {
        var word = words[i];
        if (Regex.IsMatch(word, "^(i|ii|iii|iv|v|vi|vii|viii|ix|x|xi|xii|xiii|xiv|xv)$"))
        {
            words[i] = word.ToUpperInvariant();
        }
        else if (i == 0 || !small.Contains(word))
        {
            words[i] = string.Join("-", word.Split('-').Select(part => part.Length > 0 ? char.ToUpperInvariant(part[0]) + part.Substring(1) : part));
        }
    }

    return string.Join(" ", words);
}

string Option(string name)
{
    int index = Array.IndexOf(args, name);
    return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
}

static string FindRepository()
{
    for (var folder = new DirectoryInfo(Directory.GetCurrentDirectory()); folder != null; folder = folder.Parent)
    {
        if (File.Exists(Path.Combine(folder.FullName, "AGENTS.md")) && Directory.Exists(Path.Combine(folder.FullName, "Main")))
        {
            return folder.FullName;
        }
    }

    throw new Exception("run this from inside the repository");
}

class Puzzle
{
    public string Name;
    public string Title;
    public DateTime Date;
    public string Thumbnail;
    public string Image;
    public string Statement;
    public string Thread;
    public List<string> Categories = new();
    public List<string> Techniques = new();

    public string Slug => Regex.Replace(Normalize(Name).ToLowerInvariant().Replace("'", ""), "[^a-z0-9]+", "-").Trim('-');

    static string Normalize(string name) => Regex.Replace(WebUtility.HtmlDecode(name).Replace('’', '\''), @"\s+", " ").Trim();
}
