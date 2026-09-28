#:property TargetFramework=net10.0
#:property Nullable=disable
#:property PublishAot=false
#:package SkiaSharp@3.119.4

// emoji - writes the list the Emoji tab of a point style searches: every emoji of Unicode's
// emoji-test.txt that is a single character (with its FE0F, if any) and that the emoji font
// has a glyph for, in the file's order (CLDR order, which keeps kinds together), one per line:
// the emoji, a tab, its name, a tab, its subgroup ("food-fruit", which search also reads).
//
//   dotnet tools/emoji.cs -- <emoji-test.txt> <font.ttf> <Emoji.txt>
//
// emoji-test.txt: https://unicode.org/Public/emoji/latest/emoji-test.txt

using System.Globalization;
using System.Text;
using SkiaSharp;

if (args.Length < 3)
{
    Console.WriteLine("usage: emoji <emoji-test.txt> <font.ttf> <Emoji.txt>");
    return 1;
}

using var typeface = SKTypeface.FromFile(args[1]);
if (typeface == null)
{
    Console.WriteLine($"can't read the font {args[1]}");
    return 1;
}

var sb = new StringBuilder();
string subgroup = "";
int count = 0;
var missing = new List<string>();
foreach (var line in File.ReadLines(args[0]))
{
    if (line.StartsWith("# subgroup:"))
    {
        subgroup = line.Substring("# subgroup:".Length).Trim();
        continue;
    }

    if (line.Length == 0 || line.StartsWith('#'))
    {
        continue;
    }

    // 2764 FE0F  ; fully-qualified  # ❤️ E0.6 red heart
    int semicolon = line.IndexOf(';');
    int hash = line.IndexOf('#');
    if (semicolon < 0 || hash < semicolon)
    {
        continue;
    }

    string status = line.Substring(semicolon + 1, hash - semicolon - 1).Trim();
    if (status != "fully-qualified")
    {
        continue;
    }

    var codePoints = line.Substring(0, semicolon)
        .Split(' ', StringSplitOptions.RemoveEmptyEntries)
        .Select(hex => int.Parse(hex, NumberStyles.HexNumber))
        .ToArray();
    var characters = codePoints.Where(c => c != 0xFE0F).ToArray();
    if (characters.Length != 1)
    {
        continue;
    }

    // "❤️ E0.6 red heart": the emoji, its version, the name
    var comment = line.Substring(hash + 1).Trim().Split(' ', 3);
    string name = comment[2];
    if (!typeface.ContainsGlyph(characters[0]))
    {
        missing.Add(name);
        continue;
    }

    string text = string.Concat(codePoints.Select(char.ConvertFromUtf32));
    sb.Append(text).Append('\t').Append(name).Append('\t').Append(subgroup).Append('\n');
    count++;
}

File.WriteAllText(args[2], sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
Console.WriteLine($"{count} emoji written to {args[2]}");
if (missing.Count > 0)
{
    Console.WriteLine($"{missing.Count} not in the font: {string.Join(", ", missing)}");
}

return 0;
