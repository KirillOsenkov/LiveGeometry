#:property Nullable=disable
#:property PublishAot=false

// mathfont - makes the font of math symbols the app falls back on (Main/Avalonia/Fonts/
// NotoSansMath.ttf): a subset of Noto Sans Math (SIL Open Font License 1.1, no reserved
// name; https://github.com/notofonts/math) with the symbols Inter, the font the app
// embeds, doesn't have - ∠ ⊥ ∥ ≅ ∼ ∈ ∩ ∪ ⌒ and the rest of the Mathematical Operators.
// The browser has no system fonts, and showed a box for each of them.
//
//   dotnet run tools/mathfont.cs -- <NotoSansMath-Regular.ttf> <out.ttf> [inter.ttf]
//
// The source is the unhinted TrueType file of notofonts.github.io
// (fonts/NotoSansMath/unhinted/ttf/NotoSansMath-Regular.ttf). With an Inter file, the
// characters Inter has are left out. The subset keeps the outlines, the advance widths,
// the cmap and the names; layout tables (GSUB, GPOS, MATH) are dropped, which a label's
// text doesn't use.

using System.Buffers.Binary;
using System.Text;

if (args.Length < 2)
{
    Console.WriteLine("usage: mathfont <NotoSansMath-Regular.ttf> <out.ttf> [inter.ttf]");
    return 1;
}

var source = new Font(File.ReadAllBytes(args[0]));
var inter = args.Length > 2 ? new Font(File.ReadAllBytes(args[2])) : null;

var wanted = new List<int>();
for (int c = 0x2200; c <= 0x22FF; c++)
{
    wanted.Add(c);
}

// ⌀ ⌈ ⌉ ⌊ ⌋ ⌒ ⌓, ⟂ ⟨ ⟩ ⟦ ⟧, △ ▱ ◯, ℕ ℤ ℚ ℝ ℂ ℙ, ′ ″ ‴
wanted.AddRange([0x2300, 0x2308, 0x2309, 0x230A, 0x230B, 0x2312, 0x2313, 0x27C2, 0x27E8, 0x27E9, 0x27E6, 0x27E7,
    0x25B3, 0x25B1, 0x25EF, 0x2115, 0x2124, 0x211A, 0x211D, 0x2102, 0x2119, 0x2032, 0x2033, 0x2034]);

var map = source.ReadCmap();
var interMap = inter?.ReadCmap();
var kept = wanted
    .Where(c => map.ContainsKey(c) && (interMap == null || !interMap.ContainsKey(c)))
    .ToList();
Console.WriteLine($"{kept.Count} characters of {wanted.Count} ({wanted.Count(c => !map.ContainsKey(c))} not in the source, " +
    $"{wanted.Count(c => interMap != null && interMap.ContainsKey(c))} in Inter)");

var output = Subset(source, kept.ToDictionary(c => c, c => map[c]));
File.WriteAllBytes(args[1], output);
Console.WriteLine($"wrote {args[1]}: {output.Length} bytes");
return 0;

static byte[] Subset(Font font, Dictionary<int, int> characters)
{
    var head = font.Table("head");
    var maxp = font.Table("maxp");
    var hhea = font.Table("hhea");
    var hmtx = font.Table("hmtx");
    var loca = font.Table("loca");
    var glyf = font.Table("glyf");
    int glyphCount = U16(maxp, 4);
    bool longOffsets = U16(head, 50) == 1;
    int metricCount = U16(hhea, 34);

    int GlyphStart(int glyph) => longOffsets ? (int)U32(loca, glyph * 4) : U16(loca, glyph * 2) * 2;
    byte[] GlyphData(int glyph) => glyf.AsSpan(GlyphStart(glyph), GlyphStart(glyph + 1) - GlyphStart(glyph)).ToArray();

    // the glyphs: .notdef, the characters', and the parts of composite glyphs
    var glyphs = new SortedSet<int> { 0 };
    var pending = new Stack<int>(characters.Values);
    while (pending.Count > 0)
    {
        int glyph = pending.Pop();
        if (!glyphs.Add(glyph) && glyph != 0)
        {
            continue;
        }

        foreach (var (_, component) in Components(GlyphData(glyph)))
        {
            if (!glyphs.Contains(component))
            {
                pending.Push(component);
            }
        }
    }

    var order = glyphs.ToList();
    var newIndex = order.Select((old, index) => (old, index)).ToDictionary(pair => pair.old, pair => pair.index);

    // glyf and loca (long offsets), composites pointing at the new indexes
    var newGlyf = new List<byte>();
    var newLoca = new List<byte>();
    foreach (int old in order)
    {
        AppendU32(newLoca, (uint)newGlyf.Count);
        var data = GlyphData(old);
        foreach (var (offset, component) in Components(data))
        {
            BinaryPrimitives.WriteUInt16BigEndian(data.AsSpan(offset), (ushort)newIndex[component]);
        }

        newGlyf.AddRange(data);
        while (newGlyf.Count % 4 != 0)
        {
            newGlyf.Add(0);
        }
    }

    AppendU32(newLoca, (uint)newGlyf.Count);

    // hmtx: a full metric for every glyph
    var newHmtx = new List<byte>();
    foreach (int old in order)
    {
        int advance = U16(hmtx, System.Math.Min(old, metricCount - 1) * 4);
        int leftSideBearing = old < metricCount
            ? U16(hmtx, old * 4 + 2)
            : U16(hmtx, metricCount * 4 + (old - metricCount) * 2);
        AppendU16(newHmtx, advance);
        AppendU16(newHmtx, leftSideBearing);
    }

    var newHead = (byte[])head.Clone();
    BinaryPrimitives.WriteUInt16BigEndian(newHead.AsSpan(50), 1);
    BinaryPrimitives.WriteUInt32BigEndian(newHead.AsSpan(8), 0);
    var newMaxp = (byte[])maxp.Clone();
    BinaryPrimitives.WriteUInt16BigEndian(newMaxp.AsSpan(4), (ushort)order.Count);
    var newHhea = (byte[])hhea.Clone();
    BinaryPrimitives.WriteUInt16BigEndian(newHhea.AsSpan(34), (ushort)order.Count);

    // post format 3: no glyph names
    var newPost = font.Table("post").AsSpan(0, 32).ToArray();
    BinaryPrimitives.WriteUInt32BigEndian(newPost, 0x00030000);

    // OS/2: the first and the last character
    var newOs2 = (byte[])font.Table("OS/2").Clone();
    BinaryPrimitives.WriteUInt16BigEndian(newOs2.AsSpan(64), (ushort)characters.Keys.Min());
    BinaryPrimitives.WriteUInt16BigEndian(newOs2.AsSpan(66), (ushort)characters.Keys.Max());

    var tables = new SortedDictionary<string, byte[]>(StringComparer.Ordinal)
    {
        ["OS/2"] = newOs2,
        ["cmap"] = Cmap(characters.ToDictionary(pair => pair.Key, pair => newIndex[pair.Value])),
        ["glyf"] = newGlyf.ToArray(),
        ["head"] = newHead,
        ["hhea"] = newHhea,
        ["hmtx"] = newHmtx.ToArray(),
        ["loca"] = newLoca.ToArray(),
        ["maxp"] = newMaxp,
        ["name"] = font.Table("name"),
        ["post"] = newPost
    };

    var result = WriteFont(tables);

    // head's checkSumAdjustment: 0xB1B0AFBA less the sum of the whole file
    int headOffset = TableOffset(result, "head");
    BinaryPrimitives.WriteUInt32BigEndian(result.AsSpan(headOffset + 8), 0xB1B0AFBA - Checksum(result));
    return result;
}

// where a composite glyph names its parts: the offset of each glyph index in the data
static IEnumerable<(int Offset, int Glyph)> Components(byte[] data)
{
    if (data.Length < 10 || (short)U16(data, 0) >= 0)
    {
        yield break;
    }

    int offset = 10;
    while (true)
    {
        int flags = U16(data, offset);
        yield return (offset + 2, U16(data, offset + 2));
        offset += 4 + ((flags & 0x0001) != 0 ? 4 : 2);
        if ((flags & 0x0008) != 0)
        {
            offset += 2;
        }
        else if ((flags & 0x0040) != 0)
        {
            offset += 4;
        }
        else if ((flags & 0x0080) != 0)
        {
            offset += 8;
        }

        if ((flags & 0x0020) == 0)
        {
            yield break;
        }
    }
}

// a cmap of one format 4 subtable, for Unicode (0, 3) and Windows (3, 1): a segment per run
// of characters whose glyphs run on too
static byte[] Cmap(Dictionary<int, int> characters)
{
    var segments = new List<(int Start, int End, int Delta)>();
    foreach (var c in characters.Keys.Order())
    {
        int delta = characters[c] - c;
        if (segments.Count > 0 && segments[^1].End == c - 1 && segments[^1].Delta == delta)
        {
            segments[^1] = (segments[^1].Start, c, delta);
        }
        else
        {
            segments.Add((c, c, delta));
        }
    }

    segments.Add((0xFFFF, 0xFFFF, 1));
    int segmentCount = segments.Count;
    int searchRange = 2 * (1 << (int)System.Math.Floor(System.Math.Log2(segmentCount)));
    var table = new List<byte>();
    AppendU16(table, 4);
    AppendU16(table, 16 + segmentCount * 8);
    AppendU16(table, 0);
    AppendU16(table, segmentCount * 2);
    AppendU16(table, searchRange);
    AppendU16(table, (int)System.Math.Log2(searchRange / 2));
    AppendU16(table, segmentCount * 2 - searchRange);
    segments.ForEach(segment => AppendU16(table, segment.End));
    AppendU16(table, 0);
    segments.ForEach(segment => AppendU16(table, segment.Start));
    segments.ForEach(segment => AppendU16(table, segment.Delta & 0xFFFF));
    segments.ForEach(segment => AppendU16(table, 0));

    var cmap = new List<byte>();
    AppendU16(cmap, 0);
    AppendU16(cmap, 2);
    foreach (var (platform, encoding) in new[] { (0, 3), (3, 1) })
    {
        AppendU16(cmap, platform);
        AppendU16(cmap, encoding);
        AppendU32(cmap, 4 + 2 * 8);
    }

    cmap.AddRange(table);
    return cmap.ToArray();
}

static byte[] WriteFont(SortedDictionary<string, byte[]> tables)
{
    int count = tables.Count;
    int searchRange = 16 * (1 << (int)System.Math.Floor(System.Math.Log2(count)));
    var header = new List<byte>();
    AppendU32(header, 0x00010000);
    AppendU16(header, count);
    AppendU16(header, searchRange);
    AppendU16(header, (int)System.Math.Log2(searchRange / 16));
    AppendU16(header, count * 16 - searchRange);

    var body = new List<byte>();
    int offset = 12 + count * 16;
    foreach (var (tag, data) in tables)
    {
        header.AddRange(Encoding.ASCII.GetBytes(tag));
        AppendU32(header, Checksum(data));
        AppendU32(header, (uint)(offset + body.Count));
        AppendU32(header, (uint)data.Length);
        body.AddRange(data);
        while (body.Count % 4 != 0)
        {
            body.Add(0);
        }
    }

    return header.Concat(body).ToArray();
}

static int TableOffset(byte[] font, string tag)
{
    int count = U16(font, 4);
    for (int i = 0; i < count; i++)
    {
        int record = 12 + i * 16;
        if (Encoding.ASCII.GetString(font, record, 4) == tag)
        {
            return (int)U32(font, record + 8);
        }
    }

    throw new InvalidDataException("no table " + tag);
}

static uint Checksum(byte[] data)
{
    uint sum = 0;
    for (int i = 0; i < data.Length; i += 4)
    {
        uint word = 0;
        for (int j = 0; j < 4; j++)
        {
            word = (word << 8) | (i + j < data.Length ? data[i + j] : 0u);
        }

        sum += word;
    }

    return sum;
}

static int U16(byte[] data, int offset) => BinaryPrimitives.ReadUInt16BigEndian(data.AsSpan(offset));

static uint U32(byte[] data, int offset) => BinaryPrimitives.ReadUInt32BigEndian(data.AsSpan(offset));

static void AppendU16(List<byte> list, int value)
{
    list.Add((byte)(value >> 8));
    list.Add((byte)value);
}

static void AppendU32(List<byte> list, uint value)
{
    AppendU16(list, (int)(value >> 16));
    AppendU16(list, (int)(value & 0xFFFF));
}

/// <summary>A TrueType file: its tables by tag, and the characters of its cmap</summary>
class Font
{
    readonly byte[] data;
    readonly Dictionary<string, (int Offset, int Length)> tables = new();

    static int U16(byte[] data, int offset) => BinaryPrimitives.ReadUInt16BigEndian(data.AsSpan(offset));

    static uint U32(byte[] data, int offset) => BinaryPrimitives.ReadUInt32BigEndian(data.AsSpan(offset));

    public Font(byte[] data)
    {
        this.data = data;
        int count = U16(data, 4);
        for (int i = 0; i < count; i++)
        {
            int record = 12 + i * 16;
            tables[Encoding.ASCII.GetString(data, record, 4)] = ((int)U32(data, record + 8), (int)U32(data, record + 12));
        }
    }

    public byte[] Table(string tag)
    {
        var (offset, length) = tables[tag];
        return data.AsSpan(offset, length).ToArray();
    }

    /// <summary>Character to glyph, from a Unicode subtable of format 4 or 12</summary>
    public Dictionary<int, int> ReadCmap()
    {
        var cmap = Table("cmap");
        var result = new Dictionary<int, int>();
        int count = U16(cmap, 2);
        for (int i = 0; i < count; i++)
        {
            int platform = U16(cmap, 4 + i * 8);
            int encoding = U16(cmap, 6 + i * 8);
            int offset = (int)U32(cmap, 8 + i * 8);
            if (!(platform == 0 || platform == 3 && (encoding == 1 || encoding == 10)))
            {
                continue;
            }

            int format = U16(cmap, offset);
            if (format == 12)
            {
                int groups = (int)U32(cmap, offset + 12);
                for (int g = 0; g < groups; g++)
                {
                    int group = offset + 16 + g * 12;
                    uint start = U32(cmap, group);
                    uint end = U32(cmap, group + 4);
                    uint glyph = U32(cmap, group + 8);
                    for (uint c = start; c <= end; c++)
                    {
                        result[(int)c] = (int)(glyph + c - start);
                    }
                }
            }
            else if (format == 4)
            {
                int segments = U16(cmap, offset + 6) / 2;
                int ends = offset + 14;
                int starts = ends + segments * 2 + 2;
                int deltas = starts + segments * 2;
                int rangeOffsets = deltas + segments * 2;
                for (int s = 0; s < segments; s++)
                {
                    int end = U16(cmap, ends + s * 2);
                    int start = U16(cmap, starts + s * 2);
                    int delta = U16(cmap, deltas + s * 2);
                    int rangeOffset = U16(cmap, rangeOffsets + s * 2);
                    for (int c = start; c <= end && c != 0xFFFF; c++)
                    {
                        int glyph = rangeOffset == 0
                            ? (c + delta) & 0xFFFF
                            : U16(cmap, rangeOffsets + s * 2 + rangeOffset + (c - start) * 2);
                        if (glyph != 0 && !result.ContainsKey(c))
                        {
                            result[c] = rangeOffset == 0 || glyph == 0 ? glyph : (glyph + delta) & 0xFFFF;
                        }
                    }
                }
            }
        }

        return result;
    }
}
