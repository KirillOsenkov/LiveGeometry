using System;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using Avalonia.Media;
using SkiaSharp;

namespace LiveGeometry;

/// <summary>
/// Skia writes text into an SVG as text in a font, which looks as it does here only where
/// that font is installed - and the fonts the app brings along (Inter, the emoji) are
/// installed nowhere: a viewer draws its own emoji instead. So every text element is
/// replaced by the outlines of its characters, and a color emoji by its layers, each in its
/// color. The picture is then the same everywhere; the price is that its text is shapes.
/// A text that can't be outlined (its font isn't found, a character isn't in it) stays text.
/// </summary>
public static class SvgTextOutlines
{
    static readonly XNamespace Svg = "http://www.w3.org/2000/svg";

    /// <summary>What places and styles the characters; the rest (fill, transform...) goes on with the outlines</summary>
    static readonly HashSet<string> textAttributes = new HashSet<string>()
    {
        "x", "y", "font-size", "font-family", "font-weight", "font-style", "font-stretch"
    };

    public static void Convert(XDocument document)
    {
        using var fonts = new FontFaces();
        foreach (var text in document.Descendants(Svg + "text").ToArray())
        {
            // an empty line of a label
            if (string.IsNullOrWhiteSpace(text.Value))
            {
                text.Remove();
                continue;
            }

            var outlines = CreateOutlines(text, fonts);
            if (outlines != null)
            {
                text.ReplaceWith(outlines);
            }
        }
    }

    /// <summary>
    /// Skia's text element is one run of characters in one font: the characters as its
    /// content, where each one starts in x, the baseline (or one per character) in y
    /// </summary>
    static XElement CreateOutlines(XElement text, FontFaces fonts)
    {
        var face = fonts.Find(
            (string)text.Attribute("font-family"),
            ReadWeight((string)text.Attribute("font-weight")),
            isItalic: (string)text.Attribute("font-style") is "italic" or "oblique");
        var size = ReadNumbers((string)text.Attribute("font-size"));
        var x = ReadNumbers((string)text.Attribute("x"));
        var y = ReadNumbers((string)text.Attribute("y"));

        // around the characters is the indentation of the file
        var characters = ReadCodePoints(text.Value.Trim('\r', '\n', '\t'));
        if (face == null || size.Count != 1 || y.Count == 0 || x.Count != characters.Count)
        {
            return null;
        }

        using var font = new SKFont(face.Typeface, size[0])
        {
            Hinting = SKFontHinting.None,
            LinearMetrics = true,
            Subpixel = true
        };
        var group = new XElement(Svg + "g", text.Attributes().Where(a => !textAttributes.Contains(a.Name.LocalName)));

        // characters of one color are one path, until a color emoji comes between them
        using var plain = new SKPath();
        for (int i = 0; i < characters.Count; i++)
        {
            ushort glyph = font.GetGlyph(characters[i]);
            if (glyph == 0)
            {
                return null;
            }

            float left = x[i];
            float baseline = y[System.Math.Min(i, y.Count - 1)];
            var layers = face.Colors?.GetLayers(glyph);
            if (layers == null)
            {
                using var outline = font.GetGlyphPath(glyph);
                if (outline != null)
                {
                    plain.AddPath(outline, left, baseline);
                }

                continue;
            }

            AddPath(group, plain, color: null);
            plain.Reset();
            foreach (var layer in layers)
            {
                using var outline = font.GetGlyphPath(layer.Glyph);
                if (outline != null)
                {
                    outline.Transform(SKMatrix.CreateTranslation(left, baseline));
                    AddPath(group, outline, layer.Color);
                }
            }
        }

        AddPath(group, plain, color: null);
        return group;
    }

    /// <param name="color">Null: in the color of the text, which the group has</param>
    static void AddPath(XElement group, SKPath outline, SKColor? color)
    {
        if (outline.IsEmpty)
        {
            return;
        }

        var path = new XElement(Svg + "path");
        if (color != null)
        {
            var value = color.Value;
            path.SetAttributeValue("fill", $"#{value.Red:X2}{value.Green:X2}{value.Blue:X2}");
            if (value.Alpha < 255)
            {
                path.SetAttributeValue("fill-opacity", (value.Alpha / 255.0).ToString("0.###", CultureInfo.InvariantCulture));
            }
        }

        path.SetAttributeValue("d", outline.ToSvgPathData());
        group.Add(path);
    }

    /// <summary>
    /// The weight of the font as Skia writes it, which from 500 on is one step too light
    /// (SkSVGDevice looks the name up in "100", "200", "300", "normal", "400", "500", "600",
    /// "bold", "800", "900" by the hundreds of the weight): "600" is a bold 700, "bold" is 800
    /// </summary>
    static int ReadWeight(string text)
    {
        switch (text)
        {
            case null:
            case "normal":
                return 400;
            case "bold":
                return 800;
        }

        if (!int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int weight))
        {
            return 400;
        }

        return weight >= 400 ? weight + 100 : weight;
    }

    /// <summary>"0, 8.5, 17.9, ": an empty list for anything else</summary>
    static List<float> ReadNumbers(string text)
    {
        var numbers = new List<float>();
        if (text == null)
        {
            return numbers;
        }

        foreach (var part in text.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries))
        {
            if (!float.TryParse(part, NumberStyles.Float, CultureInfo.InvariantCulture, out float number))
            {
                return new List<float>();
            }

            numbers.Add(number);
        }

        return numbers;
    }

    static List<int> ReadCodePoints(string text)
    {
        var codePoints = new List<int>();
        for (int i = 0; i < text.Length; i++)
        {
            if (char.IsSurrogatePair(text, i))
            {
                codePoints.Add(char.ConvertToUtf32(text, i));
                i++;
            }
            else
            {
                codePoints.Add(text[i]);
            }
        }

        return codePoints;
    }

    /// <summary>The fonts of one document, each opened once</summary>
    class FontFaces : IDisposable
    {
        // where Avalonia keeps the Inter it embeds (WithInterFont)
        const string InterCollection = "fonts:Inter#";

        readonly Dictionary<(string, int, bool), FontFace> faces = new Dictionary<(string, int, bool), FontFace>();

        /// <param name="families">As Skia wrote them: the names of one font, separated by commas</param>
        public FontFace Find(string families, int weight, bool isItalic)
        {
            if (string.IsNullOrEmpty(families))
            {
                return null;
            }

            var key = (families, weight, isItalic);
            if (faces.TryGetValue(key, out var face))
            {
                return face;
            }

            foreach (var family in families.Split(','))
            {
                var typeface = OpenEmbedded(family.Trim(), weight, isItalic) ?? OpenInstalled(family.Trim(), weight, isItalic);
                if (typeface != null)
                {
                    face = new FontFace(typeface);
                    break;
                }
            }

            faces[key] = face;
            return face;
        }

        /// <summary>The fonts the app brings along: they come first, being what the screen is drawn with</summary>
        static SKTypeface OpenEmbedded(string family, int weight, bool isItalic)
        {
            GlyphTypeface found = DynamicGeometry.EmojiFont.Typeface;
            if (found == null || found.PlatformTypeface.FamilyName != family)
            {
                var typeface = new Typeface(
                    InterCollection + family,
                    isItalic ? FontStyle.Italic : FontStyle.Normal,
                    (FontWeight)weight);
                if (!FontManager.Current.TryGetGlyphTypeface(typeface, out found) || found.PlatformTypeface.FamilyName != family)
                {
                    return null;
                }
            }

            if (!found.PlatformTypeface.TryGetStream(out var stream))
            {
                return null;
            }

            return SKTypeface.FromStream(stream);
        }

        static SKTypeface OpenInstalled(string family, int weight, bool isItalic)
        {
            var style = new SKFontStyle(
                weight,
                (int)SKFontStyleWidth.Normal,
                isItalic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright);
            var typeface = SKFontManager.Default.MatchFamily(family, style);

            // a family that isn't there is answered with another one
            if (typeface != null && typeface.FamilyName != family)
            {
                typeface.Dispose();
                return null;
            }

            return typeface;
        }

        public void Dispose()
        {
            foreach (var face in faces.Values)
            {
                face?.Typeface.Dispose();
            }

            faces.Clear();
        }
    }

    class FontFace
    {
        public FontFace(SKTypeface typeface)
        {
            Typeface = typeface;
            Colors = ColorGlyphs.Read(typeface);
        }

        public SKTypeface Typeface { get; }

        /// <summary>Null for a font of one color</summary>
        public ColorGlyphs Colors { get; }
    }

    /// <summary>One layer of a color glyph: the outline of a glyph of its own, in one color</summary>
    /// <param name="Color">Null for the color of the text</param>
    record struct ColorLayer(ushort Glyph, SKColor? Color);

    /// <summary>
    /// The COLR and CPAL tables of a color font (version 0, which is what Twemoji and Segoe UI
    /// Emoji have): a color glyph is a stack of layers, bottom first, and a layer is a glyph
    /// filled with a color of the palette
    /// </summary>
    class ColorGlyphs
    {
        static readonly uint ColrTag = Tag("COLR");
        static readonly uint CpalTag = Tag("CPAL");

        const int BaseGlyphRecordSize = 6;
        const int LayerRecordSize = 4;
        const int ColorRecordSize = 4;
        const ushort TextColorIndex = 0xFFFF;

        readonly Dictionary<ushort, ColorLayer[]> glyphs = new Dictionary<ushort, ColorLayer[]>();

        /// <summary>Null when the font has no color glyphs of this kind</summary>
        public static ColorGlyphs Read(SKTypeface typeface)
        {
            if (!typeface.TryGetTableData(ColrTag, out var colr)
                || !typeface.TryGetTableData(CpalTag, out var cpal)
                || colr.Length < 14
                || cpal.Length < 14)
            {
                return null;
            }

            int glyphCount = ReadUInt16(colr, 2);
            int glyphsOffset = (int)ReadUInt32(colr, 4);
            int layersOffset = (int)ReadUInt32(colr, 8);
            int layerCount = ReadUInt16(colr, 12);

            // the first palette: where it starts among the colors
            int colorCount = ReadUInt16(cpal, 6);
            int colorsOffset = (int)ReadUInt32(cpal, 8);
            int firstColor = ReadUInt16(cpal, 12);
            if (glyphCount == 0
                || glyphsOffset + glyphCount * BaseGlyphRecordSize > colr.Length
                || layersOffset + layerCount * LayerRecordSize > colr.Length
                || colorsOffset + colorCount * ColorRecordSize > cpal.Length)
            {
                return null;
            }

            var result = new ColorGlyphs();
            for (int i = 0; i < glyphCount; i++)
            {
                int record = glyphsOffset + i * BaseGlyphRecordSize;
                int firstLayer = ReadUInt16(colr, record + 2);
                int count = ReadUInt16(colr, record + 4);
                if (firstLayer + count > layerCount)
                {
                    continue;
                }

                var layers = new ColorLayer[count];
                for (int j = 0; j < count; j++)
                {
                    int layer = layersOffset + (firstLayer + j) * LayerRecordSize;
                    int colorIndex = ReadUInt16(colr, layer + 2);
                    SKColor? color = null;
                    if (colorIndex != TextColorIndex && firstColor + colorIndex < colorCount)
                    {
                        // blue, green, red, alpha
                        int at = colorsOffset + (firstColor + colorIndex) * ColorRecordSize;
                        color = new SKColor(cpal[at + 2], cpal[at + 1], cpal[at], cpal[at + 3]);
                    }

                    layers[j] = new ColorLayer(ReadUInt16(colr, layer), color);
                }

                result.glyphs[ReadUInt16(colr, record)] = layers;
            }

            return result;
        }

        /// <summary>Null for a glyph of one color</summary>
        public ColorLayer[] GetLayers(ushort glyph)
        {
            return glyphs.TryGetValue(glyph, out var layers) ? layers : null;
        }

        static uint Tag(string name)
        {
            return (uint)(name[0] << 24 | name[1] << 16 | name[2] << 8 | name[3]);
        }

        static ushort ReadUInt16(byte[] table, int offset)
        {
            return BinaryPrimitives.ReadUInt16BigEndian(table.AsSpan(offset));
        }

        static uint ReadUInt32(byte[] table, int offset)
        {
            return BinaryPrimitives.ReadUInt32BigEndian(table.AsSpan(offset));
        }
    }
}
