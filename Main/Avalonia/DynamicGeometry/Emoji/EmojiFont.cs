using System;
using System.IO;
using System.Threading.Tasks;
using Avalonia.Media;
using Avalonia.Media.Fonts;

namespace DynamicGeometry;

/// <summary>
/// The font emoji are drawn with (Twemoji, a color font): the same on every platform, and the
/// browser has no system fonts to fall back on. It is 1.5 MB, so it is loaded on first use,
/// not at startup: a point with a character, or the Emoji tab of a style. Until then
/// <see cref="Family"/> falls back to whatever the platform finds. A character that isn't an
/// emoji comes from Inter, which the app embeds.
/// </summary>
public static class EmojiFont
{
    const string FamilyName = "Twemoji Mozilla";

    static readonly Uri CollectionKey = new Uri("fonts:Emoji");

    /// <summary>
    /// For a character that isn't an emoji (★, Ω): the font the app embeds on every platform.
    /// Without it named, the browser draws a missing glyph, having no system fonts to find one in.
    /// </summary>
    const string TextFamily = "fonts:Inter#Inter";

    static readonly FontFamily loadedFamily = new FontFamily(CollectionKey + "#" + FamilyName + ", " + TextFamily);

    /// <summary>
    /// The emoji font once it is loaded, the default until then: a family in a collection
    /// that isn't there yet makes text layout throw (while drawing, which ends the app)
    /// </summary>
    public static FontFamily Family => IsLoaded ? loadedFamily : FontFamily.Default;

    static GlyphTypeface emojiTypeface;

    /// <summary>
    /// The emoji font or the text font has every character of the text (an FE0F asks for the
    /// emoji look and needs none), so it looks the same everywhere. False until the font is
    /// loaded.
    /// </summary>
    public static bool CanDraw(string text)
    {
        if (emojiTypeface == null || string.IsNullOrEmpty(text))
        {
            return false;
        }

        FontManager.Current.TryGetGlyphTypeface(new Typeface(TextFamily), out var textTypeface);
        for (int i = 0; i < text.Length; i++)
        {
            int codePoint = text[i];
            if (char.IsSurrogatePair(text, i))
            {
                codePoint = char.ConvertToUtf32(text, i);
                i++;
            }

            if (codePoint != 0xFE0F
                && !emojiTypeface.CharacterToGlyphMap.ContainsGlyph(codePoint)
                && textTypeface?.CharacterToGlyphMap.ContainsGlyph(codePoint) != true)
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Opens the font file. The app sets it: desktop reads it from beside the executable, the
    /// browser fetches it from the site. Null: there is none, characters use the fallback.
    /// </summary>
    public static Func<Task<Stream>> Open { get; set; }

    public static bool IsLoaded { get; private set; }

    static Task loading;

    /// <summary>Starts loading on the first call; completes when the font is there (or failed to come)</summary>
    public static Task EnsureLoaded()
    {
        if (loading == null)
        {
            loading = Load();
        }

        return loading;
    }

    static async Task Load()
    {
        if (Open == null)
        {
            return;
        }

        try
        {
            // in memory, and kept: a download can't seek, and the typeface may read it later
            var data = new MemoryStream();
            using (var stream = await Open())
            {
                await stream.CopyToAsync(data);
            }

            data.Position = 0;
            var collection = new EmojiFontCollection(CollectionKey);
            if (collection.TryAddGlyphTypeface(data, out emojiTypeface))
            {
                FontManager.Current.AddFontCollection(collection);
                IsLoaded = true;
            }
        }
        catch (Exception ex)
        {
            // characters stay in the fallback font; the exception is reported as every one is
            Console.WriteLine("Emoji font: " + ex.Message);
        }
    }

    class EmojiFontCollection : FontCollectionBase
    {
        public EmojiFontCollection(Uri key)
        {
            Key = key;
        }

        public override Uri Key { get; }
    }
}
