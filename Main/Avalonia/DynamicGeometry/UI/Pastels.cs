using Avalonia.Media;

namespace DynamicGeometry;

/// <summary>
/// Hand-picked near-white tints for plates that carry content (the gallery tiles): pale enough
/// for anything drawn on them, different enough from each other that neighbors read as
/// different. The order is for cycling: picked from a few shuffles as the one where no two
/// neighbors share a hue and the grays don't line up in a column of six-wide rows.
/// </summary>
public static class Pastels
{
    public static readonly Color[] Colors =
    {
        Color.Parse("#FDEEE3"), // peach
        Color.Parse("#E4F7EA"), // mint
        Color.Parse("#EAE8E4"), // stone
        Color.Parse("#EFE6DA"), // mocha
        Color.Parse("#FBEFD9"), // honey
        Color.Parse("#F1E8FA"), // lilac
        Color.Parse("#FCE9E9"), // rose
        Color.Parse("#F5F5F5"), // light gray
        Color.Parse("#EDF2E8"), // sage
        Color.Parse("#F9FADC"), // lemon
        Color.Parse("#EAE9FA"), // periwinkle
        Color.Parse("#F0E6D2"), // sand
        Color.Parse("#EFF8DE"), // lime
        Color.Parse("#E9ECEF"), // slate
        Color.Parse("#E8F5F0"), // eucalyptus
        Color.Parse("#E2F2FB"), // sky
        Color.Parse("#EEEEF4"), // fog
        Color.Parse("#FDF2E0"), // apricot
        Color.Parse("#FDF7D8"), // butter
        Color.Parse("#F5ECD5"), // wheat
        Color.Parse("#F8E6F6"), // orchid
        Color.Parse("#EDEEF0"), // silver
        Color.Parse("#E4EEFA"), // cornflower
        Color.Parse("#E1F6F3"), // seafoam
        Color.Parse("#FBE7EF"), // blush
        Color.Parse("#F3EADC"), // tan
        Color.Parse("#E8EEF9"), // powder blue
        Color.Parse("#F1EFEC"), // warm gray
    };

    /// <summary>The color at an index, wrapping around</summary>
    public static Color At(int index)
    {
        return Colors[((index % Colors.Length) + Colors.Length) % Colors.Length];
    }
}
