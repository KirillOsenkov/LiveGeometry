using System;
using Avalonia;
using Avalonia.Media;
using Avalonia.Media.Fonts;

namespace LiveGeometry;

/// <summary>
/// The font of the math symbols Inter (the font the app embeds) doesn't have - ∠ ⊥ ∥ ≅ ∼ ∈
/// ∩ ∪ ⌒ and the rest of the Mathematical Operators: a subset of Noto Sans Math
/// (Main/Avalonia/Fonts, made by tools/mathfont.cs), the first font any text falls back on.
/// The browser has no system fonts and drew a box for each of them; the desktop found them
/// in the system's fonts, and draws them as the browser does now.
/// </summary>
public static class MathFont
{
    /// <summary>The family, in the collection the app registers it in</summary>
    public const string FamilyName = "fonts:Math#Noto Sans Math";

    public static AppBuilder WithMathFont(this AppBuilder builder)
    {
        return builder
            .ConfigureFonts(fontManager => fontManager.AddFontCollection(
                new EmbeddedFontCollection(new Uri("fonts:Math"), new Uri("avares://LiveGeometry.App/Assets/Fonts"))))
            .With(new FontManagerOptions()
            {
                FontFallbacks = [new FontFallback() { FontFamily = new FontFamily(FamilyName) }]
            });
    }
}
