// Port of Main/Avalonia/DynamicGeometry/Styles/StyleHue.cs: a hue of the default styles,
// a stroke for each paper, its fills worked out from the stroke.

class StyleHue {
    constructor(name, stroke, darkStroke, fillHue = null, darkFillHue = null) {
        this.name = name;
        this.stroke = Color.parse(stroke);
        this.darkStroke = Color.parse(darkStroke);
        this.fillHue = fillHue;
        this.darkFillHue = darkFillHue;
    }

    static Gray = new StyleHue("Gray", "#6E6E6E", "#9E9E9E");

    /** A golden brown outline, filled cream to gold: the column of the yellow point and the yellow fill */
    static Brown = new StyleHue("Brown", "#A87B05", "#E3B341", 52, 45);

    /** The column of the green point and the classic green fill */
    static Green = new StyleHue("Green", "#2E9E4F", "#4CD27A");

    static All = [
        StyleHue.Gray,
        new StyleHue("Red", "#D83B3B", "#FF6B6B"),
        new StyleHue("Orange", "#E08A00", "#FFA940"),
        StyleHue.Brown,
        StyleHue.Green,
        new StyleHue("Cyan", "#00A3BF", "#33D1EB"),
        new StyleHue("Blue", "#2F7BD6", "#6AA5FF"),
        new StyleHue("Purple", "#8854D0", "#B48CFF")
    ];

    /** The hues but gray, whose styles are the defaults of the theme */
    static get Colors() {
        return StyleHue.All.filter(hue => hue !== StyleHue.Gray);
    }

    static FillAlpha = 0xC8;

    /** Light tints of the strongest hues glare next to the others: a tint is at most this saturated */
    static MaxTintSaturation = 0.75;

    /** A shape's fill on the light paper: light at the top left, deeper at the bottom right */
    get fill() {
        return StyleHue.gradient(this.lightFillEnd, this.tint(0.78, StyleHue.FillAlpha));
    }

    /** A shape's fill on the dark paper: bright at the top left, deep at the bottom right */
    get darkFill() {
        return StyleHue.gradient(this.darkFillLightEnd, this.darkTint(0.17, StyleHue.FillAlpha, 1));
    }

    get solidFill() {
        return new SolidColorBrush(this.lightFillEnd);
    }

    get darkSolidFill() {
        return new SolidColorBrush(this.darkFillLightEnd);
    }

    get lightFillEnd() {
        return this.tint(0.97, StyleHue.FillAlpha);
    }

    get darkFillLightEnd() {
        return this.darkTint(0.6, StyleHue.FillAlpha, 1);
    }

    get beadFill() {
        return StyleHue.gradient(this.tint(0.9), this.tint(0.55));
    }

    get darkBeadFill() {
        return StyleHue.gradient(this.darkTint(0.88), this.darkTint(0.42));
    }

    get beadRim() {
        return this.tint(0.3);
    }

    get darkBeadRim() {
        return this.darkTint(0.25);
    }

    tint(lightness, alpha = 0xFF, maxSaturation = StyleHue.MaxTintSaturation) {
        return StyleHue.tintOf(this.stroke, this.fillHue, lightness, alpha, maxSaturation);
    }

    darkTint(lightness, alpha = 0xFF, maxSaturation = StyleHue.MaxTintSaturation) {
        return StyleHue.tintOf(this.darkStroke, this.darkFillHue, lightness, alpha, maxSaturation);
    }

    static tintOf(color, hue, lightness, alpha, maxSaturation) {
        const hsl = color.toHsl();
        const result = Color.fromHsl(hue ?? hsl.h, Math.min(hsl.s, maxSaturation), lightness);
        return Color.fromArgb(alpha, result.r, result.g, result.b);
    }

    /** A diagonal gradient across the box of whatever it fills, top left to bottom right */
    static gradient(from, to) {
        return new LinearGradientBrush(new Point(0, 0), new Point(1, 1), [new GradientStop(from, 0), new GradientStop(to, 1)]);
    }
}
