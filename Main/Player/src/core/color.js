// Colors and brushes: Avalonia's Color, SolidColorBrush and LinearGradientBrush as the
// library uses them, and ColorText (Controls/ColorPicker/ColorText.cs) plus the ToColor
// extension of Utilities.cs for reading them from files.

class Color {
    constructor(a, r, g, b) {
        this.a = a;
        this.r = r;
        this.g = g;
        this.b = b;
    }

    static fromArgb(a, r, g, b) {
        return new Color(a, r, g, b);
    }

    static fromRgb(r, g, b) {
        return new Color(255, r, g, b);
    }

    static fromUInt32(argb) {
        return new Color((argb >>> 24) & 0xFF, (argb >>> 16) & 0xFF, (argb >>> 8) & 0xFF, argb & 0xFF);
    }

    /** Color.Parse: a name or a hex code */
    static parse(text) {
        const color = ColorText.tryParse(text);
        if (color == null) {
            throw new Error("Not a color: " + text);
        }

        return color;
    }

    static get black() {
        return new Color(255, 0, 0, 0);
    }

    static get white() {
        return new Color(255, 255, 255, 255);
    }

    static get transparent() {
        return new Color(0, 255, 255, 255);
    }

    withAlpha(alpha) {
        return new Color(alpha, this.r, this.g, this.b);
    }

    equals(other) {
        return other instanceof Color && this.a === other.a && this.r === other.r && this.g === other.g && this.b === other.b;
    }

    /** What a canvas takes */
    toCss() {
        return "rgba(" + this.r + "," + this.g + "," + this.b + "," + (this.a / 255) + ")";
    }

    toString() {
        return ColorText.toArgbHex(this);
    }

    /** Avalonia's Color.ToHsl: hue in degrees, saturation and lightness in 0..1 */
    toHsl() {
        const r = this.r / 255;
        const g = this.g / 255;
        const b = this.b / 255;
        const max = Math.max(r, g, b);
        const min = Math.min(r, g, b);
        const l = (max + min) / 2;
        let h = 0;
        let s = 0;
        if (max !== min) {
            const d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            if (max === r) {
                h = (g - b) / d + (g < b ? 6 : 0);
            } else if (max === g) {
                h = (b - r) / d + 2;
            } else {
                h = (r - g) / d + 4;
            }

            h *= 60;
        }

        return { h, s, l };
    }

    /** HslColor(...).ToRgb() */
    static fromHsl(hue, saturation, lightness, alpha = 255) {
        const h = ((hue % 360) + 360) % 360 / 360;
        const s = saturation;
        const l = lightness;
        let r;
        let g;
        let b;
        if (s === 0) {
            r = g = b = l;
        } else {
            const hueToRgb = (p, q, t) => {
                if (t < 0) {
                    t += 1;
                }

                if (t > 1) {
                    t -= 1;
                }

                if (t < 1 / 6) {
                    return p + (q - p) * 6 * t;
                }

                if (t < 1 / 2) {
                    return q;
                }

                if (t < 2 / 3) {
                    return p + (q - p) * (2 / 3 - t) * 6;
                }

                return p;
            };
            const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            const p = 2 * l - q;
            r = hueToRgb(p, q, h + 1 / 3);
            g = hueToRgb(p, q, h);
            b = hueToRgb(p, q, h - 1 / 3);
        }

        return new Color(alpha, Math.round(r * 255), Math.round(g * 255), Math.round(b * 255));
    }
}

const ColorText = {
    /** #RRGGBB, or #AARRGGBB when the color isn't opaque */
    toHex(color) {
        return color.a === 255
            ? "#" + hex2(color.r) + hex2(color.g) + hex2(color.b)
            : ColorText.toArgbHex(color);
    },

    /** #AARRGGBB, always: how colors go into drawing files */
    toArgbHex(color) {
        return "#" + hex2(color.a) + hex2(color.r) + hex2(color.g) + hex2(color.b);
    },

    /** A web color name, RGB, RRGGBB or AARRGGBB hex with or without the #; null when it is none */
    tryParse(text) {
        if (text == null || text.trim() === "") {
            return null;
        }

        text = text.trim();
        const named = NamedColors[text.toLowerCase()];
        if (named !== undefined) {
            return Color.fromUInt32(named);
        }

        let hex = text.replace(/^#/, "");
        if (hex.length === 3) {
            hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
        }

        if (hex.length === 6) {
            hex = "FF" + hex;
        }

        if (hex.length === 8 && /^[0-9a-fA-F]{8}$/.test(hex)) {
            return Color.fromUInt32(parseInt(hex, 16));
        }

        return null;
    },

    /** The ToColor extension: what a file says, black when it says nothing readable */
    toColor(text) {
        if (text == null || text === "") {
            return Color.black;
        }

        return ColorText.tryParse(text) ?? Color.black;
    }
};

function hex2(value) {
    return value.toString(16).toUpperCase().padStart(2, "0");
}

// The names Avalonia's Color.ToString() wrote into old files (and ColorPalette.ColorsByName
// reads): the web colors.
const NamedColors = {
    aliceblue: 0xFFF0F8FF, antiquewhite: 0xFFFAEBD7, aqua: 0xFF00FFFF, aquamarine: 0xFF7FFFD4, azure: 0xFFF0FFFF,
    beige: 0xFFF5F5DC, bisque: 0xFFFFE4C4, black: 0xFF000000, blanchedalmond: 0xFFFFEBCD, blue: 0xFF0000FF,
    blueviolet: 0xFF8A2BE2, brown: 0xFFA52A2A, burlywood: 0xFFDEB887, cadetblue: 0xFF5F9EA0, chartreuse: 0xFF7FFF00,
    chocolate: 0xFFD2691E, coral: 0xFFFF7F50, cornflowerblue: 0xFF6495ED, cornsilk: 0xFFFFF8DC, crimson: 0xFFDC143C,
    cyan: 0xFF00FFFF, darkblue: 0xFF00008B, darkcyan: 0xFF008B8B, darkgoldenrod: 0xFFB8860B, darkgray: 0xFFA9A9A9,
    darkgreen: 0xFF006400, darkkhaki: 0xFFBDB76B, darkmagenta: 0xFF8B008B, darkolivegreen: 0xFF556B2F, darkorange: 0xFFFF8C00,
    darkorchid: 0xFF9932CC, darkred: 0xFF8B0000, darksalmon: 0xFFE9967A, darkseagreen: 0xFF8FBC8F, darkslateblue: 0xFF483D8B,
    darkslategray: 0xFF2F4F4F, darkturquoise: 0xFF00CED1, darkviolet: 0xFF9400D3, deeppink: 0xFFFF1493, deepskyblue: 0xFF00BFFF,
    dimgray: 0xFF696969, dodgerblue: 0xFF1E90FF, firebrick: 0xFFB22222, floralwhite: 0xFFFFFAF0, forestgreen: 0xFF228B22,
    fuchsia: 0xFFFF00FF, gainsboro: 0xFFDCDCDC, ghostwhite: 0xFFF8F8FF, gold: 0xFFFFD700, goldenrod: 0xFFDAA520,
    gray: 0xFF808080, green: 0xFF008000, greenyellow: 0xFFADFF2F, honeydew: 0xFFF0FFF0, hotpink: 0xFFFF69B4,
    indianred: 0xFFCD5C5C, indigo: 0xFF4B0082, ivory: 0xFFFFFFF0, khaki: 0xFFF0E68C, lavender: 0xFFE6E6FA,
    lavenderblush: 0xFFFFF0F5, lawngreen: 0xFF7CFC00, lemonchiffon: 0xFFFFFACD, lightblue: 0xFFADD8E6, lightcoral: 0xFFF08080,
    lightcyan: 0xFFE0FFFF, lightgoldenrodyellow: 0xFFFAFAD2, lightgray: 0xFFD3D3D3, lightgreen: 0xFF90EE90, lightpink: 0xFFFFB6C1,
    lightsalmon: 0xFFFFA07A, lightseagreen: 0xFF20B2AA, lightskyblue: 0xFF87CEFA, lightslategray: 0xFF778899, lightsteelblue: 0xFFB0C4DE,
    lightyellow: 0xFFFFFFE0, lime: 0xFF00FF00, limegreen: 0xFF32CD32, linen: 0xFFFAF0E6, magenta: 0xFFFF00FF,
    maroon: 0xFF800000, mediumaquamarine: 0xFF66CDAA, mediumblue: 0xFF0000CD, mediumorchid: 0xFFBA55D3, mediumpurple: 0xFF9370DB,
    mediumseagreen: 0xFF3CB371, mediumslateblue: 0xFF7B68EE, mediumspringgreen: 0xFF00FA9A, mediumturquoise: 0xFF48D1CC, mediumvioletred: 0xFFC71585,
    midnightblue: 0xFF191970, mintcream: 0xFFF5FFFA, mistyrose: 0xFFFFE4E1, moccasin: 0xFFFFE4B5, navajowhite: 0xFFFFDEAD,
    navy: 0xFF000080, oldlace: 0xFFFDF5E6, olive: 0xFF808000, olivedrab: 0xFF6B8E23, orange: 0xFFFFA500,
    orangered: 0xFFFF4500, orchid: 0xFFDA70D6, palegoldenrod: 0xFFEEE8AA, palegreen: 0xFF98FB98, paleturquoise: 0xFFAFEEEE,
    palevioletred: 0xFFDB7093, papayawhip: 0xFFFFEFD5, peachpuff: 0xFFFFDAB9, peru: 0xFFCD853F, pink: 0xFFFFC0CB,
    plum: 0xFFDDA0DD, powderblue: 0xFFB0E0E6, purple: 0xFF800080, red: 0xFFFF0000, rosybrown: 0xFFBC8F8F,
    royalblue: 0xFF4169E1, saddlebrown: 0xFF8B4513, salmon: 0xFFFA8072, sandybrown: 0xFFF4A460, seagreen: 0xFF2E8B57,
    seashell: 0xFFFFF5EE, sienna: 0xFFA0522D, silver: 0xFFC0C0C0, skyblue: 0xFF87CEEB, slateblue: 0xFF6A5ACD,
    slategray: 0xFF708090, snow: 0xFFFFFAFA, springgreen: 0xFF00FF7F, steelblue: 0xFF4682B4, tan: 0xFFD2B48C,
    teal: 0xFF008080, thistle: 0xFFD8BFD8, tomato: 0xFFFF6347, transparent: 0x00FFFFFF, turquoise: 0xFF40E0D0,
    violet: 0xFFEE82EE, wheat: 0xFFF5DEB3, white: 0xFFFFFFFF, whitesmoke: 0xFFF5F5F5, yellow: 0xFFFFFF00,
    yellowgreen: 0xFF9ACD32
};

/** A flat color (SolidColorBrush) */
class SolidColorBrush {
    constructor(color) {
        this.color = color;
    }

    get isSolid() {
        return true;
    }

    equals(other) {
        return other instanceof SolidColorBrush && this.color.equals(other.color);
    }

    toString() {
        return ColorText.toArgbHex(this.color);
    }
}

/**
 * A linear gradient: start and end relative to the box of whatever it fills (0,0 the top
 * left, 1,1 the bottom right), unless `absolute`, and stops at offsets 0..1
 */
class LinearGradientBrush {
    constructor(startPoint = new Point(0, 0), endPoint = new Point(1, 1), stops = []) {
        this.startPoint = startPoint;
        this.endPoint = endPoint;
        this.gradientStops = stops;
        this.absolute = false;
    }

    get isSolid() {
        return false;
    }

    equals(other) {
        return other instanceof LinearGradientBrush && this.toString() === other.toString();
    }

    toString() {
        return "gradient " + this.startPoint + "-" + this.endPoint + " "
            + this.gradientStops.map(stop => ColorText.toArgbHex(stop.color) + "@" + stop.offset).join(" ");
    }
}

class GradientStop {
    constructor(color, offset) {
        this.color = color;
        this.offset = offset;
    }
}

/** Drawing.IsWhite: a paper read from a file that is plain white counts as no paper of its own */
function isWhiteBrush(brush) {
    return brush instanceof SolidColorBrush && brush.color.equals(Color.white);
}
