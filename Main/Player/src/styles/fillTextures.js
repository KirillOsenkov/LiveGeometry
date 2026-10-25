// Port of Main/Avalonia/LiveGeometry/Styles/FillTextures.cs: fills that look laid down by hand
// rather than poured. Each texture's tile is a scan of a hand-colored area (CrayonTile,
// FeltTipTile, written by tools/texturetile.cs): how much pigment each pixel has against the
// average. A fill of a color is the paper moved toward the color by that much, so it averages
// to the color, its pale spots fade to the paper and where the pen or pencil went over again it
// deepens.

/** ShapeStyle.Texture: how the fill is laid down */
const FillTexture = {
    None: "None",
    /** With a colored pencil: the paper's tooth shows through in flecks */
    Crayon: "Crayon",
    /** With a felt-tip pen: smooth ink, soft blotches and the bands of the strokes */
    FeltTip: "FeltTip"
};

/** A fill of one color with a texture (ShapeStyle.FillBrush when the style has a Texture) */
class TexturedBrush {
    constructor(color, texture) {
        this.color = color;
        this.texture = texture;
    }

    get isSolid() {
        return false;
    }

    equals(other) {
        return other instanceof TexturedBrush && this.color.equals(other.color) && this.texture === other.texture;
    }

    toString() {
        return this.texture + " " + ColorText.toArgbHex(this.color);
    }

    /** The one color of a fill: a solid brush's, the middle of a gradient's stops */
    static colorOf(brush) {
        if (brush instanceof SolidColorBrush) {
            return brush.color;
        }

        const stops = brush.gradientStops ?? [];
        if (stops.length === 0) {
            return Color.transparent;
        }

        const first = stops[0].color;
        const last = stops[stops.length - 1].color;
        return Color.fromArgb(Math.round((first.a + last.a) / 2), Math.round((first.r + last.r) / 2), Math.round((first.g + last.g) / 2), Math.round((first.b + last.b) / 2));
    }
}

const FillTextures = {
    /** How much of the canvas a pixel of a tile covers: a pixel each (drawn smaller, it blurred and lost its grain) */
    TileScale: 1,

    tiles: new Map(),
    densities: new Map(),

    /** The scanned tile of a texture: its size and its densities as base64 */
    source(texture) {
        return texture === FillTexture.FeltTip ? FeltTipTile : CrayonTile;
    },

    /** The tile's densities, 0..255, 128 the average */
    density(texture) {
        let densities = this.densities.get(texture);
        if (densities == null) {
            const text = atob(this.source(texture).data);
            densities = new Uint8Array(text.length);
            for (let i = 0; i < text.length; i++) {
                densities[i] = text.charCodeAt(i);
            }

            this.densities.set(texture, densities);
        }

        return densities;
    },

    /**
     * A canvas with the texture's tile in the color on the paper (white when none is given),
     * for a pattern: opaque, a pale spot shows the paper and not whatever is under the shape,
     * as each region of a hand-colored drawing is colored on the paper itself
     */
    tile(texture, color, paper = null) {
        paper ??= Color.white;
        const key = texture + " " + ColorText.toArgbHex(color) + " on " + ColorText.toArgbHex(paper);
        let tile = this.tiles.get(key);
        if (tile != null) {
            return tile;
        }

        const density = this.density(texture);
        const size = this.source(texture).size;
        tile = document.createElement("canvas");
        tile.width = size;
        tile.height = size;
        const context = tile.getContext("2d");
        const image = context.createImageData(size, size);
        const data = image.data;
        // a translucent color is that much less of a step from the paper
        const strength = color.a / 255;
        for (let i = 0; i < density.length; i++) {
            const amount = density[i] / 128 * strength;
            data[4 * i] = paper.r + (color.r - paper.r) * amount;
            data[4 * i + 1] = paper.g + (color.g - paper.g) * amount;
            data[4 * i + 2] = paper.b + (color.b - paper.b) * amount;
            data[4 * i + 3] = 255;
        }

        context.putImageData(image, 0, 0);
        if (this.tiles.size > 200) {
            this.tiles.clear();
        }

        this.tiles.set(key, tile);
        return tile;
    }
};
