// Text laid out for a label: what Avalonia's TextBlock measures in the app. An offscreen 2D
// context measures the text in the font, wraps it at a width (greedy, at spaces, a word
// longer than the width broken by characters), and gives the lines with their widths, the
// box of the whole, and for a single line where the letters are (the ink, for a point
// label's orbit).

class TextMeasurer {
    constructor() {
        this.context = document.createElement("canvas").getContext("2d");
        this.fontMetrics = new Map();
    }

    /** The line height and the baseline's distance from the top of a line, for the font */
    metricsOf(font) {
        let metrics = this.fontMetrics.get(font);
        if (metrics == null) {
            this.context.font = font;
            const measured = this.context.measureText("Hg");
            const ascent = measured.fontBoundingBoxAscent;
            const descent = measured.fontBoundingBoxDescent;
            if (isValidValue(ascent) && isValidValue(descent) && ascent + descent > 0) {
                metrics = { ascent, lineHeight: ascent + descent };
            } else {
                const size = TextMeasurer.fontSize(font);
                metrics = { ascent: size * 0.95, lineHeight: size * 1.2 };
            }

            this.fontMetrics.set(font, metrics);
        }

        return metrics;
    }

    static fontSize(font) {
        const match = /(\d+(\.\d+)?)px/.exec(font);
        return match != null ? Number(match[1]) : 16;
    }

    /** The layout: { text, font, wrapWidth, lines: [{ text, width, ink }], width, height, lineHeight, ascent } */
    layout(text, font, wrapWidth) {
        const context = this.context;
        context.font = font;
        const metrics = this.metricsOf(font);
        const lines = [];
        for (const paragraph of (text ?? "").split("\n")) {
            if (wrapWidth > 0) {
                for (const line of this.wrap(paragraph, wrapWidth)) {
                    lines.push(line);
                }
            } else {
                lines.push(paragraph);
            }
        }

        const measured = lines.map(line => {
            const m = context.measureText(line);
            return {
                text: line,
                width: m.width,
                ink: {
                    left: -m.actualBoundingBoxLeft,
                    right: m.actualBoundingBoxRight,
                    top: metrics.ascent - m.actualBoundingBoxAscent,
                    bottom: metrics.ascent + m.actualBoundingBoxDescent
                }
            };
        });
        const width = wrapWidth > 0 ? wrapWidth : Math.max(0, ...measured.map(line => line.width));
        return {
            text,
            font,
            wrapWidth,
            lines: measured,
            width,
            height: measured.length * metrics.lineHeight,
            lineHeight: metrics.lineHeight,
            ascent: metrics.ascent
        };
    }

    /** The paragraph in lines no wider than the width, broken at spaces; a word that doesn't fit is broken by characters */
    wrap(paragraph, width) {
        const context = this.context;
        const lines = [];
        let line = "";
        const words = paragraph.split(" ");
        for (let i = 0; i < words.length; i++) {
            const word = words[i];
            const candidate = line === "" ? word : line + " " + word;
            if (context.measureText(candidate).width <= width || line === "" && word === "") {
                line = candidate;
                continue;
            }

            if (line !== "") {
                lines.push(line);
                line = "";
            }

            if (context.measureText(word).width <= width) {
                line = word;
                continue;
            }

            // a word wider than the line: as many characters as fit, then the rest
            let piece = "";
            for (const character of word) {
                if (context.measureText(piece + character).width > width && piece !== "") {
                    lines.push(piece);
                    piece = "";
                }

                piece += character;
            }

            line = piece;
        }

        lines.push(line);
        return lines;
    }
}
