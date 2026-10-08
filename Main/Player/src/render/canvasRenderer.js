// What a figure draws with: the one renderer, a 2D canvas context. Everything is in pixels
// (CSS pixels; the canvas is scaled by the device pixel ratio once per frame). A stroke is
// { color, width, dash, join? } with a Color; a fill is a SolidColorBrush or a
// LinearGradientBrush, whose relative points span the box of what it fills.
//
// The methods are the renderer interface another renderer (SVG) would implement:
// beginFrame, endFrame, drawLine, drawPolyline, drawPolygon, drawEllipse, drawPath,
// drawMarker, drawText, drawCheckBox.

class CanvasRenderer {
    constructor(context) {
        this.context = context;
    }

    beginFrame(width, height, pixelRatio, paper) {
        const context = this.context;
        context.setTransform(pixelRatio, 0, 0, pixelRatio, 0, 0);
        context.clearRect(0, 0, width, height);
        if (paper != null) {
            context.fillStyle = this.toFillStyle(paper, new Rect(0, 0, width, height));
            context.fillRect(0, 0, width, height);
        }
    }

    endFrame() {
    }

    /** A brush as the context takes it, a gradient laid across the box */
    toFillStyle(brush, box) {
        if (brush instanceof SolidColorBrush) {
            return brush.color.toCss();
        }

        if (brush instanceof LinearGradientBrush) {
            const start = brush.absolute ? brush.startPoint : new Point(box.x + brush.startPoint.x * box.width, box.y + brush.startPoint.y * box.height);
            const end = brush.absolute ? brush.endPoint : new Point(box.x + brush.endPoint.x * box.width, box.y + brush.endPoint.y * box.height);
            if (start.equals(end)) {
                return brush.gradientStops.length > 0 ? brush.gradientStops[brush.gradientStops.length - 1].color.toCss() : "transparent";
            }

            const gradient = this.context.createLinearGradient(start.x, start.y, end.x, end.y);
            for (const stop of brush.gradientStops) {
                gradient.addColorStop(Math.max(0, Math.min(1, stop.offset)), stop.color.toCss());
            }

            return gradient;
        }

        return "transparent";
    }

    applyStroke(stroke) {
        const context = this.context;
        context.strokeStyle = stroke.color.toCss();
        context.lineWidth = stroke.width;
        context.setLineDash(stroke.dash ?? []);
        context.lineJoin = stroke.join ?? "round";
        context.lineCap = stroke.cap ?? "butt";
    }

    static hasStroke(stroke) {
        return stroke != null && stroke.color != null && stroke.color.a > 0 && stroke.width > 0;
    }

    static boxOf(points) {
        let minX = Infinity;
        let minY = Infinity;
        let maxX = -Infinity;
        let maxY = -Infinity;
        for (const p of points) {
            minX = Math.min(minX, p.x);
            minY = Math.min(minY, p.y);
            maxX = Math.max(maxX, p.x);
            maxY = Math.max(maxY, p.y);
        }

        return new Rect(minX, minY, maxX - minX, maxY - minY);
    }

    drawLine(p1, p2, stroke) {
        if (!CanvasRenderer.hasStroke(stroke)) {
            return;
        }

        const context = this.context;
        this.applyStroke(stroke);
        context.beginPath();
        context.moveTo(p1.x, p1.y);
        context.lineTo(p2.x, p2.y);
        context.stroke();
    }

    drawPolyline(points, stroke) {
        if (!CanvasRenderer.hasStroke(stroke) || points.length < 2) {
            return;
        }

        const context = this.context;
        this.applyStroke(stroke);
        context.beginPath();
        context.moveTo(points[0].x, points[0].y);
        for (let i = 1; i < points.length; i++) {
            context.lineTo(points[i].x, points[i].y);
        }

        context.stroke();
    }

    /** A polygon filled with the fill (even-odd, as Avalonia fills) and outlined with the stroke */
    drawPolygon(points, stroke, fill, closed = true) {
        if (points.length < 2) {
            return;
        }

        const context = this.context;
        context.beginPath();
        context.moveTo(points[0].x, points[0].y);
        for (let i = 1; i < points.length; i++) {
            context.lineTo(points[i].x, points[i].y);
        }

        if (closed) {
            context.closePath();
        }

        if (fill != null) {
            context.fillStyle = this.toFillStyle(fill, CanvasRenderer.boxOf(points));
            context.fill("evenodd");
        }

        if (CanvasRenderer.hasStroke(stroke)) {
            this.applyStroke(stroke);
            context.stroke();
        }
    }

    /** An ellipse around the center with the radii, turned by the rotation (radians, clockwise on screen) */
    drawEllipse(center, radiusX, radiusY, rotation, stroke, fill) {
        const context = this.context;
        context.beginPath();
        context.ellipse(center.x, center.y, Math.max(0, radiusX), Math.max(0, radiusY), rotation, 0, 2 * Math.PI);
        if (fill != null) {
            const reach = Math.max(radiusX, radiusY);
            context.fillStyle = this.toFillStyle(fill, new Rect(center.x - reach, center.y - reach, 2 * reach, 2 * reach));
            context.fill();
        }

        if (CanvasRenderer.hasStroke(stroke)) {
            this.applyStroke(stroke);
            context.stroke();
        }
    }

    /**
     * A path of commands: { op: "move", x, y }, { op: "line", x, y }, { op: "cubic", x1, y1,
     * x2, y2, x, y }, { op: "arc", cx, cy, rx, ry, rotation, start, end, counterclockwise },
     * { op: "close" }. Filled with the fill (even-odd), outlined with the stroke.
     */
    drawPath(commands, stroke, fill, box = null) {
        const context = this.context;
        context.beginPath();
        const points = [];
        for (const command of commands) {
            switch (command.op) {
                case "move":
                    context.moveTo(command.x, command.y);
                    points.push(new Point(command.x, command.y));
                    break;
                case "line":
                    context.lineTo(command.x, command.y);
                    points.push(new Point(command.x, command.y));
                    break;
                case "cubic":
                    context.bezierCurveTo(command.x1, command.y1, command.x2, command.y2, command.x, command.y);
                    points.push(new Point(command.x, command.y));
                    break;
                case "arc":
                    context.ellipse(command.cx, command.cy, command.rx, command.ry, command.rotation, command.start, command.end, command.counterclockwise);
                    points.push(new Point(command.cx - command.rx, command.cy - command.ry), new Point(command.cx + command.rx, command.cy + command.ry));
                    break;
                case "close":
                    context.closePath();
                    break;
            }
        }

        if (fill != null && points.length > 0) {
            context.fillStyle = this.toFillStyle(fill, box ?? CanvasRenderer.boxOf(points));
            context.fill("evenodd");
        }

        if (CanvasRenderer.hasStroke(stroke)) {
            this.applyStroke(stroke);
            context.stroke();
        }
    }

    /** A point's marker: a shape of the size around the center, or a character in the font size of the size */
    drawMarker(center, size, kind, character, stroke, fill) {
        const context = this.context;
        if (character != null && character !== "") {
            context.font = size + "px " + Fonts.family;
            context.textAlign = "center";
            context.textBaseline = "middle";
            // a color emoji is painted with the style's fill: the fill's alpha is the emoji's opacity
            context.fillStyle = fill != null ? this.toFillStyle(fill, new Rect(center.x - size / 2, center.y - size / 2, size, size)) : "black";
            context.fillText(character, center.x, center.y);
            context.textAlign = "start";
            context.textBaseline = "alphabetic";
            return;
        }

        const strokeWidth = CanvasRenderer.hasStroke(stroke) ? stroke.width : 0;
        const polygon = PointMarker.getPolygon(kind, center, size, strokeWidth);
        if (polygon != null) {
            this.drawPolygon(polygon, stroke, fill, true);
            return;
        }

        const radius = PointMarker.getRadius(size, strokeWidth);
        this.drawEllipse(center, radius, radius, 0, stroke, fill);
    }

    /** Text laid out by the measurer at the top-left, with a plate of the backdrop behind it when given */
    drawText(layout, topLeft, font, color, padding, backdrop, underline) {
        const context = this.context;
        if (backdrop != null) {
            context.fillStyle = this.toFillStyle(backdrop, new Rect(topLeft.x, topLeft.y, layout.width + 2 * padding, layout.height + 2 * padding));
            context.fillRect(topLeft.x, topLeft.y, layout.width + 2 * padding, layout.height + 2 * padding);
        }

        context.font = font;
        context.fillStyle = color.toCss();
        context.textAlign = "start";
        context.textBaseline = "alphabetic";
        let y = topLeft.y + padding + layout.ascent;
        for (const line of layout.lines) {
            context.fillText(line.text, topLeft.x + padding, y);
            if (underline) {
                context.fillRect(topLeft.x + padding, y + 1, line.width, Math.max(1, TextMeasurer.fontSize(font) / 14));
            }

            y += layout.lineHeight;
        }
    }

    /** A check box with a caption, as a show/hide box draws: the box outlined (filled when checked) in the caption's color */
    drawCheckBox(topLeft, size, isChecked, layout, font, color, boxSize, gap) {
        const context = this.context;
        const box = new Rect(topLeft.x, topLeft.y + (size.height - boxSize) / 2, boxSize, boxSize);
        context.beginPath();
        context.roundRect(box.x + 0.5, box.y + 0.5, box.width - 1, box.height - 1, 3);
        if (isChecked) {
            context.fillStyle = color.toCss();
            context.fill();
            context.strokeStyle = color.toCss();
            context.lineWidth = 1;
            context.setLineDash([]);
            context.stroke();

            // the tick, in the paper's white
            context.beginPath();
            context.moveTo(box.x + boxSize * 0.25, box.y + boxSize * 0.52);
            context.lineTo(box.x + boxSize * 0.43, box.y + boxSize * 0.7);
            context.lineTo(box.x + boxSize * 0.76, box.y + boxSize * 0.32);
            context.strokeStyle = "white";
            context.lineWidth = 2;
            context.lineCap = "round";
            context.lineJoin = "round";
            context.stroke();
        } else {
            context.strokeStyle = color.toCss();
            context.lineWidth = 1;
            context.setLineDash([]);
            context.stroke();
        }

        if (layout != null && layout.text !== "") {
            const textTop = new Point(topLeft.x + boxSize + gap, topLeft.y + (size.height - layout.height) / 2);
            this.drawText(layout, textTop, font, color, 0, null, false);
        }
    }
}
