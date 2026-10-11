// Port of Main/Avalonia/DynamicGeometry/Figures/Coordinates/AxisLabelsCollection.cs: the
// numbers along the axes. Each label is { text, topLeft, bold } in pixels.

class AxisLabelsCollection extends FigureBase {
    constructor() {
        super();
        this.layer = ZOrder.Grid;
        this.xAxisLabels = [];
        this.yAxisLabels = [];
    }

    get font() {
        const resolved = this.style?.resolve(AppTheme.of(this.drawing).name);
        return resolved != null && resolved.toCanvasFont != null ? resolved.toCanvasFont() : "12px " + Fonts.family;
    }

    boldFont() {
        const resolved = this.style?.resolve(AppTheme.of(this.drawing).name);
        const size = resolved != null ? resolved.fontSize : 12;
        return "bold " + size + "px " + Fonts.family;
    }

    updateVisual() {
        const coordinateSystem = this.drawing.coordinateSystem;
        const canvas = this.canvas;
        if (canvas == null) {
            return;
        }

        // the values are rounded (CoordinateSystem.gridValue), so whole numbers are exact
        const fractional = coordinateSystem.majorGridStep < 1;
        const font = this.font;
        const bold = this.boldFont();
        this.xAxisLabels = coordinateSystem.getVisibleXPoints().map(x => {
            const emphasis = fractional && x === Math.round(x);
            const text = NumberFormat.toString(x);
            const size = canvas.measureText(text, emphasis ? bold : font, 0);
            let coordinates = coordinateSystem.toPhysical(new Point(x, 0)).offset(-size.width / 2, 0);
            if (equalsWithPrecision(x, 0)) {
                coordinates = coordinates.offset(-size.width / 2 - 2, 0);
            }

            return { text, topLeft: coordinates, bold: emphasis, layout: size };
        });
        this.yAxisLabels = coordinateSystem.getVisibleYPoints().filter(y => Math.abs(y) > 0.001).map(y => {
            const emphasis = fractional && y === Math.round(y);
            const text = NumberFormat.toString(y);
            const size = canvas.measureText(text, emphasis ? bold : font, 0);
            const coordinates = coordinateSystem.toPhysical(new Point(0, y)).plus(new Point(-size.width - 2, -size.height / 2));
            return { text, topLeft: coordinates, bold: emphasis, layout: size };
        });
    }

    hitTest(point) {
        return null;
    }

    applyStyle() {
    }

    render(renderer) {
        if (!this.visible) {
            return;
        }

        const resolved = this.style?.resolve(AppTheme.of(this.drawing).name);
        const color = resolved?.color ?? Axis.colorOf(this.drawing);
        const font = this.font;
        const bold = this.boldFont();
        for (const label of this.xAxisLabels.concat(this.yAxisLabels)) {
            renderer.drawText(label.layout, label.topLeft, label.bold ? bold : font, color, 0, null, false);
        }
    }
}
