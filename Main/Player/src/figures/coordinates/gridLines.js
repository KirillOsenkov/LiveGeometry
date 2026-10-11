// Port of Main/Avalonia/DynamicGeometry/Figures/Coordinates/GridLinesCollection.cs and
// RectangularGridLinesCollection.cs: the lines of the grid, the labeled ones and the finer
// ones between them, across the visible rectangle.

class RectangularGridLinesCollection extends FigureBase {
    constructor() {
        super();
        this.layer = ZOrder.Grid;

        /** The finer lines between the labeled ones, fainter than style */
        this.minorStyle = null;
        this.lines = [];
        this.minorLines = [];
    }

    updateVisual() {
        const coordinateSystem = this.drawing.coordinateSystem;
        this.lines = RectangularGridLinesCollection.place(coordinateSystem.getVisibleXPoints(), coordinateSystem.getVisibleYPoints(), coordinateSystem);
        this.minorLines = RectangularGridLinesCollection.place(coordinateSystem.getMinorXPoints(), coordinateSystem.getMinorYPoints(), coordinateSystem);
    }

    /** The lines in pixels: one per x across the height in view, one per y across the width */
    static place(xPoints, yPoints, coordinateSystem) {
        const lines = [];
        for (const x of xPoints) {
            lines.push(coordinateSystem.toPhysicalPair(new PointPair(
                new Point(x, coordinateSystem.minimalVisibleY),
                new Point(x, coordinateSystem.maximalVisibleY))));
        }

        for (const y of yPoints) {
            lines.push(coordinateSystem.toPhysicalPair(new PointPair(
                new Point(coordinateSystem.minimalVisibleX, y),
                new Point(coordinateSystem.maximalVisibleX, y))));
        }

        return lines;
    }

    applyStyle() {
    }

    hitTest(point) {
        return null;
    }

    render(renderer) {
        if (!this.visible) {
            return;
        }

        const themeName = AppTheme.of(this.drawing).name;
        RectangularGridLinesCollection.draw(renderer, this.minorLines, this.minorStyle, themeName);
        RectangularGridLinesCollection.draw(renderer, this.lines, this.style, themeName);
    }

    static draw(renderer, lines, style, themeName) {
        if (style == null) {
            return;
        }

        const resolved = style.resolve(themeName);
        const stroke = { color: resolved.color, width: resolved.strokeWidth, dash: null };
        for (const line of lines) {
            renderer.drawLine(line.p1, line.p2, stroke);
        }
    }
}
