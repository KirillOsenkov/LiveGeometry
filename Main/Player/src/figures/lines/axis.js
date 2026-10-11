// Port of Main/Avalonia/DynamicGeometry/Figures/Lines/Axis.cs: an axis of the coordinate
// grid, a line drawn as a vector's arrow, the shaft across the window and the head where the
// axis leaves it. The line itself is invisible.

class Axis extends CompositeFigure {
    constructor() {
        super();
        this.line = new LineTwoPoints();
        const transparent = new LineStyle();
        transparent.strokeWidth = 0;
        transparent.color = Color.transparent;
        this.line.style = transparent;
        this.line.layer = ZOrder.Axes;
        this.arrow = new Arrow();
        this.arrow.layer = ZOrder.Axes;
        this.arrow.dependencies = [this.line];
        this.children.push(this.line, this.arrow);
    }

    hitTestWith(point, filter) {
        return null;
    }

    hitTest(point) {
        return null;
    }

    /** The axis color of the theme the drawing is shown under */
    static colorOf(drawing) {
        return AppTheme.of(drawing).axis;
    }
}
