// Port of Main/Avalonia/DynamicGeometry/Figures/Lines/AxisLine.cs: an axis of the grid as a
// line to build on. The grid draws the axes; this line is never drawn itself. A drawing has
// one of each for its whole life (Drawing.getAxisLine).

const AxisDirection = {
    X: "X",
    Y: "Y"
};

class AxisLine extends LineBase {
    constructor(direction) {
        super();
        this.direction = direction;
        this.nameValue = direction === AxisDirection.X ? "x-axis" : "y-axis";
        this.auxiliary = true;
    }

    get isLine() {
        return true;
    }

    get isAxisLine() {
        return true;
    }

    defaultZOrder() {
        return ZOrder.Axes;
    }

    get coordinates() {
        const direction = this.direction === AxisDirection.X ? new Point(1, 0) : new Point(0, 1);
        return new PointPair(new Point(0, 0), direction);
    }

    get onScreenCoordinates() {
        return GeometryMath.getLineFromSegment(this.coordinates, this.canvasLogicalBorders);
    }

    /** A click takes the axis only where the grid shows it */
    get isHitTestVisible() {
        return this.drawing != null && this.drawing.coordinateGrid != null && this.drawing.coordinateGrid.showsAxes;
    }

    set isHitTestVisible(value) {
    }

    /** The grid draws the axis: the line itself draws nothing */
    applyStyle() {
        if (this.drawing != null) {
            this.updateVisual();
        }
    }

    get strokeThickness() {
        return 1;
    }

    render(renderer) {
    }

    static readDirection(element) {
        return element.getAttribute("Axis") === "Y" ? AxisDirection.Y : AxisDirection.X;
    }
}

FigureTypes.register("AxisLine", AxisLine);
