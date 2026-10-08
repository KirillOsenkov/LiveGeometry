// Port of Main/Avalonia/DynamicGeometry/Figures/Controls/FigureLabel.cs: the name of a line,
// ray, segment or circle, written next to it, at Offset pixels from an anchor on it.

class FigureLabel extends Measurement {
    /** How far in from the edge of the window a line's name sits, in pixels */
    static EdgeInset = 24;

    /** Between a segment and its name, in pixels */
    static SegmentGap = 6;

    /** The least room between a line's name and the edge of the window, in pixels */
    static EdgeMargin = 4;

    constructor() {
        super();
        this.placed = false;
        this.shown = true;
    }

    get figure() {
        return this.dependencies.length > 0 ? this.dependencies[0] : null;
    }

    onAddingToDrawing(drawing) {
        super.onAddingToDrawing(drawing);
        const figure = this.figure;
        if (figure instanceof FigureBase && figure.nameLabel == null) {
            figure.nameLabel = this;
        }
    }

    onRemovingFromDrawing(drawing) {
        super.onRemovingFromDrawing(drawing);
        const figure = this.figure;
        if (figure instanceof FigureBase && figure.nameLabel === this) {
            figure.nameLabel = null;
        }
    }

    defaultZOrder() {
        return ZOrder.PointLabels;
    }

    /** Where the offset is measured from: on the figure, and for a line where the eye finds it */
    get anchor() {
        const figure = this.figure;
        if (figure instanceof Segment) {
            return figure.coordinates.midpoint;
        }

        if (figure instanceof Ray) {
            return this.inward(figure.onScreenCoordinates.p2, figure.onScreenCoordinates.p1);
        }

        if (figure instanceof LineBase) {
            const shown = figure.onScreenCoordinates;
            return shown.p1.y >= shown.p2.y ? this.inward(shown.p1, shown.p2) : this.inward(shown.p2, shown.p1);
        }

        if (figure != null && figure.isCircle === true) {
            const center = figure.center;
            const radius = figure.radius * Math.sqrt(0.5);
            return new Point(center.x - radius, center.y + radius);
        }

        return figure != null ? figure.center : new Point();
    }

    /** The point EdgeInset pixels from the end towards the other end, logical */
    inward(end, other) {
        if (this.drawing == null) {
            return end;
        }

        const from = this.toPhysical(end);
        const direction = RightAngleMark.direction(from, this.toPhysical(other));
        if (direction == null) {
            return end;
        }

        return this.toLogical(from.plus(direction.scale(FigureLabel.EdgeInset)));
    }

    updateVisual() {
        const figure = this.figure;
        if (figure == null || this.drawing == null) {
            return;
        }

        this.setText(NameDisplay.format(figure.name));
        if (!this.placed && this.text !== "") {
            this.placed = true;
            this.offset = this.defaultOffset(figure);
        }

        super.updateVisual();
        let shown = this.visible && figure.visible;
        if (figure instanceof LineBase && !(figure instanceof Segment)) {
            shown = shown && this.isOnScreen(this.toPhysical(this.anchor));
            if (shown) {
                this.keepOnScreen();
            }
        }

        this.shown = shown;
    }

    get isShown() {
        return super.isShown && this.shown;
    }

    isOnScreen(pixel) {
        const canvas = this.drawing.coordinateSystem.physicalSize;
        return pixel.exists()
            && pixel.x >= -FigureLabel.EdgeInset && pixel.x <= canvas.x + FigureLabel.EdgeInset
            && pixel.y >= -FigureLabel.EdgeInset && pixel.y <= canvas.y + FigureLabel.EdgeInset;
    }

    /** A line's anchor sits by the edge of the window, and the offset may put the name beyond it: the name is pushed back in, whole */
    keepOnScreen() {
        const canvas = this.drawing.coordinateSystem.physicalSize;
        const size = this.measureSize();
        const topLeft = this.toPhysical(this.coordinates);
        const kept = new Point(
            FigureLabel.clamp(topLeft.x, FigureLabel.EdgeMargin, canvas.x - size.width - FigureLabel.EdgeMargin),
            FigureLabel.clamp(topLeft.y, FigureLabel.EdgeMargin, canvas.y - size.height - FigureLabel.EdgeMargin));
        if (!kept.equals(topLeft) && kept.exists()) {
            this.coordinates = this.toLogical(kept);
        }
    }

    static clamp(value, low, high) {
        return high < low ? low : Math.min(Math.max(value, low), high);
    }

    /** Where a new name goes, in pixels from the anchor */
    defaultOffset(figure) {
        const size = this.measureSize();
        if (figure instanceof LineBase) {
            const coordinates = figure instanceof Segment ? figure.coordinates : figure.onScreenCoordinates;
            const direction = RightAngleMark.direction(this.toPhysical(coordinates.p1), this.toPhysical(coordinates.p2)) ?? new Point(1, 0);
            const normal = new Point(direction.y, -direction.x);
            const distance = FigureLabel.SegmentGap + Math.abs(normal.x) * size.width / 2 + Math.abs(normal.y) * size.height / 2;
            return normal.scale(distance).minus(new Point(size.width / 2, size.height / 2));
        }

        if (figure.isCircle === true) {
            return new Point(4, 2);
        }

        return new Point(6, 2);
    }

    readXml(element) {
        super.readXml(element);
        this.placed = true;
    }
}

FigureTypes.register("FigureLabel", FigureLabel);
