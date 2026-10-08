// Port of Main/Avalonia/DynamicGeometry/Figures/Controls/PointLabel.cs: a point's name
// (and/or coordinates), kept in an orbit around the point.

class PointLabel extends Measurement {
    /** The least room between the point's rim and the letters, in pixels */
    static Clearance = 0;

    /** In pixels: what the orbit and the first place of a label add to the size of the text and the point */
    static Margin = 5;

    constructor() {
        super();
        this.placed = false;
        this.showNameValue = false;
        this.showCoordinatesValue = false;
    }

    getPoint() {
        return this.dependencies[0];
    }

    onAddingToDrawing(drawing) {
        super.onAddingToDrawing(drawing);
        const point = this.getPoint();
        if (point.label == null) {
            point.label = this;
        }
    }

    allowMove() {
        return !this.locked;
    }

    get anchor() {
        return this.point(0);
    }

    /**
     * Keeps the label in orbit around its point: the center of the label may not get further
     * from the point than the label's larger dimension plus the point radius, and no part of
     * the letters may come closer to the point than Clearance.
     */
    clampPosition(newPosition) {
        const size = this.measureSize();
        const width = size.width;
        const height = size.height;
        if (width === 0 || height === 0) {
            return newPosition;
        }

        const point = this.toPhysical(this.anchor);
        const halfSize = new Point(width / 2, height / 2);
        const pointRadius = this.getPoint().pointSize / 2;
        const radius = Math.max(width, height) + pointRadius + PointLabel.Margin;
        let fromPoint = this.toPhysical(newPosition).plus(halfSize).minus(point);
        fromPoint = fromPoint.trimToMaxLength(radius);

        // the letters are kept clear, not the line box around them
        const ink = this.getInkBounds(width, height);
        const inkShift = ink.center.minus(halfSize);
        const inkHalfSize = new Point(ink.width / 2, ink.height / 2);
        fromPoint = PointLabel.pushClear(fromPoint.plus(inkShift), inkHalfSize, pointRadius + PointLabel.Clearance).minus(inkShift);
        return this.toLogical(point.plus(fromPoint).minus(halfSize));
    }

    /** Where the letters are in the label, in pixels from its top left corner; the whole label if the text doesn't say */
    getInkBounds(width, height) {
        const whole = new Rect(0, 0, width, height);
        const layout = this.getTextLayout();
        if (layout == null || layout.lines.length !== 1) {
            return whole;
        }

        const line = layout.lines[0];
        if (line.ink == null) {
            return whole;
        }

        const ink = new Rect(line.ink.left, line.ink.top, line.ink.right - line.ink.left, line.ink.bottom - line.ink.top)
            .translate(new Point(this.padding, this.padding));
        if (!(ink.width > 0 && ink.height > 0) || !whole.contains(ink.center)) {
            return whole;
        }

        return ink;
    }

    /** Moves the center of a box out along its own direction from the origin until no part of the box is nearer to the origin than the minimum */
    static pushClear(center, halfSize, minimum) {
        if (center.x === 0 && center.y === 0) {
            center = new Point(0, -1);
        }

        if (PointLabel.gap(center, halfSize) >= minimum) {
            return center;
        }

        let low = 1;
        let high = 2;
        while (PointLabel.gap(center.scale(high), halfSize) < minimum) {
            high *= 2;
        }

        for (let i = 0; i < 20; i++) {
            const middle = (low + high) / 2;
            if (PointLabel.gap(center.scale(middle), halfSize) < minimum) {
                low = middle;
            } else {
                high = middle;
            }
        }

        return center.scale(high);
    }

    /** The distance from the origin to the nearest point of a box */
    static gap(center, halfSize) {
        const dx = Math.max(Math.abs(center.x) - halfSize.x, 0);
        const dy = Math.max(Math.abs(center.y) - halfSize.y, 0);
        return Math.sqrt(dx * dx + dy * dy);
    }

    moveToCore(newPosition) {
        super.moveToCore(this.clampPosition(newPosition));
    }

    defaultZOrder() {
        return ZOrder.PointLabels;
    }

    updateVisual() {
        if (this.dependencies.length === 0) {
            return;
        }

        this.updateText();
        if (!this.placed && this.text !== "") {
            // a new label: centered above its point, just clear of it
            this.placed = true;
            const size = this.measureSize();
            this.offset = new Point(-size.width / 2, -(size.height + this.getPoint().pointSize / 2 + PointLabel.Margin));
        }

        super.updateVisual();
    }

    updateText() {
        let text = "";
        if (this.showName) {
            // A_1 shows as A₁
            text = NameDisplay.format(this.dependencies[0].name);
        }

        if (this.showCoordinates) {
            const coordinates = this.point(0);
            const x = GeometryMath.round(coordinates.x, this.decimalsToShow);
            const y = GeometryMath.round(coordinates.y, this.decimalsToShow);
            const coordinatesText = "(" + NumberFormat.toString(x) + ", " + NumberFormat.toString(y) + ")";
            if (text !== "") {
                text += " ";
            }

            text += coordinatesText;
        }

        this.setText(text);
    }

    get showName() {
        return this.showNameValue;
    }

    set showName(value) {
        this.showNameValue = value;
        this.updateVisual();
    }

    get showCoordinates() {
        return this.showCoordinatesValue;
    }

    set showCoordinates(value) {
        this.showCoordinatesValue = value;
        this.updateVisual();
    }

    readXml(element) {
        super.readXml(element);
        this.showNameValue = Xml.readBool(element, "ShowName", true);
        this.showCoordinatesValue = Xml.readBool(element, "ShowCoordinates", false);
        this.placed = true;
        this.updateText();
    }
}

FigureTypes.register("PointLabel", PointLabel);
