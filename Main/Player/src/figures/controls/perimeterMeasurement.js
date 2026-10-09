// Port of Main/Avalonia/DynamicGeometry/Figures/Controls/PerimeterMeasurement.cs: the perimeter
// of a shape, written beside it. A shape with a perimeter says isPerimeterProvider and has a
// perimeter getter (NaN where it has none right now: an open path).

class PerimeterMeasurement extends Measurement {
    /** Between the outline and the number, in pixels */
    static Gap = 6;

    /** How far past the edge of the window a point still counts as on screen, in pixels */
    static EdgeInset = 24;

    constructor() {
        super();
        this.prefix = "";
        this.placed = false;
    }

    get isLengthProvider() {
        return true;
    }

    get measured() {
        const figure = this.dependencies.length > 0 ? this.dependencies[0] : null;
        return figure != null && figure.isPerimeterProvider === true ? figure : null;
    }

    /** A shape with a perimeter right now; anything else has nothing to measure */
    get hasSomethingToMeasure() {
        const measured = this.measured;
        return measured != null && isValidValue(measured.perimeter);
    }

    updateExistence() {
        super.updateExistence();
        if (this.exists && !this.hasSomethingToMeasure) {
            this.exists = false;
        }
    }

    get length() {
        return this.hasSomethingToMeasure ? this.measured.perimeter : NaN;
    }

    /**
     * A point of the outline: the middle of a polygon's first side, the lower right of a circle
     * or an ellipse, the middle of a sector's or segment's arc, the middle of a path's first
     * piece. A circle or an ellipse too big for the window: the point of it nearest the middle
     * of the window.
     */
    get anchor() {
        const measured = this.measured;
        if (measured == null) {
            return Point.infinite;
        }

        if (measured.isPolygonalChain === true) {
            const vertices = measured.vertexCoordinates;
            return vertices != null && vertices.length >= 2
                ? GeometryMath.midpoint(vertices[0], vertices[1])
                : measured.center;
        }

        if (measured.isArc === true) {
            return measured.arcMiddle;
        }

        if (measured.isEllipse === true) {
            return this.onScreenPoint(measured);
        }

        if (measured instanceof BezierPath) {
            return measured.getPointFromParameter(0.5);
        }

        return measured.center;
    }

    onScreenPoint(ellipse) {
        const center = ellipse.center;
        const fixedPoint = GeometryMath.pointOnEllipse(center, ellipse.semiMajor, ellipse.semiMinor, ellipse.inclination, -Math.PI / 4);
        if (this.drawing == null || this.isOnScreen(this.toPhysical(fixedPoint))) {
            return fixedPoint;
        }

        const canvas = this.drawing.coordinateSystem.physicalSize;
        const middle = this.toLogical(new Point(canvas.x / 2, canvas.y / 2));
        if (!middle.exists() || middle.distance(center) < 1e-9) {
            return fixedPoint;
        }

        const crossings = GeometryMath.getIntersectionOfEllipseAndLine(center, ellipse.semiMajor, ellipse.semiMinor, ellipse.inclination, new PointPair(center, middle));
        if (!crossings.p1.exists() || !crossings.p2.exists()) {
            return fixedPoint;
        }

        return crossings.p1.distance(middle) <= crossings.p2.distance(middle) ? crossings.p1 : crossings.p2;
    }

    isOnScreen(pixel) {
        const canvas = this.drawing.coordinateSystem.physicalSize;
        return pixel.exists()
            && pixel.x >= -PerimeterMeasurement.EdgeInset && pixel.x <= canvas.x + PerimeterMeasurement.EdgeInset
            && pixel.y >= -PerimeterMeasurement.EdgeInset && pixel.y <= canvas.y + PerimeterMeasurement.EdgeInset;
    }

    updateVisual() {
        if (!this.hasSomethingToMeasure) {
            return;
        }

        this.setText(this.prefix + NumberFormat.toString(GeometryMath.round(this.length, this.decimalsToShow)));

        // (placed once the canvas has a size: before the first layout the anchor of a
        // circle could only be the fallback, and the offset worked out from it was wrong)
        if (!this.placed && this.drawing != null && this.drawing.coordinateSystem.physicalSize.x > 0) {
            this.placed = true;
            this.offset = this.defaultOffset();
        }

        super.updateVisual();
    }

    /** Where a new number goes, in pixels from the anchor: just outside the outline, clear of it */
    defaultOffset() {
        const size = this.measureSize();
        const direction = this.outwardDirection();
        const distance = PerimeterMeasurement.Gap + Math.abs(direction.x) * size.width / 2 + Math.abs(direction.y) * size.height / 2;
        return direction.scale(distance).minus(new Point(size.width / 2, size.height / 2));
    }

    /** A unit vector in pixels pointing from the anchor out of the shape */
    outwardDirection() {
        const measured = this.measured;
        const anchor = this.toPhysical(this.anchor);
        const center = this.toPhysical(measured.center);
        if (measured.isPolygonalChain === true && measured.vertexCoordinates != null && measured.vertexCoordinates.length >= 2) {
            const vertices = measured.vertexCoordinates;
            const along = RightAngleMark.direction(this.toPhysical(vertices[0]), this.toPhysical(vertices[1]));
            if (along != null) {
                let normal = new Point(along.y, -along.x);
                if ((anchor.x - center.x) * normal.x + (anchor.y - center.y) * normal.y < 0) {
                    normal = new Point(-normal.x, -normal.y);
                }

                return normal;
            }
        }

        return RightAngleMark.direction(center, anchor) ?? new Point(0, 1);
    }

    readXml(element) {
        super.readXml(element);

        // (a file written by hand without an offset gets the default place)
        this.placed = element.hasAttribute("OffsetX");
        this.prefix = element.getAttribute("Prefix") ?? "";
    }
}

FigureTypes.register("PerimeterMeasurement", PerimeterMeasurement);
