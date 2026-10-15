// Port of Main/Avalonia/DynamicGeometry/Figures/Controls/LengthMeasurement.cs: the distance
// between two points, or a length of one figure - which one is the quantity (a circle's
// radius, diameter or circumference; an arc's length or radius; a sector's arc length,
// radius or perimeter). A shape with a perimeter says isPerimeterProvider and has a perimeter
// getter (NaN where it has none right now: an open path).

const LengthQuantity = Object.freeze({
    Distance: "Distance",
    Length: "Length",
    ArcLength: "ArcLength",
    Radius: "Radius",
    Diameter: "Diameter",
    Circumference: "Circumference",
    Perimeter: "Perimeter",
});

class LengthMeasurement extends Measurement {
    /** Between the outline and the number, in pixels */
    static Gap = 6;

    /** How far past the edge of the window a point still counts as on screen, in pixels */
    static EdgeInset = 24;

    constructor() {
        super();
        this.quantityValue = null;
        this.prefixValue = null;
        this.placed = false;
    }

    get isLengthProvider() {
        return true;
    }

    /** The one figure measured; null for the distance between two points */
    get measured() {
        return this.dependencies.length === 1 ? this.dependencies[0] : null;
    }

    get isBetweenPoints() {
        return this.dependencies.length === 2 && this.dependencies[0].isPoint === true && this.dependencies[1].isPoint === true;
    }

    /** The lengths a figure has, the first being what a click on it measures */
    static quantitiesOf(figure) {
        if (figure == null || figure instanceof AngleArc) {
            return [];
        }

        if (figure.isArc === true) {
            const result = [LengthQuantity.ArcLength];
            if (figure.isCircle === true) {
                result.push(LengthQuantity.Radius);
            }

            if (figure.isPerimeterProvider === true) {
                result.push(LengthQuantity.Perimeter);
            }

            return result;
        }

        if (figure.isCircle === true) {
            return [LengthQuantity.Radius, LengthQuantity.Diameter, LengthQuantity.Circumference];
        }

        if (figure.isEllipse === true) {
            return [LengthQuantity.Circumference];
        }

        // (opened or closed after it is measured, a path's number follows)
        if (figure instanceof BezierPath) {
            return [figure.closedValue ? LengthQuantity.Perimeter : LengthQuantity.Length];
        }

        if (figure instanceof Bezier) {
            return [LengthQuantity.Length];
        }

        if (figure.isPerimeterProvider === true) {
            return [LengthQuantity.Perimeter];
        }

        if (figure.isLengthProvider === true) {
            return [LengthQuantity.Length];
        }

        return [];
    }

    get quantities() {
        return this.isBetweenPoints ? [LengthQuantity.Distance] : LengthMeasurement.quantitiesOf(this.measured);
    }

    get quantity() {
        if (this.quantityValue != null) {
            return this.quantityValue;
        }

        const quantities = this.quantities;
        return quantities.length > 0 ? quantities[0] : LengthQuantity.Length;
    }

    static defaultPrefix(quantity) {
        switch (quantity) {
            case LengthQuantity.Radius:
                return "r = ";
            case LengthQuantity.Diameter:
                return "d = ";
            case LengthQuantity.Circumference:
                return "C = ";
            case LengthQuantity.Perimeter:
                return "P = ";
            default:
                return "";
        }
    }

    get prefix() {
        return this.prefixValue ?? LengthMeasurement.defaultPrefix(this.quantity);
    }

    /** The quantity is one the figure has, and it has a value right now */
    get hasSomethingToMeasure() {
        return this.quantities.includes(this.quantity) && isValidValue(this.length);
    }

    updateExistence() {
        super.updateExistence();
        if (this.exists && !this.hasSomethingToMeasure) {
            this.exists = false;
        }
    }

    get length() {
        if (this.isBetweenPoints) {
            return this.point(0).distance(this.point(1));
        }

        const measured = this.measured;
        if (measured == null) {
            return NaN;
        }

        switch (this.quantity) {
            case LengthQuantity.Length:
                return measured.isLengthProvider === true || measured instanceof Bezier || measured instanceof BezierPath ? measured.length : NaN;
            case LengthQuantity.ArcLength:
                return measured.isArc === true ? measured.length : NaN;
            case LengthQuantity.Radius:
                return LengthMeasurement.radiusOf(measured);
            case LengthQuantity.Diameter:
                return 2 * LengthMeasurement.radiusOf(measured);
            case LengthQuantity.Circumference:
                return measured.isEllipse === true && measured.isArc !== true ? measured.perimeter : NaN;
            case LengthQuantity.Perimeter:
                return measured.isPerimeterProvider === true ? measured.perimeter : NaN;
            default:
                return NaN;
        }
    }

    static radiusOf(figure) {
        return figure.isCircle === true ? figure.radius : NaN;
    }

    /** A distance and a length sit at the middle, as they always did; the rest are placed outside */
    get placesItself() {
        const quantity = this.quantity;
        return quantity !== LengthQuantity.Distance && quantity !== LengthQuantity.Length;
    }

    /**
     * What the offset is from: the middle of the two points or of the figure for a distance or
     * a length; the middle of an arc for its length; the middle of a radius for a radius; the
     * center for a diameter; a point of the outline for a circumference or a perimeter. A circle
     * or an ellipse too big for the window: the point of it nearest the middle of the window.
     */
    get anchor() {
        if (this.isBetweenPoints) {
            return GeometryMath.midpoint(this.point(0), this.point(1));
        }

        const measured = this.measured;
        if (measured == null) {
            return Point.infinite;
        }

        switch (this.quantity) {
            case LengthQuantity.ArcLength:
                return measured.isArc === true ? measured.arcMiddle : measured.center;
            case LengthQuantity.Radius:
                return LengthMeasurement.radiusMiddle(measured);
            case LengthQuantity.Diameter:
                return measured.center;
            case LengthQuantity.Circumference:
                return measured.isEllipse === true && measured.isArc !== true ? this.onScreenPoint(measured) : measured.center;
            case LengthQuantity.Perimeter:
                return LengthMeasurement.outlinePoint(measured);
            case LengthQuantity.Length:
                if (measured instanceof Polyline || measured instanceof Bezier || measured instanceof BezierPath) {
                    // a curve's middle is on it, where its center may be far off
                    const domain = measured.getParameterDomain();
                    const middle = measured.getPointFromParameter((domain[0] + domain[1]) / 2);
                    return middle != null && middle.exists() ? middle : measured.center;
                }

                return measured.center;
            default:
                return measured.center;
        }
    }

    static radiusMiddle(measured) {
        if (measured instanceof Circle) {
            return GeometryMath.midpoint(measured.center, measured.point(1));
        }

        if (measured.isArc === true) {
            return GeometryMath.midpoint(measured.center, measured.beginLocation);
        }

        if (measured.isCircle === true) {
            const center = measured.center;
            return new Point(center.x + measured.radius / 2, center.y);
        }

        return measured.center;
    }

    static outlinePoint(measured) {
        if (measured.isPolygonalChain === true) {
            const vertices = measured.vertexCoordinates;
            return vertices != null && vertices.length >= 2
                ? GeometryMath.midpoint(vertices[0], vertices[1])
                : measured.center;
        }

        if (measured.isArc === true) {
            return measured.arcMiddle;
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
            && pixel.x >= -LengthMeasurement.EdgeInset && pixel.x <= canvas.x + LengthMeasurement.EdgeInset
            && pixel.y >= -LengthMeasurement.EdgeInset && pixel.y <= canvas.y + LengthMeasurement.EdgeInset;
    }

    updateVisual() {
        if (!this.hasSomethingToMeasure) {
            return;
        }

        this.setText(this.prefix + NumberFormat.toString(GeometryMath.round(this.length, this.decimalsToShow)));

        // where the number goes is decided once, at its first redraw, whatever its quantity: a
        // distance or a length keeps its offset (so a path closed later doesn't move it), the
        // rest get the default place; once the canvas has a size (before the first layout the
        // anchor of a circle could only be the fallback, and the offset worked out from it was wrong)
        if (!this.placed && this.drawing != null && this.drawing.coordinateSystem.physicalSize.x > 0) {
            this.placed = true;
            if (this.placesItself) {
                this.offset = this.defaultOffset();
            }
        }

        super.updateVisual();
    }

    /** Where a new number goes, in pixels from the anchor: clear of what it measures */
    defaultOffset() {
        const size = this.measureSize();
        const direction = this.outwardDirection();
        const distance = LengthMeasurement.Gap + Math.abs(direction.x) * size.width / 2 + Math.abs(direction.y) * size.height / 2;
        return direction.scale(distance).minus(new Point(size.width / 2, size.height / 2));
    }

    /** A unit vector in pixels pointing from the anchor to where the number goes */
    outwardDirection() {
        const measured = this.measured;
        const anchor = this.toPhysical(this.anchor);
        const center = this.toPhysical(measured.center);
        if (this.quantity === LengthQuantity.Diameter) {
            return new Point(0, 1);
        }

        if (this.quantity === LengthQuantity.Radius) {
            // across the radius, on its upper side
            const along = RightAngleMark.direction(center, anchor);
            if (along == null) {
                return new Point(0, -1);
            }

            const normal = new Point(along.y, -along.x);
            return normal.y <= 0 ? normal : new Point(-normal.x, -normal.y);
        }

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
        this.prefixValue = element.getAttribute("Prefix");
        const quantity = element.getAttribute("Quantity");
        this.quantityValue = quantity != null && Object.hasOwn(LengthQuantity, quantity) ? LengthQuantity[quantity] : null;
    }
}

FigureTypes.register("LengthMeasurement", LengthMeasurement);
