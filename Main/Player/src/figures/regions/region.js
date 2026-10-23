// Port of Main/Avalonia/LiveGeometry/Figures/Regions/Region.cs: a set made of other figures
// with an inside (isRegion: polygons, circles, ellipses, sectors, circular segments, Bezier
// paths, other regions) by an operation: the first without the others, what they all share,
// or all of them together. The inside alone, filled; its boundary is traced from the
// operands' (RegionBoundary) for the fill, the area, the perimeter and a point kept inside.
// Left out: the tools, the Operation row.

class Region extends ShapeBase {
    constructor() {
        super();
        this.operation = RegionOperation.Difference;
        this.boundary = RegionBoundary.empty;
    }

    /** What the tracer asks a nested region for (isRegion is every figure with an inside) */
    get isRegionFigure() {
        return true;
    }

    get isShapeWithInterior() {
        return true;
    }

    get isPerimeterProvider() {
        return true;
    }

    get isRegion() {
        return true;
    }

    /** The figures the region is made of, in order; null where a dependency has no inside */
    get operands() {
        return this.dependencies.map(dependency => dependency.isRegion === true ? dependency : null);
    }

    defaultLayer() {
        return ZOrder.Polygons;
    }

    readXml(element) {
        super.readXml(element);
        const operation = element.getAttribute("Operation");
        this.operation = Object.values(RegionOperation).includes(operation) ? operation : RegionOperation.Difference;
    }

    /** Not while a dependency is no figure with an inside */
    updateExistence() {
        super.updateExistence();
        if (this.exists && (this.dependencies.length === 0 || this.dependencies.some(dependency => dependency.isRegion !== true))) {
            this.exists = false;
        }
    }

    recalculate() {
        this.boundary = this.exists ? RegionBoundary.trace(this.operation, this.operands) : RegionBoundary.empty;
    }

    get area() {
        return this.boundary.area;
    }

    get perimeter() {
        return this.boundary.perimeter;
    }

    get hasInside() {
        return this.exists && this.boundary.hasInside;
    }

    containsPoint(point) {
        return this.exists && RegionBoundary.contains(this.operation, this.operands, point);
    }

    nearestPointInside(point) {
        if (!this.hasInside) {
            return Point.infinite;
        }

        return this.containsPoint(point) ? point : this.boundary.nearestPoint(point);
    }

    /** The centroid of the set; the first operand's center while there is none */
    get center() {
        const centroid = this.boundary.centroid;
        if (centroid.exists()) {
            return centroid;
        }

        return this.dependencies.length > 0 ? this.dependencies[0].center : new Point(0, 0);
    }

    /** Inside where it is filled, and on its edge where that is drawn */
    hitTest(point) {
        if (!this.visible || !this.exists || !this.boundary.hasInside) {
            return null;
        }

        const fill = this.fill;
        if (fill != null && this.containsPoint(point)) {
            return this;
        }

        const stroke = this.stroke;
        if (CanvasRenderer.hasStroke(stroke) && this.boundary.distanceTo(point) <= this.cursorTolerance + this.toLogicalLength(stroke.width) / 2) {
            return this;
        }

        return null;
    }

    render(renderer) {
        if (!this.isShown || this.boundary.loops.length === 0) {
            return;
        }

        const path = this.boundary.buildPath(this.drawing.coordinateSystem);
        renderer.drawPath(path.commands, this.stroke, this.fill, path.box);
    }
}

FigureTypes.register("Region", Region);
