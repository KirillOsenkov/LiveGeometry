// Port of Main/Avalonia/DynamicGeometry/Figures/Shapes/PolygonBase.cs

class PolygonBase extends ShapeBase {
    constructor() {
        super();
        this.vertices = null;
        this.vertexCoordinates = null;
    }

    get isPolygonalChain() {
        return true;
    }

    get isShapeWithInterior() {
        return true;
    }

    onDependenciesChanged() {
        this.updatePointCache();
    }

    updatePointCache() {
        this.vertices = this.dependencies.filter(f => f.isPoint === true);
        this.vertexCoordinates = this.vertices.map(() => new Point());
    }

    updateVisual() {
        if (this.vertices == null) {
            this.updatePointCache();
        }

        for (let i = 0; i < this.vertices.length; i++) {
            this.vertexCoordinates[i] = this.vertices[i].coordinates;
        }
    }

    render(renderer) {
        if (!this.isShown || this.vertexCoordinates == null) {
            return;
        }

        const points = this.vertexCoordinates.map(p => this.toPhysical(p));
        if (points.some(p => !p.exists())) {
            return;
        }

        renderer.drawPolygon(points, this.stroke, this.fill, true);
    }

    defaultLayer() {
        return ZOrder.Polygons;
    }

    hitTest(point) {
        const isInside = this.vertexCoordinates != null && GeometryMath.isPointInPolygon(this.vertexCoordinates, point);
        return isInside ? this : null;
    }

    get area() {
        return GeometryMath.area(this.vertexCoordinates);
    }

    get isPerimeterProvider() {
        return true;
    }

    get perimeter() {
        return GeometryMath.distanceAround(this.vertexCoordinates);
    }

    get center() {
        return GeometryMath.midpoint(this.vertexCoordinates);
    }
}
