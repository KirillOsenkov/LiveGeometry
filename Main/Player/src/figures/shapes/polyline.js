// Port of Main/Avalonia/DynamicGeometry/Figures/Shapes/Polyline.cs

class Polyline extends ShapeBase {
    constructor() {
        super();
        this.vertices = null;
        this.vertexCoordinates = null;
    }

    get isLinearFigure() {
        return true;
    }

    get isPolygonalChain() {
        return true;
    }

    get isLengthProvider() {
        return true;
    }

    get center() {
        return GeometryMath.midpoint(this.vertexCoordinates ?? []);
    }

    /** From the first point to the last, along the way */
    get length() {
        const points = toPoints(this.dependencies);
        let length = 0;
        for (let i = 1; i < points.length; i++) {
            length += points[i - 1].distance(points[i]);
        }

        return length;
    }

    onDependenciesChanged() {
        this.updatePointCache();
        if (this.drawing != null) {
            this.updateVisual();
        }
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

        renderer.drawPolygon(points, this.stroke, this.fill, false);
    }

    hitTest(point) {
        const epsilon = this.toLogicalLength(this.strokeThickness / 2 + GeometryMath.cursorTolerance);
        if (GeometryMath.isPointOnPolygonalChain(toPoints(this.dependencies), point, epsilon, false)) {
            return this;
        }

        return null;
    }

    getNearestParameterFromPoint(point) {
        return GeometryMath.getNearestParameterFromPointOnPolyline(toPoints(this.dependencies), point);
    }

    getPointFromParameter(parameter) {
        return GeometryMath.getPointOnPolylineFromParameter(toPoints(this.dependencies), parameter);
    }

    getParameterDomain() {
        return [0, 1];
    }
}

FigureTypes.register("Polyline", Polyline);
