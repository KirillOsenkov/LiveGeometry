// Port of Main/Avalonia/DynamicGeometry/Figures/Shapes/Bezier.cs: a cubic through four points

class Bezier extends ShapeBase {
    constructor() {
        super();
        this.info = null;
    }

    get isLinearFigure() {
        return true;
    }

    get isBezier() {
        return true;
    }

    /** The curve worked out, also before its first recalculate: a point on it from a file asks where it is while the file is read */
    get curveInfo() {
        if (this.info == null && this.dependencies.length === 4) {
            this.recalculate();
        }

        return this.info;
    }

    recalculate() {
        this.info = new BezierInfo(this.point(0), this.point(1), this.point(2), this.point(3));
    }

    updateVisual() {
    }

    render(renderer) {
        if (!this.isShown || this.dependencies.length < 4) {
            return;
        }

        const p = [0, 1, 2, 3].map(i => this.toPhysical(this.point(i)));
        if (p.some(point => !point.exists())) {
            return;
        }

        const commands = [
            { op: "move", x: p[0].x, y: p[0].y },
            { op: "cubic", x1: p[1].x, y1: p[1].y, x2: p[2].x, y2: p[2].y, x: p[3].x, y: p[3].y }
        ];
        const box = CanvasRenderer.boxOf(p);
        renderer.drawPath(commands, this.stroke, this.fill, box);
    }

    get center() {
        return GeometryMath.midpoint(this.point(0), this.point(1), this.point(2), this.point(3));
    }

    hitTest(point) {
        const curve = this.curveInfo;
        if (curve == null || curve.points == null) {
            return null;
        }

        if (GeometryMath.isPointOnPolygonalChain(curve.points, point, this.toLogicalLength(this.strokeThickness / 2 + GeometryMath.cursorTolerance), false)) {
            return this;
        }

        // the fill: the curve closed by a straight line
        const style = this.resolvedStyle;
        if (style != null && style.isFilled === true && GeometryMath.isPointInPolygon(curve.points, point)) {
            return this;
        }

        return null;
    }

    getNearestParameterFromPoint(point) {
        const curve = this.curveInfo;
        return curve == null ? 0 : curve.getNearestParameterFromPoint(point);
    }

    getPointFromParameter(parameter) {
        return GeometryMath.getPointOnPolylineFromParameter(this.curveInfo?.points ?? null, parameter);
    }

    getParameterDomain() {
        return [0, 1];
    }
}

FigureTypes.register("Bezier", Bezier);
