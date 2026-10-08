// Port of Main/Avalonia/DynamicGeometry/Figures/Shapes/Curve.cs: a curve drawn through the
// points getPoints gives, with a Gap wherever it is interrupted.

class Curve extends ShapeBase {
    /** Among the points of a curve: the curve is interrupted here */
    static Gap = new Point(NaN, NaN);

    constructor() {
        super();

        // as getPoints gave them, in the units of the drawing, gaps included
        this.logicalPoints = [];
    }

    get isLinearFigure() {
        return true;
    }

    recalculate() {
    }

    updateVisual() {
        this.logicalPoints = [];
        this.getPoints(this.logicalPoints);
    }

    render(renderer) {
        if (!this.isShown) {
            return;
        }

        const stroke = this.stroke;
        const points = this.logicalPoints;
        let run = [];
        const flush = () => {
            if (run.length > 1) {
                renderer.drawPolyline(run, stroke);
            }

            run = [];
        };
        for (let i = 0; i < points.length; i++) {
            if (!points[i].exists()) {
                flush();
                continue;
            }

            const physical = this.toPhysical(points[i]);
            if (!physical.exists()) {
                flush();
                continue;
            }

            run.push(physical);
        }

        flush();
    }

    get center() {
        // the middle one of the points, or the nearest to it that is not a gap
        const points = this.logicalPoints;
        for (let offset = 0; offset < points.length; offset++) {
            const middle = Math.floor(points.length / 2);
            for (const index of [middle + offset, middle - offset]) {
                if (index >= 0 && index < points.length && points[index].exists()) {
                    return points[index];
                }
            }
        }

        return super.center;
    }

    /** The points of the curve, in the units of the drawing, with a Gap wherever it is interrupted */
    getPoints(points) {
    }

    hitTest(point) {
        const epsilon = this.toLogicalLength(this.strokeThickness / 2 + GeometryMath.cursorTolerance);
        const points = this.logicalPoints;
        for (let i = 1; i < points.length; i++) {
            const from = points[i - 1];
            const to = points[i];
            if (!from.exists() || !to.exists()) {
                continue;
            }

            if (GeometryMath.isPointOnSegment(new PointPair(from, to), point, epsilon)
                || from.distance(point) < epsilon
                || to.distance(point) < epsilon) {
                return this;
            }
        }

        return null;
    }

    getNearestParameterFromPoint(point) {
        return GeometryMath.getNearestParameterFromPointOnPolyline(this.logicalPoints, point);
    }

    getPointFromParameter(parameter) {
        return GeometryMath.getPointOnPolylineFromParameter(this.logicalPoints, parameter);
    }

    getParameterDomain() {
        return [0, 1];
    }
}

class CustomCurve extends Curve {
    constructor() {
        super();
        this.points = [];
    }

    getPoints(points) {
        points.push(...this.points);
    }
}
