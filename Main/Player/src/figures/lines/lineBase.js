// Port of Main/Avalonia/DynamicGeometry/Figures/Lines/LineBase.cs

class LineBase extends ShapeBase {
    constructor() {
        super();

        /** The part drawn on screen, in pixels, as updateVisual last worked it out */
        this.screenLine = null;
    }

    get isLinearFigure() {
        return true;
    }

    /** The part of the line that is drawn: all of it for a segment, clipped to the canvas for a line or a ray */
    get onScreenCoordinates() {
        return this.coordinates;
    }

    updateVisual() {
        if (this.isShown) {
            this.screenLine = this.toPhysicalPair(this.onScreenCoordinates);
        } else {
            this.screenLine = null;
        }
    }

    render(renderer) {
        if (!this.isShown || this.screenLine == null) {
            return;
        }

        const stroke = this.stroke;
        if (stroke == null) {
            return;
        }

        renderer.drawLine(this.screenLine.p1, this.screenLine.p2, stroke);
    }

    get coordinates() {
        return new PointPair(this.point(0), this.point(1));
    }

    get center() {
        return this.coordinates.midpoint;
    }

    hitTest(point) {
        const epsilon = this.toLogicalLength(this.strokeThickness) / 2 + this.cursorTolerance;
        if (GeometryMath.isPointOnLine(this.coordinates, point, epsilon)) {
            return this;
        }

        return null;
    }

    getNearestParameterFromPoint(point) {
        return GeometryMath.getProjection(point, this.coordinates).ratio;
    }

    getPointFromParameter(parameter) {
        return GeometryMath.scalePointBetweenTwo(this.coordinates, parameter);
    }

    /** [from, to] of the parameter */
    getParameterDomain() {
        const coordinates = this.onScreenCoordinates;
        const p1 = this.getNearestParameterFromPoint(coordinates.p1);
        const p2 = this.getNearestParameterFromPoint(coordinates.p2);
        return [p1 * 2, p2 * 2];
    }

    /** The direction from the first point to the second, in degrees */
    get angle() {
        return GeometryMath.toDegrees(GeometryMath.getAngle(this.coordinates.p1, this.coordinates.p2));
    }
}
