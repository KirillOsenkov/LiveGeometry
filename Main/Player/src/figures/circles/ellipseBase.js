// Port of Main/Avalonia/DynamicGeometry/Figures/Circles/EllipseBase.cs

class EllipseBase extends ShapeBase {
    get isLinearFigure() {
        return true;
    }

    get isEllipse() {
        return true;
    }

    get semiMajor() {
        return 0;
    }

    get semiMinor() {
        return 0;
    }

    /** Angle of inclination in radians */
    get inclination() {
        return 0;
    }

    get area() {
        return Math.PI * this.semiMajor * this.semiMinor;
    }

    hitTest(point) {
        const width = this.logicalWidth();
        const fromEdge = GeometryMath.radialDistanceToEllipse(this.center, this.semiMajor, this.semiMinor, this.inclination, point);

        // the edge
        if (Math.abs(fromEdge) < this.cursorTolerance + width / 2) {
            return this;
        }

        // the fill
        const style = this.resolvedStyle;
        if (style != null && style.isFilled === true && fromEdge < 0) {
            return this;
        }

        return null;
    }

    logicalWidth() {
        return this.toLogicalLength(this.strokeThickness);
    }

    getNearestParameterFromPoint(point) {
        let result;
        if (Settings.pointsOnEllipticalsUseAbsoluteAngle) {
            result = GeometryMath.getAngle(this.center, point);
        } else {
            result = GeometryMath.getAngle(this.center, point) - this.inclination;
        }

        if (this.flipped) {
            result = -result;
        }

        return result;
    }

    getPointFromParameter(parameter) {
        const center = this.center;
        let angleToPoint = parameter;
        if (this.flipped) {
            angleToPoint = -angleToPoint;
        }

        if (!Settings.pointsOnEllipticalsUseAbsoluteAngle) {
            angleToPoint += this.inclination;
        }

        const intersections = GeometryMath.getIntersectionOfEllipseFigureAndLine(
            this,
            new PointPair(center, GeometryMath.getTranslationPoint(center, 1, angleToPoint)));
        const direction1 = GeometryMath.getAngle(center, intersections.p1);
        const cDiff = Math.cos(angleToPoint) - Math.cos(direction1);
        const sDiff = Math.sin(angleToPoint) - Math.sin(direction1);
        if (isWithinEpsilon(cDiff) && isWithinEpsilon(sDiff)) {
            return intersections.p1;
        }

        return intersections.p2;
    }

    getParameterDomain() {
        return [0, GeometryMath.DOUBLEPI];
    }

    updateVisual() {
    }

    render(renderer) {
        if (!this.isShown) {
            return;
        }

        const center = this.toPhysical(this.center);
        const major = this.toPhysicalLength(this.semiMajor);
        const minor = this.toPhysicalLength(this.semiMinor);
        const inclination = this.inclination;
        if (!center.exists() || !isValidValue(major) || !isValidValue(minor) || !isValidValue(inclination)) {
            return;
        }

        renderer.drawEllipse(center, major, minor, -inclination, this.stroke, this.fill);
    }
}
