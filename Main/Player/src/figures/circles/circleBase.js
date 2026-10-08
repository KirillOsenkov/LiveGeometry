// Port of Main/Avalonia/DynamicGeometry/Figures/Circles/CircleBase.cs. Left out: setting and fixing the radius.

class CircleBase extends EllipseBase {
    get isCircle() {
        return true;
    }

    get radius() {
        return 0;
    }

    /** The radius, as the grid's Length row and an expression's AB.Length read it */
    get length() {
        return this.radius;
    }

    get inclination() {
        return 0;
    }

    get semiMajor() {
        return this.radius;
    }

    get semiMinor() {
        return this.radius;
    }

    getPointFromParameter(parameter) {
        if (Settings.pointsOnEllipticalsUseAbsoluteAngle) {
            const center = this.center;
            const radius = this.radius;
            return new Point(center.x + radius * Math.cos(parameter), center.y + radius * Math.sin(parameter));
        }

        return super.getPointFromParameter(parameter);
    }

    render(renderer) {
        if (!this.isShown) {
            return;
        }

        const center = this.toPhysical(this.center);
        const radius = this.toPhysicalLength(this.radius);
        if (!center.exists() || !isValidValue(radius)) {
            return;
        }

        renderer.drawEllipse(center, radius, radius, 0, this.stroke, this.fill);
    }
}
