// Port of Main/Avalonia/DynamicGeometry/Figures/Circles/CircleByRadius.cs: the radius is
// either the distance between two points or the length of a figure; the center is the last
// dependency either way.

class CircleByRadius extends CircleBase {
    get isShapeWithInterior() {
        return true;
    }

    get center() {
        return this.point(this.dependencies.length - 1);
    }

    get radius() {
        const first = this.dependencies[0];
        if (first.isLengthProvider === true) {
            return first.length;
        }

        return this.point(0).distance(this.point(1));
    }

    /** No circle of a radius that is no length: a label or a Number that says a negative number or "undefined" */
    updateExistence() {
        super.updateExistence();
        if (this.exists && !isValidNonNegativeValue(this.radius)) {
            this.exists = false;
        }
    }
}

FigureTypes.register("CircleByRadius", CircleByRadius);
