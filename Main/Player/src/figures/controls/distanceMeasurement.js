// Port of Main/Avalonia/DynamicGeometry/Figures/Controls/DistanceMeasurement.cs

class DistanceMeasurement extends Measurement {
    get isLengthProvider() {
        return true;
    }

    get anchor() {
        return this.midpoint();
    }

    /** A figure with a length, or two points; anything else has nothing to measure */
    get hasSomethingToMeasure() {
        return this.dependencies.length > 0
            && (this.dependencies[0].isLengthProvider === true
                || (this.dependencies.length > 1 && this.dependencies[0].isPoint === true && this.dependencies[1].isPoint === true));
    }

    updateExistence() {
        super.updateExistence();
        if (this.exists && !this.hasSomethingToMeasure) {
            this.exists = false;
        }
    }

    midpoint() {
        if (!this.hasSomethingToMeasure) {
            return Point.infinite;
        }

        if (this.dependencies[0].isLengthProvider === true) {
            return this.dependencies[0].center;
        }

        return GeometryMath.midpoint(this.point(0), this.point(1));
    }

    get distance() {
        if (!this.hasSomethingToMeasure) {
            return NaN;
        }

        if (this.dependencies[0].isLengthProvider === true) {
            return this.dependencies[0].length;
        }

        return this.point(0).distance(this.point(1));
    }

    get length() {
        return this.distance;
    }

    updateVisual() {
        if (!this.hasSomethingToMeasure) {
            return;
        }

        super.updateVisual();
        this.setText(NumberFormat.toString(GeometryMath.round(this.distance, this.decimalsToShow)));
    }
}

FigureTypes.register("DistanceMeasurement", DistanceMeasurement);
