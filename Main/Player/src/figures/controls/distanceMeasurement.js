// Port of Main/Avalonia/DynamicGeometry/Figures/Controls/DistanceMeasurement.cs

class DistanceMeasurement extends Measurement {
    constructor() {
        super();
        this.units = Settings.distanceUnit;
    }

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

    get measure() {
        return this.distance * this.conversionFactor;
    }

    get length() {
        return this.distance;
    }

    updateVisual() {
        if (!this.hasSomethingToMeasure) {
            return;
        }

        super.updateVisual();
        const distance = NumberFormat.toString(GeometryMath.round(this.measure, this.decimalsToShow));
        if (this.units === LengthUnit.Inches) {
            this.setText(distance + "\"");
        } else if (this.units === LengthUnit.Centimeter) {
            this.setText(distance + "cm");
        } else {
            this.setText(distance);
        }
    }

    get conversionFactor() {
        if (this.units === LengthUnit.Inches) {
            return 1 / GeometryMath.inchesLogicalLength;
        }

        if (this.units === LengthUnit.Centimeter) {
            return 1 / GeometryMath.centimeterLogicalLength;
        }

        return 1;
    }

    readXml(element) {
        super.readXml(element);
        const unitsAsString = element.getAttribute("Units");
        if (unitsAsString === "Inches") {
            this.units = LengthUnit.Inches;
        } else if (unitsAsString === "Centimeters") {
            this.units = LengthUnit.Centimeter;
        }
    }
}

FigureTypes.register("DistanceMeasurement", DistanceMeasurement);
