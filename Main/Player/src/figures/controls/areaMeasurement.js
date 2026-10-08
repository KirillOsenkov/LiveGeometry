// Port of Main/Avalonia/DynamicGeometry/Figures/Controls/AreaMeasurement.cs

class AreaMeasurement extends Measurement {
    constructor() {
        super();
        this.units = Settings.distanceUnit;
    }

    get conversionFactor() {
        if (this.units === LengthUnit.Inches) {
            return 1 / GeometryMath.sqr(GeometryMath.inchesLogicalLength);
        }

        if (this.units === LengthUnit.Centimeter) {
            return 1 / GeometryMath.sqr(GeometryMath.centimeterLogicalLength);
        }

        return 1;
    }

    get measure() {
        if (this.dependencies[0].isShapeWithInterior === true) {
            return this.dependencies[0].area * this.conversionFactor;
        }

        return GeometryMath.area(toPoints(this.dependencies)) * this.conversionFactor;
    }

    get anchor() {
        return this.origin;
    }

    updateVisual() {
        super.updateVisual();
        const areaText = NumberFormat.toString(GeometryMath.round(this.measure, this.decimalsToShow));
        if (this.units === LengthUnit.Inches) {
            this.setText(areaText + "in²");
        } else if (this.units === LengthUnit.Centimeter) {
            this.setText(areaText + "cm²");
        } else {
            this.setText(areaText);
        }
    }

    get origin() {
        if (this.dependencies[0] instanceof PointBase) {
            return GeometryMath.midpoint(toPoints(this.dependencies));
        }

        return this.dependencies[0].center;
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

FigureTypes.register("AreaMeasurement", AreaMeasurement);
