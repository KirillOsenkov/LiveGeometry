// Port of Main/Avalonia/DynamicGeometry/Figures/Controls/AreaMeasurement.cs

class AreaMeasurement extends Measurement {
    get measure() {
        if (this.dependencies[0].isShapeWithInterior === true) {
            return this.dependencies[0].area;
        }

        return GeometryMath.area(toPoints(this.dependencies));
    }

    get anchor() {
        return this.origin;
    }

    updateVisual() {
        super.updateVisual();
        this.setText(NumberFormat.toString(GeometryMath.round(this.measure, this.decimalsToShow)));
    }

    get origin() {
        if (this.dependencies[0] instanceof PointBase) {
            return GeometryMath.midpoint(toPoints(this.dependencies));
        }

        return this.dependencies[0].center;
    }
}

FigureTypes.register("AreaMeasurement", AreaMeasurement);
