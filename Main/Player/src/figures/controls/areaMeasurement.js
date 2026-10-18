// Port of Main/Avalonia/LiveGeometry/Figures/Controls/AreaMeasurement.cs

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

    /** Not while there is no area: a path opened and unfilled after it was measured */
    updateExistence() {
        super.updateExistence();
        if (this.exists && this.dependencies.length > 0 && !isValidValue(this.measure)) {
            this.exists = false;
        }
    }

    updateVisual() {
        if (!this.hasAnchor) {
            return;
        }

        this.setText(NumberFormat.toString(GeometryMath.round(this.measure, this.decimalsToShow)));
        super.updateVisual();
    }

    /** A new one: the number in the middle of the region */
    defaultOffset() {
        return FigureAnchor.centered(this.measureSize());
    }

    /** The middle of the region: a polygon's centroid by area, the center of a circle, an ellipse, a sector or a path */
    get origin() {
        if (this.dependencies[0] instanceof PointBase) {
            return FigureAnchor.centroid(toPoints(this.dependencies));
        }

        if (this.dependencies[0] instanceof PolygonBase) {
            return FigureAnchor.centroid(this.dependencies[0].vertexCoordinates);
        }

        return this.dependencies[0].center;
    }
}

FigureTypes.register("AreaMeasurement", AreaMeasurement);
