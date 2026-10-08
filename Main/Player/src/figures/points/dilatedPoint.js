// Port of Main/Avalonia/DynamicGeometry/Figures/Points/DilatedPoint.cs: a point stretched
// from a center by a factor, a figure it depends on: a Number, or anything with a length -
// and a fourth dependency when the factor is the ratio of two lengths.

class DilatedPoint extends PointBase {
    get source() {
        return this.dependencies.length >= 1 && this.dependencies[0].isPoint === true ? this.dependencies[0] : null;
    }

    get dilationCenter() {
        return this.dependencies.length >= 2 && this.dependencies[1].isPoint === true ? this.dependencies[1] : null;
    }

    /** A Number or a length provider */
    get factorSource() {
        return this.dependencies.length >= 3 ? this.dependencies[2] : null;
    }

    /** The factor is the ratio of two lengths */
    get isRatio() {
        return this.dependencies.length >= 4;
    }

    get factor() {
        if (this.isRatio) {
            const denominator = this.dependencies[3].isLengthProvider === true ? this.dependencies[3].length : 1;
            // over a length of 0 there is no factor, and no point
            return denominator !== 0 ? (this.dependencies[2].isLengthProvider === true ? this.dependencies[2].length : 1) / denominator : NaN;
        }

        const source = this.factorSource;
        if (source instanceof NumberFigure) {
            return source.value;
        }

        if (source != null && source.isLengthProvider === true) {
            return source.length;
        }

        return 1;
    }

    recalculate() {
        const source = this.source;
        const center = this.dilationCenter;
        if (source != null && center != null) {
            this.coordinates = GeometryMath.getDilationPoint(source.coordinates, center.coordinates, this.factor);
        }

        this.exists = allExist(this.dependencies) && this.coordinates.exists();
    }
}

FigureTypes.register("DilatedPoint", DilatedPoint);
