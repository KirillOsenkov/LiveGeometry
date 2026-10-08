// Port of Main/Avalonia/DynamicGeometry/Figures/Points/ReflectedPoint.cs

class ReflectedPoint extends PointBase {
    get source() {
        return this.dependencies[0];
    }

    get mirror() {
        return this.dependencies[1];
    }

    onDependenciesChanged() {
        const mirror = this.dependencies.length > 1 ? this.mirror : null;
        this.mirrorPoint = mirror != null && mirror.isPoint === true ? mirror : null;
        this.mirrorLine = mirror != null && mirror.isLine === true ? mirror : null;
        this.mirrorCircle = mirror != null && mirror.isCircle === true ? mirror : null;
    }

    recalculate() {
        const source = this.point(0);
        if (this.mirrorPoint != null) {
            this.coordinates = GeometryMath.getSymmetricPointThroughPoint(source, this.mirrorPoint.coordinates);
        } else if (this.mirrorLine != null) {
            this.coordinates = GeometryMath.getSymmetricPointAcrossLine(source, this.mirrorLine.coordinates);
        } else if (this.mirrorCircle != null) {
            this.coordinates = GeometryMath.getSymmetricPointInCircle(source, this.mirrorCircle.center, this.mirrorCircle.radius);
        }

        // the image of a point that is not there is not there either
        this.exists = allExist(this.dependencies) && this.coordinates.exists();
    }
}

FigureTypes.register("ReflectedPoint", ReflectedPoint);
