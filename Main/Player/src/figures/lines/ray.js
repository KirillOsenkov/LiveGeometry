// Port of Main/Avalonia/DynamicGeometry/Figures/Lines/Ray.cs

class Ray extends LineBase {
    get isLine() {
        return true;
    }

    get onScreenCoordinates() {
        const c = GeometryMath.getLineFromSegment(this.coordinates, this.canvasLogicalBorders);
        return new PointPair(this.coordinates.p1, c.p2);
    }

    getNearestParameterFromPoint(point) {
        let parameter = super.getNearestParameterFromPoint(point);
        if (parameter < 0) {
            parameter = 0;
        }

        return parameter;
    }

    hitTest(point) {
        const hit = super.hitTest(point) != null;
        const line = this.coordinates;

        // the start is on the ray however the rounding of a point there went
        const inside = GeometryMath.getProjection(point, line).ratio >= -GeometryMath.endTolerance(line);
        if (hit && inside) {
            return this;
        }

        return null;
    }

    getParameterDomain() {
        return [0, super.getParameterDomain()[1]];
    }
}

FigureTypes.register("Ray", Ray);
