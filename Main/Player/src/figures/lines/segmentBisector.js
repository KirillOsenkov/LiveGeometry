// Port of Main/Avalonia/DynamicGeometry/Figures/Lines/SegmentBisector.cs

class SegmentBisector extends PerpendicularLineBase {
    constructor() {
        super();
        this.coordinatesValue = new PointPair();
    }

    get coordinates() {
        return this.coordinatesValue;
    }

    recalculate() {
        const p1 = this.point(0);
        const p2 = this.point(1);
        const line = this.flipped ? new PointPair(p2, p1) : new PointPair(p1, p2);
        const midpoint = GeometryMath.midpoint(p1, p2);
        this.coordinatesValue = GeometryMath.getPerpendicularLine(line, midpoint);
    }

    tryGetRightAngle() {
        // in the middle of the two points, whether or not a segment is drawn between them
        const baseLine = new PointPair(this.point(0), this.point(1));
        const vertex = GeometryMath.midpoint(baseLine.p1, baseLine.p2);
        const along = baseLine.p2.minus(baseLine.p1);
        const pointAcross = vertex.plus(new Point(-along.y, along.x));
        return { vertex, baseLine, pointAcross, hasRightAngle: true };
    }
}

FigureTypes.register("SegmentBisector", SegmentBisector);
