// Port of Main/Avalonia/DynamicGeometry/Figures/Lines/ParallelLine.cs

class ParallelLine extends LineTwoPoints {
    get coordinates() {
        const parentLine = this.line(0);
        const point = this.point(1);
        return new PointPair(point, point.plus(parentLine.p2.minus(parentLine.p1)));
    }
}

FigureTypes.register("ParallelLine", ParallelLine);
