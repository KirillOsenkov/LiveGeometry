// Port of Main/Avalonia/DynamicGeometry/Figures/Lines/PerpendicularLine.cs

class PerpendicularLine extends PerpendicularLineBase {
    get coordinates() {
        let line = this.line(0);
        if (this.flipped) {
            line = new PointPair(line.p2, line.p1);
        }

        return GeometryMath.getPerpendicularLine(line, this.point(1));
    }

    tryGetRightAngle() {
        // where this line crosses the one it is perpendicular to - if it does: the foot can be beyond the end of a segment
        const baseLine = this.line(0);
        const pointAcross = this.point(1);
        const vertex = GeometryMath.getProjectionPoint(pointAcross, baseLine);
        const hasRightAngle = this.dependencies[0].visible && this.baseFigure.hitTest(vertex) != null;
        return { vertex, baseLine, pointAcross, hasRightAngle };
    }

    /** A vector by the segment inside it: the mark is about where the vector ends, not its head */
    get baseFigure() {
        const base = this.dependencies[0];
        return base.isVector === true ? base.line : base;
    }
}

FigureTypes.register("PerpendicularLine", PerpendicularLine);
