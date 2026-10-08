// Port of Main/Avalonia/DynamicGeometry/Figures/Lines/Segment.cs. Left out: setting the
// length and fixing it (the grid's), the decoration mark (SegmentDecoration comes later;
// the attribute is read and kept).

class Segment extends LineBase {
    constructor() {
        super();

        /** The mark at the middle: ticks, chevrons or a wave (SegmentDecoration), by name; None for none */
        this.decoration = "None";
    }

    get isLine() {
        return true;
    }

    get isLengthProvider() {
        return true;
    }

    readXml(element) {
        super.readXml(element);
        const name = element.getAttribute("Decoration");
        if (name != null) {
            this.decoration = name;
        }
    }

    render(renderer) {
        super.render(renderer);
        if (this.decoration !== "None" && this.isShown && this.screenLine != null && typeof SegmentDecorationMark !== "undefined") {
            SegmentDecorationMark.render(renderer, this.screenLine.p1, this.screenLine.p2, this.stroke, this.decoration);
        }
    }

    get length() {
        return this.coordinates.length;
    }

    getNearestParameterFromPoint(point) {
        let parameter = super.getNearestParameterFromPoint(point);
        if (parameter < 0) {
            parameter = 0;
        } else if (parameter > 1) {
            parameter = 1;
        }

        return parameter;
    }

    hitTest(point) {
        const epsilon = this.toLogicalLength(this.strokeThickness) / 2 + this.cursorTolerance;
        if (GeometryMath.isPointOnSegment(this.coordinates, point, epsilon)) {
            return this;
        }

        return null;
    }

    getParameterDomain() {
        return [0, 1];
    }
}

FigureTypes.register("Segment", Segment);
