// Port of Main/Avalonia/DynamicGeometry/Figures/Circles/Arc.cs: the arcs, segments and sectors
// of circles and ellipses. Left out: the Convert verbs.

class EllipseArc extends EllipseArcBase {
}

class CircleArc extends CircleArcBase {
}

/** The chord is drawn, not a figure of its own; unlike an arc, a segment has an area */
class CircleSegment extends CircleArcBase {
    get isShapeWithInterior() {
        return true;
    }

    get isSegmentShape() {
        return true;
    }

    get area() {
        return this.segmentArea;
    }
}

class EllipseSegment extends EllipseArcBase {
    get isShapeWithInterior() {
        return true;
    }

    get isSegmentShape() {
        return true;
    }

    get area() {
        return this.segmentArea;
    }
}

class CircleSector extends CircleArcBase {
    get isShapeWithInterior() {
        return true;
    }

    get isSectorShape() {
        return true;
    }

    get area() {
        return this.sectorArea;
    }
}

class EllipseSector extends EllipseArcBase {
    get isShapeWithInterior() {
        return true;
    }

    get isSectorShape() {
        return true;
    }

    get area() {
        return this.sectorArea;
    }
}

FigureTypes.register("EllipseArc", EllipseArc);
FigureTypes.register("CircleArc", CircleArc);
FigureTypes.register("CircleSegment", CircleSegment);
FigureTypes.register("EllipseSegment", EllipseSegment);
FigureTypes.register("CircleSector", CircleSector);
FigureTypes.register("EllipseSector", EllipseSector);
