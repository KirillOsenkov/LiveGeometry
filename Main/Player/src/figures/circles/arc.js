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

    get isPerimeterProvider() {
        return true;
    }

    get isSegmentShape() {
        return true;
    }

    get area() {
        return this.segmentArea;
    }

    /** The arc and its chord */
    get perimeter() {
        return this.length + this.beginLocation.distance(this.endLocation);
    }
}

class EllipseSegment extends EllipseArcBase {
    get isShapeWithInterior() {
        return true;
    }

    get isPerimeterProvider() {
        return true;
    }

    get isSegmentShape() {
        return true;
    }

    get area() {
        return this.segmentArea;
    }

    /** The arc and its chord */
    get perimeter() {
        return this.length + this.beginLocation.distance(this.endLocation);
    }
}

class CircleSector extends CircleArcBase {
    get isShapeWithInterior() {
        return true;
    }

    get isPerimeterProvider() {
        return true;
    }

    get isSectorShape() {
        return true;
    }

    get area() {
        return this.sectorArea;
    }

    /** The arc and the two radii */
    get perimeter() {
        return this.length + this.center.distance(this.beginLocation) + this.center.distance(this.endLocation);
    }
}

class EllipseSector extends EllipseArcBase {
    get isShapeWithInterior() {
        return true;
    }

    get isPerimeterProvider() {
        return true;
    }

    get isSectorShape() {
        return true;
    }

    get area() {
        return this.sectorArea;
    }

    /** The arc and the two radii */
    get perimeter() {
        return this.length + this.center.distance(this.beginLocation) + this.center.distance(this.endLocation);
    }
}

FigureTypes.register("EllipseArc", EllipseArc);
FigureTypes.register("CircleArc", CircleArc);
FigureTypes.register("CircleSegment", CircleSegment);
FigureTypes.register("EllipseSegment", EllipseSegment);
FigureTypes.register("CircleSector", CircleSector);
FigureTypes.register("EllipseSector", EllipseSector);
