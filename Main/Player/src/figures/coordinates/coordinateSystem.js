// Port of Main/Avalonia/DynamicGeometry/Figures/Coordinates/CoordinateSystem.cs: the view,
// an origin in pixels plus the number of pixels per unit; every zoom and fit goes through it.

class CoordinateSystem {
    static MinUnitLength = 1;
    static MaxUnitLength = 1000;
    static zoomFactor = 1.2;

    /** What "zoom to fit" leaves free around the drawing */
    static FitMarginPixels = 40;

    /** "Zoom to fit" on a tiny drawing stops here */
    static MaxFitUnitLength = 200;

    /** Labeled grid lines come no closer than this, in pixels */
    static MinimumMajorGridSpacing = 40;

    /** The finer lines between the labeled ones are left out when they would be closer than this */
    static MinimumMinorGridSpacing = 10;

    constructor(drawing) {
        this.drawing = drawing;
        this.unitLength = Settings.defaultUnitLength;
        this.scale = 1;
        this.logicalViewportVertices = null;
        this.minimalVisibleX = 0;
        this.minimalVisibleY = 0;
        this.maximalVisibleX = 0;
        this.maximalVisibleY = 0;

        /** The grid never goes finer than this, in units; 0 lets the zoom decide (Viewport GridStep) */
        this.gridStep = 0;
        this.origin = this.physicalSize.scale(0.5).snapToIntegers().minus(0.5);
    }

    get canvas() {
        return this.drawing.canvas;
    }

    get physicalSize() {
        const canvas = this.canvas;
        return canvas == null ? new Point(0, 0) : new Point(canvas.width, canvas.height);
    }

    /** The drawing's canvas changed its size: what was in the middle stays in the middle */
    onSizeChanged(previousWidth, previousHeight, newWidth, newHeight) {
        if (previousWidth > 0 && previousHeight > 0) {
            this.origin = new Point(
                this.origin.x + (newWidth - previousWidth) / 2,
                this.origin.y + (newHeight - previousHeight) / 2);
        }

        this.recalculate();
    }

    /**
     * Sets the unit length and origin so that the rectangle shows. A rectangle or a canvas
     * without a size leaves the view as it is.
     */
    setViewport(minX, maxX, minY, maxY) {
        const logicalWidth = maxX - minX;
        const logicalHeight = maxY - minY;
        const physicalSize = this.physicalSize;
        if (!isValidPositiveValue(logicalWidth)
            || !isValidPositiveValue(logicalHeight)
            || !isValidPositiveValue(physicalSize.x)
            || !isValidPositiveValue(physicalSize.y)) {
            return;
        }

        this.fit(new Rect(minX, minY, logicalWidth, logicalHeight), 0);
    }

    /** The zoom within its limits; one that is no number is the default one */
    static clampUnitLength(value) {
        if (Number.isNaN(value)) {
            return Settings.defaultUnitLength;
        }

        return Math.max(CoordinateSystem.MinUnitLength, Math.min(CoordinateSystem.MaxUnitLength, value));
    }

    /** Shows the logical rectangle as large as the canvas allows, centered */
    fit(logicalBounds, marginPixels, maxUnitLength = CoordinateSystem.MaxUnitLength, zoom = 1) {
        const physicalSize = this.physicalSize;
        const availableWidth = Math.max(physicalSize.x - 2 * marginPixels, physicalSize.x / 2);
        const availableHeight = Math.max(physicalSize.y - 2 * marginPixels, physicalSize.y / 2);

        // a single point has no size: nothing to derive the zoom from, keep it
        let newUnitLength = this.unitLength;
        if (logicalBounds.width > 0 || logicalBounds.height > 0) {
            newUnitLength = Math.min(
                logicalBounds.width > 0 ? availableWidth / logicalBounds.width : Number.MAX_VALUE,
                logicalBounds.height > 0 ? availableHeight / logicalBounds.height : Number.MAX_VALUE);
            newUnitLength = Math.min(newUnitLength, maxUnitLength) * zoom;
        }

        this.setView(logicalBounds.center, CoordinateSystem.clampUnitLength(newUnitLength));
    }

    /** A suggested view of the drawing, edge to edge (Drawing.scenes) */
    fitScene(scene) {
        this.fit(scene, 0);
    }

    /** Puts the logical point under the given pixel (the middle of the canvas by default) at the given zoom */
    setView(logicalPoint, newUnitLength, physicalPoint = this.physicalSize.scale(0.5)) {
        this.unitLength = newUnitLength;
        this.scale = this.unitLength / Settings.defaultUnitLength;
        this.origin = CoordinateSystem.snapOrigin(new Point(
            physicalPoint.x - logicalPoint.x * this.unitLength,
            physicalPoint.y + logicalPoint.y * this.unitLength));
        this.recalculate();
    }

    /** The origin sits in the middle of a pixel so that the axes are crisp */
    static snapOrigin(origin) {
        return origin.snapToIntegers().minus(0.5);
    }

    get viewCenter() {
        return this.toLogical(this.physicalSize.scale(0.5));
    }

    zoomIn() {
        this.zoom(CoordinateSystem.zoomFactor, this.physicalSize.scale(0.5));
    }

    zoomOut() {
        this.zoom(1 / CoordinateSystem.zoomFactor, this.physicalSize.scale(0.5));
    }

    /** Zooms so that whatever is under the focus (physical) stays under it */
    zoom(factor, focus) {
        const newUnitLength = CoordinateSystem.clampUnitLength(this.unitLength * factor);
        if (newUnitLength === this.unitLength) {
            return;
        }

        const ratio = newUnitLength / this.unitLength;
        this.unitLength = newUnitLength;
        this.scale = this.unitLength / Settings.defaultUnitLength;

        // not snapped: a rounded origin makes the point under the cursor creep while zooming
        this.origin = new Point(
            focus.x - (focus.x - this.origin.x) * ratio,
            focus.y - (focus.y - this.origin.y) * ratio);
        this.recalculate();
        this.drawing.viewChanged?.();
    }

    /** Two fingers on the paper: what was under `from` goes under `to`, zoomed about it by the factor */
    panAndZoom(from, to, factor) {
        const newUnitLength = CoordinateSystem.clampUnitLength(this.unitLength * factor);
        const ratio = newUnitLength / this.unitLength;
        if (ratio === 1 && from.equals(to)) {
            return;
        }

        this.unitLength = newUnitLength;
        this.scale = this.unitLength / Settings.defaultUnitLength;
        this.origin = new Point(
            to.x - (from.x - this.origin.x) * ratio,
            to.y - (from.y - this.origin.y) * ratio);
        this.recalculate();
        this.drawing.viewChanged?.();
    }

    /** Zoom to fit: everything visible, as large as possible. An empty drawing goes back to the default view. */
    zoomExtend(alsoShow = null, zoom = 1) {
        let bounds = this.tryGetBoundsToShow(alsoShow);
        if (bounds == null) {
            this.setView(new Point(), Settings.defaultUnitLength);
            return;
        }

        const margin = CoordinateSystem.FitMarginPixels + this.getPointReach();
        this.fit(bounds, margin, CoordinateSystem.MaxFitUnitLength, zoom);

        // text keeps its size in pixels, so in logical units a label grows as the view zooms
        // out: measure again at the new zoom until it settles
        for (let i = 0; i < 3; i++) {
            const refitted = this.tryGetBoundsToShow(alsoShow);
            if (refitted == null || refitted.equals(bounds)) {
                break;
            }

            bounds = refitted;
            this.fit(bounds, margin, CoordinateSystem.MaxFitUnitLength, zoom);
        }
    }

    tryGetBoundsToShow(alsoShow) {
        const bounds = this.tryGetContentBounds();
        if (alsoShow != null) {
            return bounds != null ? bounds.union(alsoShow) : alsoShow;
        }

        return bounds;
    }

    centerContent() {
        const bounds = this.tryGetContentBounds();
        this.setView(bounds != null ? bounds.center : new Point(), this.unitLength);
    }

    /**
     * The logical box around everything of a finite size that is showing: points, circles,
     * arcs, labels. Lines, rays and graphs don't end, so they don't count. Null for nothing.
     */
    tryGetContentBounds(include = null, includeHidden = null) {
        let minX = Number.MAX_VALUE;
        let minY = Number.MAX_VALUE;
        let maxX = -Number.MAX_VALUE;
        let maxY = -Number.MAX_VALUE;
        const includePoint = point => {
            if (!point.exists()) {
                return;
            }

            minX = Math.min(minX, point.x);
            minY = Math.min(minY, point.y);
            maxX = Math.max(maxX, point.x);
            maxY = Math.max(maxY, point.y);
        };

        for (const figure of this.drawing.figures.list) {
            if ((!figure.visible && (includeHidden == null || !includeHidden(figure)))
                || !figure.exists
                || (include != null && !include(figure))) {
                continue;
            }

            // a pinned label or box is on the screen, not in the plane: nothing to fit
            if (figure.pin != null && figure.pin !== LabelPin.None) {
                continue;
            }

            if (figure.isPoint === true) {
                includePoint(figure.coordinates);
            } else if (figure instanceof Segment || figure.isPolygonalChain === true || figure.isBezier === true) {
                // the corners count even when the points themselves are hidden
                for (const vertex of figure.dependencies) {
                    if (vertex.isPoint === true) {
                        includePoint(vertex.coordinates);
                    }
                }
            } else if (figure.isBezierPath === true) {
                const curve = figure.bounds;
                if (curve.width > 0 || curve.height > 0) {
                    includePoint(curve.topLeft);
                    includePoint(curve.bottomRight);
                }
            } else if (figure.isSlider === true) {
                includePoint(figure.anchor.coordinates);
                includePoint(figure.knob.coordinates);
            } else if (figure.isArc === true && equalsWithPrecision(figure.semiMajor, figure.semiMinor)) {
                // a circular arc reaches its ends and whichever of the four compass points lie on it
                includePoint(figure.beginLocation);
                includePoint(figure.endLocation);
                const center = figure.center;
                const radius = figure.semiMajor;
                for (let quarter = 0; quarter < 4; quarter++) {
                    const angle = quarter * Math.PI / 2;
                    if (GeometryMath.isAngleBetweenAngles(angle, figure.startAngle, figure.endAngle, figure.isClockwise)) {
                        includePoint(new Point(center.x + radius * Math.cos(angle), center.y + radius * Math.sin(angle)));
                    }
                }
            } else if (figure.isEllipse === true) {
                const reach = Math.max(figure.semiMajor, figure.semiMinor);
                includePoint(figure.center.plus(new Point(reach, reach)));
                includePoint(figure.center.minus(new Point(reach, reach)));
            } else if (figure instanceof ControlBase) {
                const size = figure.measureSize();
                const topLeft = figure.coordinates;
                includePoint(topLeft);
                includePoint(new Point(topLeft.x + this.toLogicalLength(size.width), topLeft.y - this.toLogicalLength(size.height)));
            }
        }

        if (minX > maxX) {
            return null;
        }

        return new Rect(minX, minY, maxX - minX, maxY - minY);
    }

    /** How far the biggest visible point reaches out from its place, in pixels */
    getPointReach(include = null) {
        let reach = 0;
        for (const figure of this.drawing.figures.list) {
            if (figure instanceof PointBase
                && figure.visible
                && figure.exists
                && (include == null || include(figure))
                && isValidValue(figure.pointSize)) {
                reach = Math.max(reach, figure.pointSize / 2);
            }
        }

        return reach;
    }

    /** The zoom, at most unitLength, at which every visible point, shape and all, stays inside a room of this size around the center */
    limitZoomByPointReach(unitLength, center, room, include = null) {
        const limit = (unit, distance, pixels) => distance > 0 && pixels > 0 ? Math.min(unit, pixels / distance) : unit;
        for (const figure of this.drawing.figures.list) {
            if (figure instanceof PointBase
                && figure.visible
                && figure.exists
                && (include == null || include(figure))
                && isValidValue(figure.pointSize)) {
                const reach = figure.pointSize / 2;
                unitLength = limit(unitLength, Math.abs(figure.coordinates.x - center.x), room.width / 2 - reach);
                unitLength = limit(unitLength, Math.abs(figure.coordinates.y - center.y), room.height / 2 - reach);
            }
        }

        return unitLength;
    }

    // Bounds

    recalculate() {
        this.logicalViewportVertices = this.getViewportVerticesInLogical();
        if (this.logicalViewportVertices != null) {
            const vertices = this.logicalViewportVertices;
            this.minimalVisibleX = Math.min(...vertices.map(p => p.x));
            this.minimalVisibleY = Math.min(...vertices.map(p => p.y));
            this.maximalVisibleX = Math.max(...vertices.map(p => p.x));
            this.maximalVisibleY = Math.max(...vertices.map(p => p.y));
            this.drawing.recalculate();
        }
    }

    getViewportVerticesInLogical() {
        const physicalSize = this.physicalSize;
        if (!isValidPositiveValue(physicalSize.x) || !isValidPositiveValue(physicalSize.y)) {
            return this.logicalViewportVertices;
        }

        return [
            this.toLogical(new Point()),
            this.toLogical(new Point(physicalSize.x, 0)),
            this.toLogical(new Point(physicalSize.x, physicalSize.y)),
            this.toLogical(new Point(0, physicalSize.y))
        ];
    }

    /** The x of every labeled grid line in view */
    getVisibleXPoints() {
        return CoordinateSystem.gridValues(this.minimalVisibleX, this.maximalVisibleX, this.majorGridStep);
    }

    getVisibleYPoints() {
        return CoordinateSystem.gridValues(this.minimalVisibleY, this.maximalVisibleY, this.majorGridStep);
    }

    /** The x of every finer grid line in view, the labeled ones left out */
    getMinorXPoints() {
        return this.minorGridValues(this.minimalVisibleX, this.maximalVisibleX);
    }

    getMinorYPoints() {
        return this.minorGridValues(this.minimalVisibleY, this.maximalVisibleY);
    }

    static gridValues(min, max, step) {
        const first = Math.ceil(min / step);
        const last = Math.floor(max / step);
        const result = [];
        for (let index = first; index <= last; index++) {
            result.push(CoordinateSystem.gridValue(index, step));
        }

        return result;
    }

    minorGridValues(min, max) {
        const chosen = this.chooseGridStep();
        if (chosen.subdivisions === 1) {
            return [];
        }

        const step = chosen.step / chosen.subdivisions;
        const first = Math.ceil(min / step);
        const last = Math.floor(max / step);
        const result = [];
        for (let index = first; index <= last; index++) {
            if (index % chosen.subdivisions !== 0) {
                result.push(CoordinateSystem.gridValue(index, step));
            }
        }

        return result;
    }

    /** index * step without the floating point dust */
    static gridValue(index, step) {
        return roundToDigits(index * step, 10);
    }

    // Grid step

    /** The distance between the labeled grid lines, in units: 1, 2 or 5 times a power of ten */
    get majorGridStep() {
        return this.chooseGridStep().step;
    }

    /** The step and how many parts the finer lines cut it into (5, 4, or 1 when there is no room) */
    chooseGridStep() {
        const minimum = CoordinateSystem.MinimumMajorGridSpacing / this.unitLength;
        const decade = Math.pow(10, Math.floor(Math.log10(minimum)));
        let step;
        let subdivisions;
        if (decade >= minimum) {
            step = decade;
            subdivisions = 5;
        } else if (2 * decade >= minimum) {
            step = 2 * decade;
            subdivisions = 4;
        } else if (5 * decade >= minimum) {
            step = 5 * decade;
            subdivisions = 5;
        } else {
            step = 10 * decade;
            subdivisions = 5;
        }

        if (this.gridStep > 0 && step < this.gridStep) {
            step = this.gridStep;
        }

        const minorStep = step / subdivisions;
        if (minorStep * this.unitLength < CoordinateSystem.MinimumMinorGridSpacing || (this.gridStep > 0 && minorStep < this.gridStep)) {
            subdivisions = 1;
        }

        return { step, subdivisions };
    }

    // Coordinate transforms

    get cursorTolerance() {
        return this.toLogicalLength(GeometryMath.cursorTolerance);
    }

    toLogical(physicalPoint) {
        return new Point(
            (physicalPoint.x - this.origin.x) / this.unitLength,
            -(physicalPoint.y - this.origin.y) / this.unitLength).roundToEpsilon();
    }

    /** The pair as a box, its corners sorted so that P1 is the lower left (ToLogical(PointPair)) */
    toLogicalPair(pointPair) {
        let result = new PointPair(this.toLogical(pointPair.p1), this.toLogical(pointPair.p2));
        if (result.p1.x > result.p2.x) {
            result = result.reverse;
        }

        if (result.p1.y > result.p2.y) {
            const temp = result.p2.y;
            result.p2 = result.p2.withY(result.p1.y);
            result.p1 = result.p1.withY(temp);
        }

        return result;
    }

    toLogicalLength(length) {
        return length / this.unitLength;
    }

    toPhysical(logicalPoint) {
        return new Point(
            this.origin.x + logicalPoint.x * this.unitLength,
            this.origin.y - logicalPoint.y * this.unitLength);
    }

    toPhysicalPair(logicalPointPair) {
        return new PointPair(this.toPhysical(logicalPointPair.p1), this.toPhysical(logicalPointPair.p2));
    }

    toPhysicalLength(length) {
        return length * this.unitLength;
    }

    // IMovable: dragging the paper

    get isMovable() {
        return true;
    }

    moveTo(position) {
        // exactly, no rounding to whole pixels: a drag arrives as many small steps
        position = this.toPhysical(position);
        if (position.equals(this.origin)) {
            return;
        }

        this.origin = position;
        this.recalculate();
    }

    allowMove() {
        return true;
    }

    get isCoordinateSystem() {
        return true;
    }

    get coordinates() {
        return new Point();
    }
}
