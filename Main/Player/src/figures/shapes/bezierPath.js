// Port of Main/Avalonia/DynamicGeometry/Figures/Shapes/BezierPath.cs: a path of cubic Bézier
// pieces through anchor points. Each anchor has a handle on either side, an offset from the
// anchor, a point of the drawing, or left to the path (automatic: BezierPathSmoother). A
// composite like a regular polygon: the anchors are points of the drawing it is built on,
// and its pieces, its inside and its handles are its parts. Other paths can be its holes.
// Left out: every change of what the path is made of (layout changes, anchors inserted or
// dropped, holes cut, points as handles), the images of transformations, undo. A handle
// shows next to the anchor the Drag tool last pressed (showHandlesWhileDragging; there is
// no selection here) and can be dragged then.

/** What a handle is: an offset from its anchor, a point of the drawing, or left to the path */
class HandleSpec {
    constructor(offset, point, auto = false) {
        this.offset = offset ?? new Point(0, 0);
        this.point = point ?? null;
        this.auto = auto;
    }

    static get automatic() {
        return new HandleSpec(new Point(0, 0), null, true);
    }
}

/** A handle of an anchor: an offset from it, which the Drag tool changes, or a point of the drawing (then hidden, the point shows itself) */
class BezierPathHandle extends PointBase {
    constructor(owner, isIn) {
        super();
        this.owner = owner;

        /** Whether the handle is on the side of the piece that comes to its anchor (else of the one that leaves it) */
        this.isIn = isIn;

        /** Where the handle is from its anchor, unless it is a point */
        this.offset = new Point(0, 0);

        /** The point of the drawing the handle is, if it is one */
        this.point = null;

        /** Whether the path works the handle out from its anchors; offset is then what it came to last */
        this.auto = false;

        /** Set by the Drag tool unless Alt is held: the handle across the anchor follows as the mirror image of this one */
        this.mirrorsOpposite = true;
    }

    get isFigurePart() {
        return true;
    }

    get spec() {
        return new HandleSpec(this.offset, this.point, this.auto);
    }

    set(spec) {
        this.offset = spec.offset;
        this.point = spec.point;
        this.auto = spec.auto && spec.point == null;
    }

    onAddingToDrawing(drawing) {
    }

    onRemovingFromDrawing(drawing) {
    }

    // over the figures, under the points: an anchor its handle is on is taken first
    defaultZOrder() {
        return ZOrder.Points - 1;
    }

    get visible() {
        return super.visible && this.owner.isHandleShown(this);
    }

    set visible(value) {
        super.visible = value;
    }

    allowMove() {
        return !this.owner.locked && this.owner.drawing != null && this.point == null;
    }

    moveToCore(newLocation) {
        this.owner.moveHandle(this, newLocation);
    }

    capturePlace() {
        return this.owner.captureHandles(this);
    }

    restorePlace(place) {
        this.owner.restoreHandles(this, place);
    }

    recalculate() {
        this.coordinates = this.owner.handleCoordinates(this);
    }

    toString() {
        return this.owner.describePart(this);
    }
}

/** A piece of the path, from an anchor to the next: a side, styled by itself */
class BezierPathPiece extends ShapeBase {
    constructor(owner) {
        super();
        this.owner = owner;
    }

    get isFigurePart() {
        return true;
    }

    /** What a line style asks for (LineStyle.supportsFigure) */
    get isBezierPathPiece() {
        return true;
    }

    onAddingToDrawing(drawing) {
    }

    defaultZOrder() {
        return ZOrder.Figures;
    }

    // the path finds its parts (BezierPath.hitTest)
    hitTest(point) {
        return this.owner.hitTest(point) === this ? this : null;
    }

    render(renderer) {
        if (!this.isShown) {
            return;
        }

        const index = this.owner.pieces.indexOf(this);
        const commands = this.owner.pieceCommands(index);
        if (commands != null) {
            renderer.drawPath(commands, this.stroke, null);
        }
    }

    toString() {
        return this.owner.describePart(this);
    }
}

/** The filled inside: to the user, the path itself */
class BezierPathInterior extends ShapeBase {
    constructor(owner) {
        super();
        this.owner = owner;
    }

    get isFigurePart() {
        return true;
    }

    onAddingToDrawing(drawing) {
    }

    defaultZOrder() {
        return ZOrder.Polygons;
    }

    /** There while the path is filled */
    updateExistence() {
        this.exists = this.owner.exists && this.owner.filled && this.owner.pieceCount > 0;
    }

    hitTest(point) {
        return this.owner.hitTest(point) === this ? this : null;
    }

    render(renderer) {
        if (!this.isShown) {
            return;
        }

        const commands = this.owner.interiorCommands();
        if (commands != null) {
            renderer.drawPath(commands, null, this.fill, this.owner.pixelBounds());
        }
    }

    toString() {
        return this.owner.name;
    }
}

class BezierPath extends CompositeFigure {
    /** How much of the ink the lines from the anchors to their handles take */
    static HandleLineOpacity = 0.3;

    static DefaultTension = 1;
    static MinimumTension = 0.5;
    static MaximumTension = 3;

    /** In pixels, besides the cursor's tolerance: how far from the cursor a handle hidden under its anchor is taken to be there */
    static HandleReach = 4;

    static InPart = "In";
    static OutPart = "Out";
    static PiecePart = "Piece";
    static InteriorPart = "Interior";
    static AutomaticToken = "a";

    constructor() {
        super();
        this.inHandles = [];
        this.outHandles = [];
        this.pieces = [];
        this.retiredPieces = [];
        this.interior = new BezierPathInterior(this);
        this.interior.dependencies = [this];
        this.children.push(this.interior);
        this.closedValue = false;
        this.filledValue = false;
        this.smoothingValue = BezierPathSmoothing.Hobby;

        // the tension typed, and the figure it comes from instead (the last dependency), if any
        this.tension = BezierPath.DefaultTension;
        this.tensionSource = null;

        // the anchor the Drag tool pressed last, whose handles show meanwhile
        this.draggedAnchor = null;
        this.curves = [];

        // the four points of each piece: its anchor, the two handles it is pulled towards, the next anchor
        this.controls = [];
        this.curvesKnown = false;
        this.partsUnregistered = true;
        this.isOnCanvas = false;
        this.zIndex = ZOrder.Polygons;
    }

    get isLinearFigure() {
        return true;
    }

    get isBezierPath() {
        return true;
    }

    get isFigureParts() {
        return true;
    }

    get isShapeWithInterior() {
        return true;
    }

    /** A new path through the anchors with these handles (HandleSpecs) and holes, not in the drawing yet */
    static create(drawing, anchors, ins, outs, holes, closed, filled) {
        const path = new BezierPath();
        path.drawing = drawing;
        path.closedValue = closed;
        path.filledValue = filled;
        path.build(anchors, ins, outs, holes);
        return path;
    }

    /** The handles, the pieces and the dependencies, for a path made or read */
    build(anchors, ins, outs, holes) {
        for (let i = 0; i < anchors.length; i++) {
            const inHandle = new BezierPathHandle(this, true);
            const outHandle = new BezierPathHandle(this, false);
            inHandle.set(ins[i]);
            outHandle.set(outs[i]);
            this.inHandles.push(inHandle);
            this.outHandles.push(outHandle);
            this.attachPart(inHandle);
            this.attachPart(outHandle);
        }

        this.syncPieces();
        this.setDependencies(this.dependenciesOf(anchors, holes, this.tensionSource));
    }

    /** The anchors, the points of the handles (handlesInOrder), the holes, the figure the tension comes from */
    dependenciesOf(anchors, holes, source) {
        const result = [...anchors, ...this.handlesInOrder.filter(h => h.point != null).map(h => h.point), ...holes];
        if (source != null) {
            result.push(source);
        }

        return result;
    }

    // Anchors, handles and pieces

    get anchorCount() {
        return this.inHandles.length;
    }

    /** One between each two anchors, and one back to the first when closed */
    get pieceCount() {
        const count = this.anchorCount;
        return count < 2 ? 0 : this.closed ? count : count - 1;
    }

    anchor(index) {
        return this.dependencies[index];
    }

    isAnchor(figure) {
        const count = Math.min(this.anchorCount, this.dependencies.length);
        for (let i = 0; i < count; i++) {
            if (this.dependencies[i] === figure) {
                return true;
            }
        }

        return false;
    }

    /** The handles in the order their points come among the dependencies: the in and the out handle of the first anchor, of the second... */
    get handlesInOrder() {
        const result = [];
        for (let i = 0; i < this.inHandles.length; i++) {
            result.push(this.inHandles[i], this.outHandles[i]);
        }

        return result;
    }

    /** Where the holes start among the dependencies: after the anchors and the points of handles */
    get holeStart() {
        return this.anchorCount + this.handlesInOrder.filter(h => h.point != null).length;
    }

    /** The dependencies after the points of handles, but for the figure the tension comes from (the last) */
    get holeList() {
        const start = this.holeStart;
        return this.dependencies.slice(start, this.dependencies.length - (this.tensionSource != null ? 1 : 0));
    }

    /** The paths whose insides the inside of this one leaves out */
    get holes() {
        return this.holeList.filter(f => f instanceof BezierPath);
    }

    get handles() {
        return [...this.inHandles, ...this.outHandles];
    }

    /** Whether the path is the image of a transformation: its handles are hidden helper points */
    get isImage() {
        return this.handles.some(h => h.point != null && h.point.auxiliary === true);
    }

    anchorPoint(index) {
        return this.dependencies[index].coordinates;
    }

    indexOf(handle) {
        return handle.isIn ? this.inHandles.indexOf(handle) : this.outHandles.indexOf(handle);
    }

    /** Where the handle is in the plane: its point, or its offset from its anchor */
    handlePosition(handle, index) {
        return handle.point != null ? handle.point.coordinates : this.anchorPoint(index).plus(handle.offset);
    }

    inPoint(index) {
        return this.handlePosition(this.inHandles[index], index);
    }

    outPoint(index) {
        return this.handlePosition(this.outHandles[index], index);
    }

    handleCoordinates(handle) {
        const index = this.indexOf(handle);
        if (index < 0 || index >= this.anchorCount || this.dependencies.length < this.anchorCount) {
            return handle.coordinates;
        }

        return this.handlePosition(handle, index);
    }

    /** The handle across the anchor from this one */
    opposite(handle) {
        const index = this.indexOf(handle);
        if (index < 0) {
            return null;
        }

        return handle.isIn ? this.outHandles[index] : this.inHandles[index];
    }

    anchorOf(handle) {
        const index = this.indexOf(handle);
        return index >= 0 && index < this.anchorCount && index < this.dependencies.length ? this.dependencies[index] : null;
    }

    /**
     * The handle goes where it is dragged: its offset from its anchor changes (it is the user's
     * from then on), and with mirrorsOpposite the handle across the anchor is its mirror image
     * through the anchor, unless that one is a point of the drawing. Without it an automatic one
     * across stays where it was worked out to (a corner).
     */
    moveHandle(handle, coordinates) {
        const anchor = this.anchorOf(handle);
        if (anchor == null || anchor.isPoint !== true || handle.point != null) {
            return;
        }

        const offset = coordinates.minus(anchor.coordinates);
        const opposite = this.opposite(handle);
        if (opposite != null && opposite.point == null) {
            if (handle.mirrorsOpposite) {
                opposite.offset = offset.negate();
            }

            opposite.auto = false;
        }

        handle.offset = offset;
        handle.auto = false;
        this.recalculateAndUpdate();
    }

    captureHandles(handle) {
        const opposite = this.opposite(handle);
        return [handle.spec, opposite != null ? opposite.spec : new HandleSpec()];
    }

    restoreHandles(handle, place) {
        if (!Array.isArray(place) || place.length !== 2) {
            return;
        }

        handle.set(place[0]);
        const opposite = this.opposite(handle);
        if (opposite != null) {
            opposite.set(place[1]);
        }

        this.recalculateAndUpdate();
    }

    /** The path worked out and drawn again, and what is built on it */
    recalculateAndUpdate() {
        if (this.drawing == null) {
            return;
        }

        this.recalculateAndUpdateVisual();
        this.recalculateAllDependents();
    }

    /** What took the place of a point that is a handle is that handle's point from now on: the dependencies say */
    onDependenciesChanged() {
        super.onDependenciesChanged?.();
        let index = this.anchorCount;
        for (const handle of this.handlesInOrder) {
            if (handle.point == null) {
                continue;
            }

            if (index < this.dependencies.length && this.dependencies[index].isPoint === true) {
                handle.point = this.dependencies[index];
            }

            index++;
        }

        if (this.tensionSource != null && this.dependencies.length > 0) {
            this.tensionSource = this.dependencies[this.dependencies.length - 1];
        }

        this.curvesKnown = false;
    }

    /** The dependencies set, and listed with what they are, if they were */
    setDependencies(dependencies) {
        const registered = this.dependencies.length > 0 && this.dependencies.every(d => d.dependents.includes(this));
        if (registered) {
            this.unregisterFromDependencies();
        }

        this.dependencies = dependencies;
        if (registered) {
            this.registerWithDependencies();
        }
    }

    /** The pieces the anchors need, made or taken away at the end; a piece that leaves is kept, and is the one that comes back */
    syncPieces() {
        const count = this.pieceCount;
        while (this.pieces.length < count) {
            const isNew = this.retiredPieces.length === 0;
            const piece = isNew ? new BezierPathPiece(this) : this.retiredPieces.pop();
            const common = isNew ? DependentPolygonBase.commonStyle(this.pieces) : null;
            if (common != null) {
                piece.style = common;
            }

            this.pieces.push(piece);
            this.attachPart(piece);
        }

        while (this.pieces.length > count) {
            const piece = this.pieces.pop();
            this.removePart(piece);
            this.retiredPieces.push(piece);
        }
    }

    // Parts: figures of the library that the drawing never sees, built on the path, listed with it only while it is in the drawing

    attachPart(part) {
        part.dependencies = [this];
        part.drawing = this.drawing;
        part.visible = this.visible;
        if (!this.partsUnregistered) {
            part.registerWithDependencies();
        }

        this.children.push(part);
        if (this.isOnCanvas) {
            part.onAddingToCanvas(this.drawing.canvas);
        }
    }

    removePart(part) {
        part.unregisterFromDependencies();
        const index = this.children.indexOf(part);
        if (index >= 0) {
            this.children.splice(index, 1);
        }

        if (this.isOnCanvas) {
            part.onRemovingFromCanvas(this.drawing.canvas);
        }
    }

    onAddingToDrawing(drawing) {
        super.onAddingToDrawing(drawing);
        if (this.partsUnregistered) {
            this.partsUnregistered = false;
            for (const part of this.children) {
                part.registerWithDependencies();
            }
        }
    }

    onRemovingFromDrawing(drawing) {
        super.onRemovingFromDrawing(drawing);
        if (!this.partsUnregistered) {
            this.partsUnregistered = true;
            for (const part of this.children) {
                part.unregisterFromDependencies();
            }
        }
    }

    onAddingToCanvas(newContainer) {
        super.onAddingToCanvas(newContainer);
        this.isOnCanvas = true;
    }

    onRemovingFromCanvas(leavingContainer) {
        super.onRemovingFromCanvas(leavingContainer);
        this.isOnCanvas = false;
    }

    /** In1, Out1 (the handles of the first anchor), Piece1 (from the first anchor to the second), Interior */
    getPartName(part) {
        let index = this.inHandles.indexOf(part);
        if (index >= 0) {
            return BezierPath.InPart + (index + 1);
        }

        index = this.outHandles.indexOf(part);
        if (index >= 0) {
            return BezierPath.OutPart + (index + 1);
        }

        index = this.pieces.indexOf(part);
        if (index >= 0) {
            return BezierPath.PiecePart + (index + 1);
        }

        return part === this.interior ? BezierPath.InteriorPart : null;
    }

    getPart(partName) {
        if (partName === BezierPath.InteriorPart) {
            return this.interior;
        }

        let index = DependentPolygonBase.tryGetIndex(partName, BezierPath.InPart);
        if (index != null) {
            return index >= 1 && index <= this.inHandles.length ? this.inHandles[index - 1] : null;
        }

        index = DependentPolygonBase.tryGetIndex(partName, BezierPath.OutPart);
        if (index != null) {
            return index >= 1 && index <= this.outHandles.length ? this.outHandles[index - 1] : null;
        }

        index = DependentPolygonBase.tryGetIndex(partName, BezierPath.PiecePart);
        if (index != null) {
            return index >= 1 && index <= this.pieces.length ? this.pieces[index - 1] : null;
        }

        return null;
    }

    /** The pieces: a click selects one by itself; the inside selects the path */
    get selectableParts() {
        return this.pieces;
    }

    /** "Side 2 of ABC", "Handle of B toward C" */
    describePart(part) {
        const index = this.pieces.indexOf(part);
        if (index >= 0) {
            return "Side " + (index + 1) + " of " + this.name;
        }

        if (part instanceof BezierPathHandle) {
            const anchor = this.anchorOf(part);
            if (anchor != null) {
                const count = this.anchorCount;
                const anchorIndex = this.indexOf(part);
                const toward = (anchorIndex + (part.isIn ? count - 1 : 1)) % count;
                return "Handle of " + anchor.name + " toward " + this.dependencies[toward].name;
            }
        }

        return this.name;
    }

    // Handles shown

    /** The anchor the Drag tool pressed (a handle: its anchor), whose handles show from then on; null for none */
    static showHandlesWhileDragging(drawing, point) {
        if (drawing == null) {
            return;
        }

        for (const path of drawing.figures.list) {
            if (!(path instanceof BezierPath)) {
                continue;
            }

            const dragged = point instanceof BezierPathHandle && point.owner === path ? path.anchorOf(point) : point;
            const anchor = dragged != null && path.isAnchor(dragged) ? dragged : null;
            if (path.draggedAnchor !== anchor) {
                path.draggedAnchor = anchor;
                path.refreshHandles();
            }
        }
    }

    /** Whether the handle shows: next to the active anchor, where it bends a piece; a handle that is a point doesn't (the point shows itself) */
    isHandleShown(handle) {
        return handle.point == null && this.isNextToActiveAnchor(handle);
    }

    isNextToActiveAnchor(handle) {
        const piece = this.bendsPiece(handle);
        return piece != null && (this.isActive(this.dependencies[piece.index]) || this.isActive(this.dependencies[piece.other]));
    }

    /** Whether the handle bends a piece (not the outer handle of an open path's end), and the anchors of that piece: { index, other }, or null */
    bendsPiece(handle) {
        const count = this.anchorCount;
        const index = this.indexOf(handle);
        if (!this.exists || !this.visible || this.drawing == null || this.dependencies.length < count || index < 0 || count < 2) {
            return null;
        }

        let other = handle.isIn ? index - 1 : index + 1;
        if (this.closed) {
            other = (other + count) % count;
        } else if (other < 0 || other >= count) {
            return null;
        }

        return { index, other };
    }

    /** The anchor's own handles that bend a piece and are at the point, shown or not */
    handlesOn(anchor, point) {
        const index = this.dependencies.indexOf(anchor);
        if (index < 0 || index >= this.anchorCount || this.drawing == null) {
            return [];
        }

        const reach = this.drawing.coordinateSystem.cursorTolerance + this.toLogicalLength(BezierPath.HandleReach);
        return [this.inHandles[index], this.outHandles[index]].filter(handle =>
            handle.point == null
            && handle.exists
            && this.bendsPiece(handle) != null
            && this.handleCoordinates(handle).distance(point) <= reach);
    }

    isActive(anchor) {
        return anchor.selected || anchor === this.draggedAnchor;
    }

    /** The handles where they are */
    refreshHandles() {
        for (const handle of this.handles) {
            if (handle.visible && handle.exists) {
                handle.updateVisual();
            }
        }

        this.drawing?.canvas?.invalidate?.();
    }

    // Properties

    /** Whether a piece goes from the last anchor back to the first */
    get closed() {
        return this.closedValue;
    }

    set closed(value) {
        if (this.closedValue === value) {
            return;
        }

        this.closedValue = value;
        this.syncPieces();
        if (this.drawing != null) {
            this.recalculateAndUpdate();
        }
    }

    /** Whether the inside is filled: an open path as if a straight line closed it */
    get filled() {
        return this.filledValue;
    }

    set filled(value) {
        if (this.filledValue === value) {
            return;
        }

        this.filledValue = value;
        if (this.drawing != null) {
            this.recalculateAndUpdate();
        }
    }

    get smoothing() {
        return this.smoothingValue;
    }

    set smoothing(value) {
        if (this.smoothingValue === value) {
            return;
        }

        this.smoothingValue = value;
        if (this.drawing != null) {
            this.recalculateAndUpdate();
        }
    }

    /** How tight the automatic handles are: typed, or a figure with a number (a slider), which it then follows */
    get tensionValue() {
        const source = this.tensionSource;
        if (source == null) {
            return this.tension;
        }

        if (source.isNumber === true) {
            return source.value;
        }

        if (source.isLengthProvider === true) {
            return source.length;
        }

        return NaN;
    }

    /** The style of the inside, which is the path's own */
    get style() {
        return this.interior.style;
    }

    set style(value) {
        this.interior.style = value;
    }

    get center() {
        const count = Math.min(this.anchorCount, this.dependencies.length);
        if (count === 0) {
            return new Point();
        }

        let x = 0;
        let y = 0;
        for (let i = 0; i < count; i++) {
            const point = this.anchorPoint(i);
            x += point.x;
            y += point.y;
        }

        return new Point(x / count, y / count);
    }

    toString() {
        return this.name;
    }

    // Working it out

    /** The pieces worked out, also before the first recalculate: a point on the path read from a file asks where it is while the file is still being read */
    get curveInfos() {
        if (!this.curvesKnown) {
            this.calculateCurves();
        }

        return this.curves;
    }

    calculateCurves() {
        const count = this.dependencies.length >= this.anchorCount ? this.pieceCount : 0;
        const anchors = this.anchorCount;
        if (count > 0) {
            this.smoothHandles();
        }

        const result = new Array(count);
        const points = new Array(count);
        for (let i = 0; i < count; i++) {
            const next = (i + 1) % anchors;
            points[i] = [this.anchorPoint(i), this.outPoint(i), this.inPoint(next), this.anchorPoint(next)];
            result[i] = new BezierInfo(points[i][0], points[i][1], points[i][2], points[i][3]);
        }

        this.curves = result;
        this.controls = points;
        this.curvesKnown = true;
    }

    /** The automatic handles worked out from where the anchors and the other handles are now */
    smoothHandles() {
        if (!this.handles.some(h => h.auto)) {
            return;
        }

        const count = this.anchorCount;
        const points = new Array(count);
        for (let i = 0; i < count; i++) {
            points[i] = this.anchorPoint(i);
            if (!points[i].exists()) {
                return;
            }
        }

        const given = (handle, index) => handle.auto ? null : this.handlePosition(handle, index).minus(points[index]);
        const result = BezierPathSmoother.smooth(
            points,
            this.inHandles.map(given),
            this.outHandles.map(given),
            this.closed,
            this.smoothing,
            this.tensionValue);
        for (let i = 0; i < count; i++) {
            if (this.inHandles[i].auto) {
                this.inHandles[i].offset = result.ins[i];
            }

            if (this.outHandles[i].auto) {
                this.outHandles[i].offset = result.outs[i];
            }
        }
    }

    /** Whether every point the pieces are drawn through is somewhere */
    curvesExist() {
        for (const curve of this.curveInfos) {
            if (curve.points == null || !curve.points.every(point => point.exists())) {
                return false;
            }
        }

        return true;
    }

    /** The path exists while its anchors and the points of its handles do, and while the tension is a number above 0 if a handle is worked out with it; a hole that doesn't is left out */
    updateExistence() {
        const count = Math.min(this.holeStart, this.dependencies.length);
        let exists = this.anchorCount >= 2 && this.dependencies.length >= this.anchorCount;
        for (let i = 0; i < count && exists; i++) {
            exists = this.dependencies[i].exists;
        }

        if (exists && this.smoothing !== BezierPathSmoothing.None && this.handles.some(h => h.auto)) {
            const value = this.tensionValue;
            exists = (this.tensionSource == null || this.tensionSource.exists) && value > 0 && isValidValue(value);
        }

        this.exists = exists;
        for (const part of this.children) {
            part.updateExistence();
        }
    }

    recalculate() {
        this.calculateCurves();
        for (const handle of this.handles) {
            handle.recalculate();
        }
    }

    updateVisual() {
        if (this.drawing == null) {
            return;
        }

        this.refreshHandles();
    }

    /** The commands of one piece, in pixels */
    pieceCommands(index) {
        if (index < 0 || index >= this.controls.length || !this.exists || !this.curvesExist()) {
            return null;
        }

        const c = this.controls[index].map(p => this.toPhysical(p));
        if (c.some(p => !p.exists())) {
            return null;
        }

        return [
            { op: "move", x: c[0].x, y: c[0].y },
            { op: "cubic", x1: c[1].x, y1: c[1].y, x2: c[2].x, y2: c[2].y, x: c[3].x, y: c[3].y }
        ];
    }

    /** The outline as one figure, closed (with a straight line, when the path is open), in pixels */
    outlineCommands() {
        if (!this.curvesKnown) {
            this.calculateCurves();
        }

        if (this.controls.length === 0) {
            return null;
        }

        const start = this.toPhysical(this.controls[0][0]);
        const commands = [{ op: "move", x: start.x, y: start.y }];
        for (const points of this.controls) {
            const c = points.map(p => this.toPhysical(p));
            if (c.some(p => !p.exists())) {
                return null;
            }

            commands.push({ op: "cubic", x1: c[1].x, y1: c[1].y, x2: c[2].x, y2: c[2].y, x: c[3].x, y: c[3].y });
        }

        commands.push({ op: "close" });
        return commands;
    }

    /** The inside: the outline filled (even-odd), the holes taken out (their outlines in the same path, which even-odd leaves out) */
    interiorCommands() {
        const outline = this.outlineCommands();
        if (outline == null) {
            return null;
        }

        const commands = [...outline];
        for (const hole of this.holes) {
            if (hole.exists && hole.pieceCount > 0 && hole.curvesExist()) {
                const holeOutline = hole.outlineCommands();
                if (holeOutline != null) {
                    commands.push(...holeOutline);
                }
            }
        }

        return commands;
    }

    /** The outline as a polygon through the points the pieces are drawn through */
    outlinePolygon() {
        const result = [];
        for (const curve of this.curveInfos) {
            if (curve.points != null) {
                result.push(...curve.points);
            }
        }

        return result;
    }

    /** Whether the point is inside: inside the outline (even-odd) and in none of the holes */
    isInside(point) {
        if (this.pieceCount === 0 || !GeometryMath.isPointInPolygon(this.outlinePolygon(), point)) {
            return false;
        }

        return !this.holes.some(hole => hole.exists && hole.pieceCount > 0 && GeometryMath.isPointInPolygon(hole.outlinePolygon(), point));
    }

    /** The smallest box around the pieces as drawn, logical */
    get bounds() {
        const points = this.outlinePolygon().filter(point => point.exists());
        if (points.length === 0) {
            return new Rect();
        }

        const left = Math.min(...points.map(p => p.x));
        const right = Math.max(...points.map(p => p.x));
        const bottom = Math.min(...points.map(p => p.y));
        const top = Math.max(...points.map(p => p.y));
        return new Rect(left, bottom, right - left, top - bottom);
    }

    /** The box around the outline in pixels (a gradient fill spans it) */
    pixelBounds() {
        const points = this.outlinePolygon().filter(point => point.exists()).map(p => this.toPhysical(p));
        return points.length > 0 ? CanvasRenderer.boxOf(points) : null;
    }

    render(renderer) {
        if (!this.visible || !this.exists) {
            return;
        }

        // the inside under the sides, the handles and their dotted lines over them
        if (this.interior.exists) {
            this.interior.render(renderer);
        }

        for (const piece of this.pieces) {
            if (piece.exists) {
                piece.render(renderer);
            }
        }

        this.renderHandles(renderer);
    }

    /** A dotted line from an anchor to each handle shown next to it, and to each point that is a handle there, then the handles */
    renderHandles(renderer) {
        const shown = this.handles.filter(handle => handle.point == null
            ? handle.visible && handle.exists
            : handle.point.visible && handle.point.exists && this.isNextToActiveAnchor(handle));
        if (shown.length === 0 || this.drawing == null) {
            return;
        }

        const ink = AppTheme.of(this.drawing).ink;
        const stroke = { color: ink.withAlpha(Math.round(ink.a * BezierPath.HandleLineOpacity)), width: 1, dash: [1, 2] };
        for (const handle of shown) {
            const anchor = this.anchorOf(handle);
            if (anchor == null || anchor.isPoint !== true) {
                continue;
            }

            const from = this.toPhysical(anchor.coordinates);
            const to = this.toPhysical(this.handleCoordinates(handle));
            if (from.exists() && to.exists()) {
                renderer.drawLine(from, to, stroke);
            }
        }

        for (const handle of shown) {
            if (handle.point == null) {
                handle.render(renderer);
            }
        }
    }

    // Hit testing

    /** A handle that shows (for the Drag tool alone), else a side, else the inside if it is filled; by the numbers, shown or not */
    hitTest(point) {
        if (this.drawing == null || this.anchorCount < 2) {
            return null;
        }

        if (this.drawing.behavior instanceof Dragger) {
            for (const handle of this.handles) {
                if (handle.visible && handle.exists && handle.hitTest(point) != null) {
                    return handle;
                }
            }
        }

        const curves = this.curveInfos;
        for (let i = 0; i < curves.length && i < this.pieces.length; i++) {
            // (the corners of the polyline count too: on the outer side of a bend a click is over neither piece of it)
            const reach = this.toLogicalLength(this.pieces[i].strokeThickness / 2 + GeometryMath.cursorTolerance);
            if (curves[i].points != null && GeometryMath.isPointOnPolygonalChain(curves[i].points, point, reach, false)) {
                return this.pieces[i];
            }
        }

        if (this.filled && this.interior.exists && this.isInside(point)) {
            return this.interior;
        }

        return null;
    }

    // A point on the path

    /** What a point put on the figure under the cursor goes on: a side of a path stands for the path */
    static pointHolder(figure) {
        return figure instanceof BezierPathPiece ? figure.owner : figure;
    }

    /** A point on the path is on a piece: the whole number of the parameter is which (from 0), the rest the cubic's own parameter */
    getParameterDomain() {
        return [0, this.pieceCount];
    }

    getPointFromParameter(parameter) {
        const curves = this.curveInfos;
        const slack = 1e-9;
        if (curves.length === 0 || Number.isNaN(parameter) || parameter < -slack || parameter > curves.length + slack) {
            return new Point(NaN, NaN);
        }

        parameter = Math.max(0, Math.min(curves.length, parameter));
        const index = Math.min(Math.floor(parameter), curves.length - 1);
        return curves[index].getPoint(parameter - index);
    }

    getNearestParameterFromPoint(point) {
        const curves = this.curveInfos;
        let bestDistance = Number.MAX_VALUE;
        let bestIndex = 0;
        let bestT = 0;
        for (let i = 0; i < curves.length; i++) {
            const points = curves[i].points;
            if (points == null) {
                continue;
            }

            for (let j = 0; j + 1 < points.length; j++) {
                const found = BezierPath.distanceToSegment(point, points[j], points[j + 1]);
                if (found.distance < bestDistance) {
                    bestDistance = found.distance;
                    bestIndex = i;
                    bestT = (j + found.ratio) / (points.length - 1);
                }
            }
        }

        if (curves.length === 0) {
            return 0;
        }

        // the polyline is a few hundredths of the curve off: the nearest place on the curve itself, in the step around the one found
        const best = curves[bestIndex];
        const step = 1 / (BezierInfo.NumberOfPoints - 1);
        let low = Math.max(0, bestT - step);
        let high = Math.min(1, bestT + step);
        for (let k = 0; k < 40; k++) {
            const third = (high - low) / 3;
            if (best.getPoint(low + third).distance(point) < best.getPoint(high - third).distance(point)) {
                high -= third;
            } else {
                low += third;
            }
        }

        // and a few steps of Newton's method on the squared distance, which the search above leaves a little off
        let t = (low + high) / 2;
        const c = this.controls[bestIndex];
        for (let k = 0; k < 4; k++) {
            const u = 1 - t;
            const at = best.getPoint(t);
            const dx = at.x - point.x;
            const dy = at.y - point.y;
            const d1x = 3 * u * u * (c[1].x - c[0].x) + 6 * u * t * (c[2].x - c[1].x) + 3 * t * t * (c[3].x - c[2].x);
            const d1y = 3 * u * u * (c[1].y - c[0].y) + 6 * u * t * (c[2].y - c[1].y) + 3 * t * t * (c[3].y - c[2].y);
            const d2x = 6 * u * (c[2].x - 2 * c[1].x + c[0].x) + 6 * t * (c[3].x - 2 * c[2].x + c[1].x);
            const d2y = 6 * u * (c[2].y - 2 * c[1].y + c[0].y) + 6 * t * (c[3].y - 2 * c[2].y + c[1].y);
            const slope = dx * d1x + dy * d1y;
            const curvature = d1x * d1x + d1y * d1y + dx * d2x + dy * d2y;
            if (curvature <= 0) {
                break;
            }

            t = Math.max(0, Math.min(1, t - slope / curvature));
        }

        return bestIndex + t;
    }

    static distanceToSegment(point, start, end) {
        const dx = end.x - start.x;
        const dy = end.y - start.y;
        const lengthSquared = dx * dx + dy * dy;
        const ratio = lengthSquared === 0
            ? 0
            : Math.max(0, Math.min(1, ((point.x - start.x) * dx + (point.y - start.y) * dy) / lengthSquared));
        const nearest = new Point(start.x + ratio * dx, start.y + ratio * dy);
        return { distance: nearest.distance(point), ratio };
    }

    // File

    /**
     * The anchors are the first dependencies, and their handles Path="C 1,0.5 a L C #4 0,2 ...",
     * a piece from each anchor to the next, the closing one included whether or not the path
     * is closed: "L" for a piece whose two handles are on their anchors, else "C", the first
     * anchor's out handle and the second one's in handle - each an offset, a point of the
     * drawing by its place among the dependencies ("#4"), or "a", automatic. The dependencies
     * after those are the holes, and last the figure the tension comes from (Tension="#7").
     */
    readXml(element) {
        this.closedValue = Xml.readBool(element, "Closed", false);
        this.filledValue = Xml.readBool(element, "Filled", false);
        const listed = [...this.dependencies];
        let tokens = (element.getAttribute("Path") ?? "").split(/\s+/).filter(token => token !== "");
        let count = tokens.filter(token => token === "L" || token === "C").length;
        if (count < 2 || count > listed.length || !listed.slice(0, count).every(d => d.isPoint === true)) {
            count = 0;
            while (count < listed.length && listed[count].isPoint === true) {
                count++;
            }

            tokens = [];
        }

        const ins = [];
        const outs = [];
        for (let i = 0; i < count; i++) {
            ins.push(new HandleSpec());
            outs.push(new HandleSpec());
        }

        const referenced = new Set();
        const read = text => {
            if (text === BezierPath.AutomaticToken) {
                return HandleSpec.automatic;
            }

            if (text.startsWith("#") && /^\d+$/.test(text.substring(1))) {
                const index = parseInt(text.substring(1), 10);
                if (index >= count && index < listed.length && listed[index].isPoint === true) {
                    referenced.add(index);
                    return new HandleSpec(new Point(0, 0), listed[index]);
                }
            }

            return new HandleSpec(BezierPath.tryParsePoint(text) ?? new Point(0, 0), null);
        };

        let piece = 0;
        for (let i = 0; i < tokens.length && piece < count; piece++) {
            if (tokens[i] === "C" && i + 2 < tokens.length) {
                outs[piece] = read(tokens[i + 1]);
                ins[(piece + 1) % count] = read(tokens[i + 2]);
                i += 3;
            } else {
                i++;
            }
        }

        const smoothingText = element.getAttribute("Smoothing");
        this.smoothingValue = smoothingText != null && BezierPathSmoothing[smoothingText] != null ? BezierPathSmoothing[smoothingText] : BezierPathSmoothing.Hobby;
        this.tension = BezierPath.DefaultTension;
        this.tensionSource = null;
        const tensionText = element.getAttribute("Tension");
        if (tensionText != null && tensionText.startsWith("#") && /^\d+$/.test(tensionText.substring(1))) {
            const tensionIndex = parseInt(tensionText.substring(1), 10);
            if (tensionIndex >= count && tensionIndex < listed.length) {
                this.tensionSource = listed[tensionIndex];
                referenced.add(tensionIndex);
            }
        } else if (tensionText != null) {
            const typed = Xml.parseDouble(tensionText);
            if (isValidValue(typed)) {
                this.tension = typed;
            }
        }

        const holes = [];
        for (let index = count; index < listed.length; index++) {
            if (!referenced.has(index) && listed[index] instanceof BezierPath) {
                holes.push(listed[index]);
            }
        }

        this.build(listed.slice(0, count), ins, outs, holes);

        // visibility and the style of the inside, onto the parts made
        super.readXml(element);
        this.readPartStyles(element);
        this.curvesKnown = false;
    }

    static tryParsePoint(text) {
        const parts = text.split(",");
        if (parts.length !== 2) {
            return null;
        }

        const x = Xml.parseDouble(parts[0]);
        const y = Xml.parseDouble(parts[1]);
        return isValidValue(x) && isValidValue(y) ? new Point(x, y) : null;
    }

    /** The styles of the sides and handles, as a regular polygon's parts: <Sides Style>, <Handles Style>, <Part Name Style> */
    readPartStyles(element) {
        const manager = this.drawing?.styleManager;
        if (manager == null) {
            return;
        }

        const apply = (part, styleName) => {
            const style = styleName != null ? manager.get(styleName) : null;
            if (part != null && style != null && style.constructor.supportsFigure(part)) {
                part.style = style;
            }
        };
        const sidesStyle = Xml.element(element, "Sides")?.getAttribute("Style") ?? null;
        for (const piece of this.pieces) {
            apply(piece, sidesStyle);
        }

        const handlesStyle = Xml.element(element, "Handles")?.getAttribute("Style") ?? null;
        for (const handle of this.handles) {
            apply(handle, handlesStyle);
        }

        for (const part of Xml.elements(element, "Part")) {
            apply(this.getPart(part.getAttribute("Name")), part.getAttribute("Style"));
        }
    }
}

FigureTypes.register("BezierPath", BezierPath);
