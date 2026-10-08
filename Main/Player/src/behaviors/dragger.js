// Port of Main/Avalonia/DynamicGeometry/Behaviors/Dragger.cs: the Drag tool, the player's one
// tool. A press takes a figure (or a part of one) and the drag moves what can move: the
// figure itself if it is free, the free points it is built on otherwise, the view when
// nothing is there. A click on a show/hide box ticks it. Left out: Alt-drag snapping and
// joining, the selection, the Tab choice of a Bezier handle under its anchor, the context
// menu, the Delete key.

class Dragger extends Behavior {
    /** In pixels: how far the cursor goes from where it was pressed before the press is a drag and not a click */
    static DragThreshold = 3;

    constructor() {
        super();
        this.moving = null;
        this.found = null;

        // a label the press found that can't be dragged: the drag moves the view, a click ticks a box
        this.pressedFixedLabel = null;
        this.toRecalculate = null;
        this.offsetFromFigureLeftTopCorner = new Point();
        this.oldCoordinates = new Point();
        this.coordinatesOnMouseDown = new Point();
        this.startedMoving = false;
        this.clickCount = 0;
        this.lastPressTime = 0;
        this.lastPressPosition = null;
    }

    get name() {
        return "Drag";
    }

    get hintText() {
        return "Use this tool to drag points and figures.";
    }

    /** Two presses within half a second at one place are a double click (e.detail is not reliable with pointer capture) */
    countClicks(e) {
        const now = performance.now();
        const position = this.position(e);
        if (now - this.lastPressTime < 500 && this.lastPressPosition != null && position.distance(this.lastPressPosition) < Dragger.DragThreshold * 2) {
            this.clickCount++;
        } else {
            this.clickCount = 1;
        }

        this.lastPressTime = now;
        this.lastPressPosition = position;
        return this.clickCount;
    }

    mouseDown(e) {
        // a drag whose release never arrived
        this.release();

        // (not on a show/hide box, which two quick clicks tick and untick)
        if (this.countClicks(e) === 2 && !(this.drawing.figures.hitTest(this.coordinatesOf(e, false, false, false)) instanceof ShowHideControl)) {
            this.drawing.zoomToFit();
            this.clickCount = 0;
            return;
        }

        this.offsetFromFigureLeftTopCorner = this.coordinatesOf(e, false, false, false);
        this.oldCoordinates = this.offsetFromFigureLeftTopCorner;
        this.coordinatesOnMouseDown = this.offsetFromFigureLeftTopCorner;
        this.startedMoving = false;

        this.moving = [];
        let roots = null;
        let isLocked = false;

        this.found = this.drawing.figures.hitTest(this.offsetFromFigureLeftTopCorner);

        // an axis is the grid's, for the tools to build on: a press on it is a press on the paper
        if (this.found != null && this.found.isAxisLine === true) {
            this.found = null;
        }

        // labels that can't be dragged are paper: the drag moves the view, and one that starts
        // on a caption takes the captions along
        let captions = null;
        if (this.found instanceof ControlBase && this.drawing.fixedLabels.has(this.found)) {
            const label = this.found;
            const isCaption = label instanceof Label && label.pin !== LabelPin.None;
            if (isCaption) {
                captions = new PinnedLabelScroll(this.drawing);
            }

            if (!isCaption || e.pointerType !== "touch") {
                this.pressedFixedLabel = label;
            }

            this.found = null;
        }

        const found = this.found;

        // the handles next to an anchor of a Bezier path show from the press on it until the
        // next press elsewhere (there is no selection to keep them shown by)
        BezierPath.showHandlesWhileDragging(this.drawing, found);

        // a figure with parts (a slider) says which of them the press takes
        const oneMovable = found != null && found.isMovableParts === true
            ? found.findMovablePart(this.offsetFromFigureLeftTopCorner)
            : found != null && found.isMovable === true ? found : null;
        if (oneMovable != null && (found.locked || oneMovable.allowMove())) {
            if (found.locked) {
                isLocked = true;
            } else if (oneMovable.allowMove()) {
                if (oneMovable.isPoint === true) {
                    // when we drag a point, we want it to snap to the cursor
                    this.offsetFromFigureLeftTopCorner = new Point();
                    this.oldCoordinates = oneMovable.coordinates;
                } else {
                    // other stuff (text labels) keeps the grab offset: no snap to the cursor
                    this.offsetFromFigureLeftTopCorner = this.offsetFromFigureLeftTopCorner.minus(oneMovable.coordinates);
                }

                roots = DependencyAlgorithms.findRoots(f => f.dependents, [found]);

                // (a show/hide box holds figures, it isn't built on them: locked, it stays
                // where it is, and they still move)
                if (roots.every(root => !root.locked || root instanceof ShowHideControl)) {
                    this.moving.push(oneMovable);
                    roots = [found];
                } else {
                    isLocked = true;
                }
            }
        } else if (found != null) {
            if (!found.locked) {
                // a Number has no place to move; the drag goes to the points
                const allRoots = DependencyAlgorithms.findRoots(f => f.dependencies, [found])
                    .filter(root => root.isNumber !== true);

                // a point by coordinates stays where its X and Y say: the drag goes to the
                // other roots, and a figure built on such points alone doesn't move at all
                roots = allRoots.filter(root => !(root instanceof PointByCoordinates));
                if (roots.length === 0 && allRoots.length > 0) {
                    isLocked = true;
                } else if (roots.every(root => root.isMovable === true)) {
                    if (roots.every(root => root.allowMove())) {
                        this.moving.push(...roots);
                    } else {
                        isLocked = true;
                    }
                }
            } else {
                isLocked = true;
            }
        }

        if (roots != null) {
            this.toRecalculate = DependencyAlgorithms.findDescendants(f => f.dependents, roots);
            this.toRecalculate.reverse();
        } else {
            this.toRecalculate = null;
        }

        if (this.moving.length === 0 && !isLocked && !this.drawing.coordinateGrid.locked) {
            this.moving.push(this.drawing.coordinateSystem);
            if (captions != null) {
                this.moving.push(captions);
            }

            this.toRecalculate = null;
        }
    }

    mouseMove(e) {
        // a move with no button down is no drag: the release of this press never arrived
        if (this.moving != null && e.pointerType !== "touch" && (e.buttons & 1) === 0) {
            this.release();
            return;
        }

        if (this.moving == null) {
            return;
        }

        const currentCoordinates = this.coordinates(e);
        if (!this.startedMoving) {
            if (currentCoordinates.equals(this.coordinatesOnMouseDown)) {
                return;
            }

            // a press that wobbles is still a click
            const coordinateSystem = this.drawing.coordinateSystem;
            const wobble = coordinateSystem.toPhysical(this.coordinatesOf(e, false, false, false))
                .distance(coordinateSystem.toPhysical(this.coordinatesOnMouseDown));
            if (wobble < Dragger.DragThreshold) {
                return;
            }

            this.startedMoving = true;
        }

        if (this.moving.length > 0) {
            let offset = currentCoordinates.minus(this.oldCoordinates);
            if (this.moving.length === 1 && this.moving[0] instanceof PointLabel) {
                // a point label is confined to an orbit around its point: move it to where the
                // cursor wants it (as far as allowed)
                const pointLabel = this.moving[0];
                const desired = currentCoordinates.minus(this.offsetFromFigureLeftTopCorner);
                offset = pointLabel.clampPosition(desired).minus(pointLabel.coordinates);
            } else if (this.moving.length === 1 && this.moving[0] instanceof PointOnFigure) {
                // a point on a figure stops at the end of a segment or ray: where it lands, not where the cursor went
                const pointOnFigure = this.moving[0];
                const figure = pointOnFigure.linearFigure;
                const landing = figure.getPointFromParameter(
                    figure.getNearestParameterFromPoint(pointOnFigure.coordinates.plus(offset)));
                offset = landing.minus(pointOnFigure.coordinates);
            }

            Actions.move(this.drawing, this.moving, offset, this.toRecalculate);
        }

        // if you're dragging the coordinate plane itself, the origin changes, so the point's
        // coordinates have to be read again in the new coordinate system
        this.oldCoordinates = this.coordinates(e);
        if (this.moving != null
            && this.moving.length === 1
            && this.moving[0].isPoint === true
            && this.found != null
            && (this.found === this.moving[0] || this.found.isMovableParts === true)) {
            this.oldCoordinates = this.moving[0].coordinates;
        }
    }

    mouseUp(e) {
        // a press that didn't become a drag: on a show/hide box it ticks the box
        if (this.moving != null && !this.startedMoving) {
            const box = this.found ?? this.pressedFixedLabel;
            if (box instanceof ShowHideControl && !this.isCtrlPressed()) {
                box.click();
                this.drawing.canvas?.invalidate?.();
            }
        }

        this.release();
    }

    stopping() {
        this.release();
        super.stopping();
    }

    get drawing() {
        return super.drawing;
    }

    set drawing(value) {
        if (value !== super.drawing) {
            this.release();
        }

        super.drawing = value;
    }

    /** The press is over: nothing is held any more */
    release() {
        this.startedMoving = false;
        this.moving = null;
        this.found = null;
        this.pressedFixedLabel = null;
        this.toRecalculate = null;
    }

    /** A hand where a press would move something or tick a box, an arrow elsewhere */
    getCursor(coordinates) {
        const found = this.drawing.figures.hitTest(coordinates);
        if (found == null || found.isAxisLine === true) {
            return "default";
        }

        if (found instanceof ShowHideControl) {
            return "pointer";
        }

        if (found instanceof ControlBase && this.drawing.fixedLabels.has(found)) {
            return "default";
        }

        const oneMovable = found.isMovableParts === true ? found.findMovablePart(coordinates) : found.isMovable === true ? found : null;
        if (oneMovable != null) {
            return !found.locked && oneMovable.allowMove() ? "pointer" : "default";
        }

        if (found.locked) {
            return "default";
        }

        const roots = DependencyAlgorithms.findRoots(f => f.dependencies, [found])
            .filter(root => root.isNumber !== true && !(root instanceof PointByCoordinates));
        return roots.length > 0 && roots.every(root => root.isMovable === true && root.allowMove()) ? "pointer" : "default";
    }
}

/**
 * Port of Behaviors/PinnedLabelScroll.cs: the pinned labels of the drawing as one movable,
 * so that a drag on a caption scrolls every pinned label along with the view, by pixels
 */
class PinnedLabelScroll {
    constructor(drawing) {
        this.drawing = drawing;
        this.coordinates = new Point();
    }

    get isMovable() {
        return true;
    }

    allowMove() {
        return true;
    }

    moveTo(position) {
        const pixels = this.drawing.coordinateSystem.toPhysical(position).minus(this.drawing.coordinateSystem.toPhysical(this.coordinates));
        for (const figure of this.drawing.figures.list) {
            if (figure instanceof Label && figure.pin !== LabelPin.None) {
                figure.scrollPinned(pixels);
            }
        }
    }
}
