// Port of Main/Avalonia/LiveGeometry/Behaviors/Dragger.cs: the Drag tool, the player's one
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

        // The presses held right now, in the order they came, one per pointer (the mouse is
        // one, each finger on a figure is one, see "Touch" in Behavior: two fingers drag the
        // two ends of a segment at once): { pointerId, moving, found, pressedFixedLabel (a
        // label the press found that can't be dragged: the drag moves the view, a click
        // ticks a box), toRecalculate, offsetFromFigureLeftTopCorner, oldCoordinates,
        // coordinatesOnMouseDown, startedMoving }
        this.drags = [];
        this.clickCount = 0;
        this.lastPressTime = 0;
        this.lastPressPosition = null;
    }

    findDrag(pointerId) {
        return this.drags.find(drag => drag.pointerId === pointerId) ?? null;
    }

    /**
     * What a press at the coordinates takes: the figure a drag moves or a click selects, or
     * null for the paper, which a drag moves - an axis (the grid's, for the tools to build
     * on), and the labels that can't be dragged (drawing.fixedLabels); a drag that starts on
     * a caption takes the captions along (isCaption). A click still ticks such a box
     * (fixedLabel; a tap leaves a caption alone)
     */
    findPressedFigure(coordinates, isTouch) {
        let found = this.drawing.figures.hitTest(coordinates);

        // an axis is the grid's, for the tools to build on: a press on it is a press on the paper
        if (found != null && found.isAxisLine === true) {
            found = null;
        }

        let fixedLabel = null;
        let isCaption = false;
        if (found instanceof ControlBase && this.drawing.fixedLabels.has(found)) {
            isCaption = found instanceof Label && found.pin !== LabelPin.None;
            if (!isCaption || !isTouch) {
                fixedLabel = found;
            }

            found = null;
        }

        return { found, fixedLabel, isCaption };
    }

    /**
     * A finger stays with what its press would drag or select, whatever the other fingers
     * do; one whose press would move the view (the paper, an axis, a caption, a graph with
     * nothing to move) pans and zooms with the other paper fingers
     */
    takesFingerOnFigure(e) {
        return !this.prepareDrag(e).moving.includes(this.drawing.coordinateSystem);
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
        // a drag of this pointer whose release never arrived
        this.release(this.findDrag(e.pointerId));

        // (not on a show/hide box, which two quick clicks tick and untick)
        if (this.countClicks(e) === 2 && !(this.drawing.figures.hitTest(this.coordinatesOf(e, false, false, false)) instanceof ShowHideControl)) {
            this.drawing.zoomToFit();
            this.clickCount = 0;
            return;
        }

        const drag = this.prepareDrag(e);

        // the handles next to an anchor of a Bezier path show from the press on it until the
        // next press elsewhere (there is no selection to keep them shown by)
        BezierPath.showHandlesWhileDragging(this.drawing, drag.found);
        this.drags.push(drag);
    }

    /** What a press here would hold: the figure under it and what a drag moves (nothing is changed) */
    prepareDrag(e) {
        const pressed = this.coordinatesOf(e, false, false, false);
        const drag = {
            pointerId: e.pointerId,
            moving: [],
            found: null,
            pressedFixedLabel: null,
            toRecalculate: null,
            offsetFromFigureLeftTopCorner: pressed,
            oldCoordinates: pressed,
            coordinatesOnMouseDown: pressed,
            startedMoving: false
        };

        let roots = null;
        let isLocked = false;

        const { found, fixedLabel, isCaption } = this.findPressedFigure(pressed, e.pointerType === "touch");
        drag.found = found;
        drag.pressedFixedLabel = fixedLabel;
        const captions = isCaption ? new PinnedLabelScroll(this.drawing) : null;

        // a figure with parts (a slider) says which of them the press takes
        const oneMovable = found != null && found.isMovableParts === true
            ? found.findMovablePart(pressed)
            : found != null && found.isMovable === true ? found : null;
        if (oneMovable != null && (found.locked || oneMovable.allowMove())) {
            if (found.locked) {
                isLocked = true;
            } else if (oneMovable.allowMove()) {
                if (oneMovable.isPoint === true) {
                    // when we drag a point, we want it to snap to the cursor
                    drag.offsetFromFigureLeftTopCorner = new Point();
                    drag.oldCoordinates = oneMovable.coordinates;
                } else {
                    // other stuff (text labels) keeps the grab offset: no snap to the cursor
                    drag.offsetFromFigureLeftTopCorner = pressed.minus(oneMovable.coordinates);
                }

                roots = DependencyAlgorithms.findRoots(f => f.dependents, [found]);

                // (a show/hide box holds figures, it isn't built on them: locked, it stays
                // where it is, and they still move)
                if (roots.every(root => !root.locked || root instanceof ShowHideControl)) {
                    drag.moving.push(oneMovable);
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
                        drag.moving.push(...roots);
                    } else {
                        isLocked = true;
                    }
                }
            } else {
                isLocked = true;
            }
        }

        if (roots != null) {
            drag.toRecalculate = DependencyAlgorithms.findDescendants(f => f.dependents, roots);
            drag.toRecalculate.reverse();
        } else {
            drag.toRecalculate = null;
        }

        if (drag.moving.length === 0 && !isLocked && !this.drawing.coordinateGrid.locked) {
            drag.moving.push(this.drawing.coordinateSystem);
            if (captions != null) {
                drag.moving.push(captions);
            }

            drag.toRecalculate = null;
        }

        return drag;
    }

    mouseMove(e) {
        const drag = this.findDrag(e.pointerId);
        if (drag == null) {
            return;
        }

        // a move with no button down is no drag: the release of this press never arrived
        if (e.pointerType !== "touch" && (e.buttons & 1) === 0) {
            this.release(drag);
            return;
        }

        const moving = drag.moving;
        const currentCoordinates = this.coordinates(e);
        if (!drag.startedMoving) {
            if (currentCoordinates.equals(drag.coordinatesOnMouseDown)) {
                return;
            }

            // a press that wobbles is still a click
            const coordinateSystem = this.drawing.coordinateSystem;
            const wobble = coordinateSystem.toPhysical(this.coordinatesOf(e, false, false, false))
                .distance(coordinateSystem.toPhysical(drag.coordinatesOnMouseDown));
            if (wobble < (e.pointerType === "touch" ? Behavior.TouchSlop : Dragger.DragThreshold)) {
                return;
            }

            drag.startedMoving = true;
        }

        if (moving.length > 0) {
            let offset = currentCoordinates.minus(drag.oldCoordinates);
            if (moving.length === 1 && moving[0] instanceof PointLabel) {
                // a point label is confined to an orbit around its point: move it to where the
                // cursor wants it (as far as allowed)
                const pointLabel = moving[0];
                const desired = currentCoordinates.minus(drag.offsetFromFigureLeftTopCorner);
                offset = pointLabel.clampPosition(desired).minus(pointLabel.coordinates);
            } else if (moving.length === 1 && moving[0] instanceof PointOnFigure) {
                // a point on a figure stops at the end of a segment or ray: where it lands, not where the cursor went
                const pointOnFigure = moving[0];
                const figure = pointOnFigure.linearFigure;
                const landing = figure.getPointFromParameter(
                    figure.getNearestParameterFromPoint(pointOnFigure.coordinates.plus(offset)));
                offset = landing.minus(pointOnFigure.coordinates);
            } else if (moving.length === 1 && moving[0] instanceof PointInRegion && moving[0].region != null) {
                // and for a point kept inside a shape, which stops at its outline
                const inside = moving[0];
                const landing = inside.region.nearestPointInside(inside.coordinates.plus(offset));
                if (landing.exists()) {
                    offset = landing.minus(inside.coordinates);
                }
            }

            Actions.move(this.drawing, moving, offset, drag.toRecalculate);
        }

        // if you're dragging the coordinate plane itself, the origin changes, so the point's
        // coordinates have to be read again in the new coordinate system
        drag.oldCoordinates = this.coordinates(e);
        if (moving.length === 1
            && moving[0].isPoint === true
            && drag.found != null
            && (drag.found === moving[0] || drag.found.isMovableParts === true)) {
            drag.oldCoordinates = moving[0].coordinates;
        }
    }

    mouseUp(e) {
        const drag = this.findDrag(e.pointerId);
        if (drag == null) {
            return;
        }

        // a press that didn't become a drag: on a show/hide box it ticks the box
        if (!drag.startedMoving) {
            const box = drag.found ?? drag.pressedFixedLabel;
            if (box instanceof ShowHideControl && !this.isCtrlPressed()) {
                box.click();
                this.drawing.canvas?.invalidate?.();
            }
        }

        this.release(drag);
    }

    stopping() {
        this.releaseAll();
        super.stopping();
    }

    get drawing() {
        return super.drawing;
    }

    set drawing(value) {
        if (value !== super.drawing) {
            this.releaseAll();
        }

        super.drawing = value;
    }

    /** The press is over: nothing is held by it any more */
    release(drag) {
        if (drag == null) {
            return;
        }

        const index = this.drags.indexOf(drag);
        if (index >= 0) {
            this.drags.splice(index, 1);
        }
    }

    releaseAll() {
        for (const drag of this.drags.slice()) {
            this.release(drag);
        }
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
