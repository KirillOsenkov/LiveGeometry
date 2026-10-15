// Port of Main/Avalonia/LiveGeometry/Behaviors/Behavior.cs: a tool. The pointer events of
// the canvas element come in here and go to mouseDown/mouseMove/mouseUp in logical
// coordinates, as the WPF-shaped virtuals the C# tools override. Fingers are told apart from
// a mouse as in the app: a finger's press is kept back until it is a tap or a drag, and each
// finger is for a figure (it drags that one, whatever the other fingers do) or for the paper
// (the paper fingers pan and zoom the view together). Left out: the click preview and the
// choice among overlapping figures (tools that construct), the keyboard, the context menu.

class Behavior {
    /** In pixels: how far a finger goes from where it came down before it is dragging and not tapping */
    static TouchSlop = 6;

    /** In pixels: what a finger reaches, where the mouse has Settings.cursorTolerance */
    static TouchTolerance = 10;

    static WheelZoomFactor = 1.2;

    constructor() {
        this.drawingValue = null;
        this.parentCanvasValue = null;
        this.errorHappened = false;
        this.modifiers = { shift: false, ctrl: false, alt: false };

        // the fingers on the canvas in the order they came down: { pointerId, position (canvas
        // pixels, where it was last seen), onFigure, press (kept back until the finger is a tap
        // or a drag; null once the finger moves the view itself), pressDelivered, last,
        // drivesView (a paper finger moving the view itself, past the tool) }
        this.fingers = [];
        this.listeners = null;
    }

    get name() {
        return "";
    }

    get hintText() {
        return "";
    }

    get drawing() {
        return this.drawingValue;
    }

    set drawing(value) {
        if (this.drawingValue != null) {
            this.parentCanvas = null;
        }

        this.drawingValue = value;
        if (value != null) {
            this.parentCanvas = value.canvas;
        }
    }

    /** The canvas the tool listens to (the player's surface, whose element gets the pointer events) */
    get parentCanvas() {
        return this.parentCanvasValue;
    }

    set parentCanvas(value) {
        if (this.parentCanvasValue != null && this.listeners != null) {
            const element = this.parentCanvasValue.element;
            for (const [type, listener, options] of this.listeners) {
                element.removeEventListener(type, listener, options);
            }

            this.listeners = null;
            this.parentCanvasValue.setCursor(null);
        }

        this.parentCanvasValue = value;
        if (value != null) {
            const element = value.element;
            this.listeners = [
                ["pointerdown", e => this.pointerPressedHandler(e), undefined],
                ["pointermove", e => this.pointerMovedHandler(e), undefined],
                ["pointerup", e => this.pointerReleasedHandler(e), undefined],
                ["pointercancel", e => this.pointerCaptureLostHandler(e), undefined],
                ["pointerleave", e => this.pointerExitedHandler(e), undefined],
                ["wheel", e => this.pointerWheelHandler(e), { passive: false }],
                ["contextmenu", e => e.preventDefault(), undefined]
            ];
            for (const [type, listener, options] of this.listeners) {
                element.addEventListener(type, listener, options);
            }
        }
    }

    started() {
    }

    stopping() {
    }

    isCtrlPressed() {
        return this.modifiers.ctrl;
    }

    isShiftPressed() {
        return this.modifiers.shift;
    }

    isAltPressed() {
        return this.modifiers.alt;
    }

    readModifiers(e) {
        this.modifiers = { shift: e.shiftKey, ctrl: e.ctrlKey || e.metaKey, alt: e.altKey };
    }

    /** The event's place on the canvas, in pixels (GetPosition) */
    position(e) {
        return this.parentCanvas.positionOf(e);
    }

    pointerPressedHandler(e) {
        this.readModifiers(e);
        if (e.pointerType === "touch") {
            this.touchPressed(e);
            return;
        }

        if (e.button === 0) {
            // the canvas keeps the pointer while the button is down, also outside the element
            this.parentCanvas.element.setPointerCapture?.(e.pointerId);
            e.preventDefault();
            this.safeMouseDown(e);
        } else if (e.button === 2) {
            try {
                this.mouseRightClick(e);
            } catch (error) {
                this.handleException(error);
            }
        }
    }

    pointerMovedHandler(e) {
        this.readModifiers(e);
        if (e.pointerType === "touch") {
            this.touchMoved(e);
            return;
        }

        this.handleMove(e);
    }

    handleMove(e) {
        this.safeMouseMove(e);
        if (!this.errorHappened && (e.buttons & 1) === 0) {
            this.updateCursor(e);
        }
    }

    pointerExitedHandler(e) {
    }

    // Cursor

    updateCursor(e) {
        if (this.parentCanvas == null || this.drawing == null) {
            return;
        }

        let cursor;
        try {
            cursor = this.getCursor(this.coordinatesOf(e, false, false, false));
        } catch (error) {
            cursor = "default";
        }

        this.parentCanvas.setCursor(cursor);
    }

    /**
     * The cursor tells what a click at this place would do: a cross, a new free point; a
     * hand, the click picks something that is already there; an arrow, everything else.
     */
    getCursor(coordinates) {
        return "default";
    }

    // Touch

    /** A tool handles what a finger did with a finger's reach */
    asTouch(action) {
        const tolerance = GeometryMath.cursorTolerance;
        GeometryMath.cursorTolerance = Behavior.TouchTolerance;
        try {
            action();
        } finally {
            GeometryMath.cursorTolerance = tolerance;
        }
    }

    /**
     * Whether a finger pressed here is for a figure - it taps or drags that figure, whatever
     * other fingers do meanwhile - rather than for the paper, which the paper fingers pan
     * and zoom together: one pans, two pan and zoom (the paper under them stays under
     * them), a third does nothing. The paper fingers are the tool's only while there is one
     * finger on the canvas; with any other they move the view themselves. False by default.
     */
    takesFingerOnFigure(e) {
        return false;
    }

    safeTakesFingerOnFigure(e) {
        if (this.errorHappened || this.parentCanvas == null || this.drawing == null) {
            return false;
        }

        try {
            return this.takesFingerOnFigure(e);
        } catch (error) {
            return false;
        }
    }

    findFinger(pointerId) {
        return this.fingers.find(finger => finger.pointerId === pointerId) ?? null;
    }

    touchPressed(e) {
        this.parentCanvas.element.setPointerCapture?.(e.pointerId);
        e.preventDefault();
        this.fingers = this.fingers.filter(finger => finger.pointerId !== e.pointerId);
        const finger = {
            pointerId: e.pointerId,
            position: this.position(e),
            onFigure: false,
            press: e,
            pressDelivered: false,
            last: e,
            drivesView: false
        };
        this.asTouch(() => finger.onFigure = this.safeTakesFingerOnFigure(e));
        this.fingers.push(finger);

        // the paper fingers are the tool's only while there is one finger
        if (this.fingers.length > 1) {
            for (const paper of this.fingers.filter(other => !other.onFigure && !other.drivesView)) {
                this.takeFromTool(paper);
            }
        }

        // A figure finger's press goes to the tool at once (a press on a figure does nothing
        // visible under the Drag tool, which waits for the slop itself). Held back, the press
        // reached the tool after the paper fingers had moved the view, and found the figure
        // no longer under the finger.
        if (finger.onFigure) {
            finger.pressDelivered = true;
            this.asTouch(() => this.safeMouseDown(e));
        }
    }

    /** A paper finger goes over from the tool to the view: a drag under way ends where the finger is */
    takeFromTool(finger) {
        if (finger.pressDelivered && finger.last != null) {
            const last = finger.last;
            this.asTouch(() => this.safeMouseUp(last));
        }

        finger.press = null;
        finger.pressDelivered = false;
        finger.last = null;
        finger.drivesView = true;
    }

    touchMoved(e) {
        const finger = this.findFinger(e.pointerId);
        if (finger == null) {
            return;
        }

        const position = this.position(e);
        if (finger.drivesView) {
            this.moveView(finger, position);
            return;
        }

        finger.position = position;
        if (finger.press == null) {
            return;
        }

        finger.last = e;
        if (!finger.pressDelivered) {
            if (position.distance(this.position(finger.press)) < Behavior.TouchSlop) {
                return;
            }

            finger.pressDelivered = true;
            const press = finger.press;
            this.asTouch(() => this.safeMouseDown(press));
        }

        this.asTouch(() => this.safeMouseMove(e));
    }

    /**
     * The paper fingers move the view: what was under the first two of them stays under
     * them. The figures held by the other fingers stay under those: each gets its last move
     * again over the moved paper.
     */
    moveView(finger, position) {
        const driving = this.fingers.filter(other => other.drivesView).slice(0, 2);
        const index = driving.indexOf(finger);
        if (index < 0 || this.drawing == null) {
            finger.position = position;
            return;
        }

        let from;
        let to;
        let factor = 1;
        if (driving.length === 1) {
            from = finger.position;
            finger.position = position;
            to = position;
        } else {
            from = GeometryMath.midpoint(driving[0].position, driving[1].position);
            const span = driving[0].position.distance(driving[1].position);
            finger.position = position;
            to = GeometryMath.midpoint(driving[0].position, driving[1].position);
            const newSpan = driving[0].position.distance(driving[1].position);

            // fingers almost on each other say nothing about the zoom
            factor = span > Behavior.TouchSlop && newSpan > Behavior.TouchSlop ? newSpan / span : 1;
        }

        this.drawing.coordinateSystem.panAndZoom(from, to, factor);
        for (const held of this.fingers.filter(other => other.pressDelivered && other.last != null)) {
            const last = held.last;
            this.asTouch(() => this.safeMouseMove(last));
        }
    }

    /** e: the release; null when the touch was taken away, and the finger is where it was last seen */
    touchReleased(pointerId, e) {
        const finger = this.findFinger(pointerId);
        if (finger == null) {
            return;
        }

        this.fingers.splice(this.fingers.indexOf(finger), 1);
        if (finger.press == null) {
            return;
        }

        if (e != null && !finger.pressDelivered) {
            // a tap: the press and the release in one go
            finger.pressDelivered = true;
            const press = finger.press;
            this.asTouch(() => this.safeMouseDown(press));
        }

        const release = e ?? finger.last;
        if (finger.pressDelivered && release != null) {
            this.asTouch(() => this.safeMouseUp(release));
        }
    }

    pointerCaptureLostHandler(e) {
        if (e.pointerType === "touch") {
            this.touchReleased(e.pointerId, null);
        }
    }

    pointerReleasedHandler(e) {
        this.readModifiers(e);
        if (e.pointerType === "touch") {
            this.touchReleased(e.pointerId, e);
            return;
        }

        if (e.button === 0) {
            this.safeMouseUp(e);
        }
    }

    pointerWheelHandler(e) {
        this.readModifiers(e);
        this.mouseWheel(e);
    }

    mouseDown(e) {
    }

    mouseMove(e) {
    }

    mouseUp(e) {
    }

    mouseRightClick(e) {
    }

    handleException(error) {
        console.error(error);
    }

    safeMouseDown(e) {
        try {
            this.mouseDown(e);
        } catch (error) {
            this.errorHappened = true;
            this.handleException(error);
        }
    }

    safeMouseMove(e) {
        if (this.errorHappened) {
            return;
        }

        try {
            this.mouseMove(e);
        } catch (error) {
            this.handleException(error);
            this.errorHappened = true;
        }
    }

    safeMouseUp(e) {
        this.errorHappened = false;
        try {
            this.mouseUp(e);
        } catch (error) {
            this.handleException(error);
        }
    }

    /**
     * The wheel zooms around the cursor, a notch a factor of 1.2 - only with Ctrl (or Cmd)
     * held, unless the player says otherwise (wheelZooms): in a page an embed that zooms on
     * every wheel hijacks the scrolling of the text around it.
     */
    mouseWheel(e) {
        if (this.drawing == null) {
            return;
        }

        if (!this.parentCanvas.wheelZooms && !(e.ctrlKey || e.metaKey)) {
            return;
        }

        e.preventDefault();
        const notches = e.deltaMode === 1 ? -e.deltaY / 3 : e.deltaMode === 2 ? -e.deltaY : -e.deltaY / 100;
        if (notches !== 0) {
            const factor = Math.pow(Behavior.WheelZoomFactor, notches);
            this.drawing.coordinateSystem.zoom(factor, this.position(e));
        }
    }

    // Coordinates

    /** The event's place in the plane; with Shift, snapped to the labeled grid lines */
    coordinates(e) {
        if (this.isShiftPressed()) {
            return this.coordinatesOf(e, false, true, false);
        }

        return this.coordinatesOf(e, false, false, false);
    }

    coordinatesOf(e, snapToPoint, snapToGrid, snapToCenter) {
        let result = this.toLogical(this.position(e));
        if (snapToGrid) {
            result = GeometryMath.getSnapToGridPosition(this.drawing.coordinateSystem.majorGridStep, result);
        }

        return result;
    }

    get cursorTolerance() {
        return this.drawing.coordinateSystem.cursorTolerance;
    }

    toPhysical(point) {
        return this.drawing.coordinateSystem.toPhysical(point);
    }

    toLogical(pixel) {
        return this.drawing.coordinateSystem.toLogical(pixel);
    }
}
