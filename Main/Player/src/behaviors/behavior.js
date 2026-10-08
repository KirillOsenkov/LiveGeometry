// Port of Main/Avalonia/DynamicGeometry/Behaviors/Behavior.cs: a tool. The pointer events of
// the canvas element come in here and go to mouseDown/mouseMove/mouseUp in logical
// coordinates, as the WPF-shaped virtuals the C# tools override. Fingers are told apart from
// a mouse as in the app: a finger's press is kept back until it is a tap, a drag or one of
// two fingers, which zoom and pan the view. Left out: the click preview and the choice among
// overlapping figures (tools that construct), the keyboard, the context menu.

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
        this.fingers = [];
        this.isPinching = false;
        this.touchPress = null;
        this.touchPressDelivered = false;
        this.touchLast = null;
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

    forgetTouch() {
        this.touchPress = null;
        this.touchPressDelivered = false;
        this.touchLast = null;
    }

    touchPressed(e) {
        this.parentCanvas.element.setPointerCapture?.(e.pointerId);
        e.preventDefault();
        this.fingers = this.fingers.filter(finger => finger.pointerId !== e.pointerId);
        this.fingers.push({ pointerId: e.pointerId, position: this.position(e) });
        if (this.fingers.length === 1) {
            this.isPinching = false;
            this.touchPress = e;
            this.touchPressDelivered = false;
            this.touchLast = e;
            return;
        }

        if (!this.isPinching) {
            this.isPinching = true;

            // a drag under way ends where the first finger is
            if (this.touchPressDelivered && this.touchLast != null) {
                const last = this.touchLast;
                this.asTouch(() => this.safeMouseUp(last));
            }

            this.forgetTouch();
        }
    }

    touchMoved(e) {
        const index = this.fingers.findIndex(finger => finger.pointerId === e.pointerId);
        if (index < 0) {
            return;
        }

        const position = this.position(e);
        if (this.isPinching) {
            if (index < 2 && this.fingers.length >= 2 && this.drawing != null) {
                const from = GeometryMath.midpoint(this.fingers[0].position, this.fingers[1].position);
                const span = this.fingers[0].position.distance(this.fingers[1].position);
                this.fingers[index].position = position;
                const to = GeometryMath.midpoint(this.fingers[0].position, this.fingers[1].position);
                const newSpan = this.fingers[0].position.distance(this.fingers[1].position);

                // fingers almost on each other say nothing about the zoom
                const factor = span > Behavior.TouchSlop && newSpan > Behavior.TouchSlop ? newSpan / span : 1;
                this.drawing.coordinateSystem.panAndZoom(from, to, factor);
            } else {
                this.fingers[index].position = position;
            }

            return;
        }

        this.fingers[index].position = position;
        if (this.touchPress == null) {
            return;
        }

        this.touchLast = e;
        if (!this.touchPressDelivered) {
            if (position.distance(this.position(this.touchPress)) < Behavior.TouchSlop) {
                return;
            }

            this.touchPressDelivered = true;
            const press = this.touchPress;
            this.asTouch(() => this.safeMouseDown(press));
        }

        this.asTouch(() => this.safeMouseMove(e));
    }

    /** e: the release; null when the touch was taken away, and the finger is where it was last seen */
    touchReleased(pointerId, e) {
        const index = this.fingers.findIndex(finger => finger.pointerId === pointerId);
        if (index < 0) {
            return;
        }

        this.fingers.splice(index, 1);
        if (this.isPinching) {
            if (this.fingers.length === 0) {
                this.isPinching = false;
            }

            return;
        }

        if (this.touchPress != null) {
            if (e != null && !this.touchPressDelivered) {
                // a tap: the press and the release in one go
                this.touchPressDelivered = true;
                const press = this.touchPress;
                this.asTouch(() => this.safeMouseDown(press));
            }

            const release = e ?? this.touchLast;
            if (this.touchPressDelivered && release != null) {
                this.asTouch(() => this.safeMouseUp(release));
            }
        }

        this.forgetTouch();
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
