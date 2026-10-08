// The player: a canvas in an element, a drawing on it, the Drag tool. What the app's
// DrawingControl and MainView are for the editor, as far as the player needs them. The
// PlayerCanvas is what the figures and the coordinate system call the canvas: its size in
// CSS pixels, the text measurer, the pointer events' element, and invalidate, which asks
// for a frame.

class PlayerCanvas {
    constructor(element, measurer) {
        this.element = element;
        this.measurer = measurer;
        this.width = 0;
        this.height = 0;
        this.invalidated = null;
        this.wheelZooms = false;
    }

    /** The canvas's borders in pixels (GetBorderRectangle) */
    getBorderRectangle() {
        return new PointPair(new Point(0, 0), new Point(this.width, this.height));
    }

    measureText(text, font, wrapWidth) {
        return this.measurer.layout(text, font, wrapWidth);
    }

    /** The event's place on the canvas, in CSS pixels */
    positionOf(e) {
        const rect = this.element.getBoundingClientRect();
        return new Point(e.clientX - rect.left, e.clientY - rect.top);
    }

    setCursor(cursor) {
        const value = cursor ?? "";
        if (this.element.style.cursor !== value) {
            this.element.style.cursor = value;
        }
    }

    /** Something changed: the next frame draws it */
    invalidate() {
        this.invalidated?.();
    }
}

class Player {
    /**
     * options: lgf (the drawing's text), src (a URL to fetch it from), theme ("light",
     * "dark" or "auto"), font (a CSS font family), wheel ("zoom" lets the wheel zoom without
     * Ctrl), fit ("content" fits what is drawn rather than the file's viewport)
     */
    constructor(element, options = {}) {
        this.element = element;
        this.options = options;
        this.canvasElement = document.createElement("canvas");
        this.canvasElement.style.display = "block";
        this.canvasElement.style.width = "100%";
        this.canvasElement.style.height = "100%";
        this.canvasElement.style.touchAction = "none";
        this.canvasElement.style.userSelect = "none";
        this.canvasElement.setAttribute("tabindex", "-1");
        element.appendChild(this.canvasElement);
        this.context = this.canvasElement.getContext("2d");
        this.measurer = new TextMeasurer();
        this.renderer = new CanvasRenderer(this.context);
        this.canvas = new PlayerCanvas(this.canvasElement, this.measurer);
        this.canvas.wheelZooms = options.wheel === "zoom";
        this.canvas.invalidated = () => this.invalidate();
        this.frameRequested = false;
        this.drawing = null;
        this.behavior = null;
        this.pixelRatio = 1;
        this.applyTheme(options.theme ?? "auto");
        this.applyFont(options.font);
        this.measureElement();
        this.observer = typeof ResizeObserver !== "undefined" ? new ResizeObserver(() => this.resize()) : null;
        this.observer?.observe(element);
        if (options.lgf != null) {
            this.load(options.lgf);
        } else if (options.src != null) {
            this.loadFrom(options.src);
        }
    }

    /** The theme the drawing is shown under: the host page's choice, or the system's */
    applyTheme(theme) {
        this.themeChoice = theme;
        let dark = theme === "dark";
        if (theme === "auto" && typeof matchMedia === "function") {
            const query = matchMedia("(prefers-color-scheme: dark)");
            dark = query.matches;
            if (this.themeQuery == null) {
                this.themeQuery = query;
                query.addEventListener?.("change", () => {
                    if (this.themeChoice === "auto") {
                        this.applyTheme("auto");
                    }
                });
            }
        }

        // the drawing's own theme, not a global: two embeds on a page may differ
        this.theme = dark ? AppTheme.Dark : AppTheme.Light;
        if (this.drawing != null) {
            this.drawing.theme = this.theme;
            this.drawing.refreshTheme();
        }

        this.invalidate();
    }

    /** The font the text is drawn in: the one given, else the host element's */
    applyFont(font) {
        if (font != null && font !== "") {
            Fonts.family = font;
        } else if (typeof getComputedStyle === "function") {
            const family = getComputedStyle(this.element).fontFamily;
            if (family != null && family !== "") {
                Fonts.family = family;
            }
        }

        this.measurer.fontMetrics.clear();
    }

    measureElement() {
        const rect = this.element.getBoundingClientRect();
        const width = Math.max(0, Math.floor(rect.width));
        const height = Math.max(0, Math.floor(rect.height));
        const pixelRatio = window.devicePixelRatio || 1;
        if (width === this.canvas.width && height === this.canvas.height && pixelRatio === this.pixelRatio) {
            return false;
        }

        const previousWidth = this.canvas.width;
        const previousHeight = this.canvas.height;
        this.canvas.width = width;
        this.canvas.height = height;
        this.pixelRatio = pixelRatio;
        this.canvasElement.width = Math.max(1, Math.round(width * pixelRatio));
        this.canvasElement.height = Math.max(1, Math.round(height * pixelRatio));
        this.drawing?.onSizeChanged(previousWidth, previousHeight, width, height);
        return true;
    }

    /** The element changed its size: the canvas follows, and the drawing is fitted again while it is still as it opened */
    resize() {
        if (this.measureElement()) {
            if (this.keepFitted) {
                this.fitInitial();
            }

            this.invalidate();
        }
    }

    /** The drawing from its text; what could not be read is in loadErrors */
    load(lgfText) {
        const drawing = new Drawing(this.canvas);
        drawing.theme = this.theme;
        drawing.addFromXml(lgfText);
        this.loadErrors = drawing.loadErrors;
        if (this.loadErrors != null) {
            console.warn("Live Geometry: not all of the drawing could be read.\n" + this.loadErrors);
        }

        this.show(drawing);
    }

    /** The drawing from a URL (an .lgf the page can fetch: same origin, or one that allows it) */
    async loadFrom(url) {
        const response = await fetch(url);
        if (!response.ok) {
            throw new Error("Live Geometry: could not load " + url + " (" + response.status + ")");
        }

        this.load(await response.text());
    }

    show(drawing) {
        if (this.drawing != null) {
            this.drawing.behavior = null;
            this.drawing.canvas = null;
        }

        this.drawing = drawing;
        drawing.canvas = this.canvas;
        this.behavior = new Dragger();
        drawing.behavior = this.behavior;

        // the view the drawing opened with is kept through resizes until the user moves
        // something or the view (as the app keeps a drawing fitted, MainView.KeepFitted)
        this.keepFitted = true;
        drawing.viewChanged = () => this.keepFitted = false;
        drawing.figuresMoved = () => this.keepFitted = false;
        this.fitInitial();
        this.invalidate();
    }

    /** The view the drawing opens with: the file's viewport fitted into the canvas, or the content when asked */
    fitInitial() {
        const drawing = this.drawing;
        if (drawing == null) {
            return;
        }

        if (this.options.fit === "content" || drawing.viewport == null) {
            drawing.zoomToFit();
        } else {
            const viewport = drawing.viewport;
            drawing.coordinateSystem.setViewport(viewport.x, viewport.right, viewport.y, viewport.bottom);
        }
    }

    /** Zoom to fit, as a double click does */
    fit() {
        this.keepFitted = false;
        this.drawing?.zoomToFit();
        this.invalidate();
    }

    invalidate() {
        if (this.frameRequested) {
            return;
        }

        this.frameRequested = true;
        requestAnimationFrame(() => {
            this.frameRequested = false;
            this.render();
        });
    }

    render() {
        const drawing = this.drawing;
        this.renderer.beginFrame(this.canvas.width, this.canvas.height, this.pixelRatio, drawing != null && drawing.paintsPaper ? drawing.paperForCanvas : null);
        if (drawing != null) {
            drawing.render(this.renderer);
        }

        this.renderer.endFrame();
    }

    /** Every figure's numbers, for the parity check */
    dump() {
        return this.drawing?.dump() ?? [];
    }

    dispose() {
        this.observer?.disconnect();
        if (this.drawing != null) {
            this.drawing.behavior = null;
            this.drawing.canvas = null;
        }

        this.canvasElement.remove();
    }
}
