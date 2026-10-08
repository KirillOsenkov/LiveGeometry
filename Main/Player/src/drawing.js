// Port of Main/Avalonia/DynamicGeometry/Drawing.cs: the figures, their styles, the view and
// the paper. Left out: undo (ActionManager), selection, copy and paste, the events of the
// editor, the theme refresh of a parked drawing.

class Drawing {
    constructor(canvas) {
        this.canvasValue = null;
        this.styleManager = new StyleManager(this);
        this.figures = new RootFigureList(this);
        this.axisLines = [null, null];
        this.backgroundValue = null;

        /** By theme name, the paper the drawing chose for that theme (null for the theme's paper) */
        this.overrides = new Map();

        /** Suggested views, in logical coordinates, for drawings whose content has no useful bounds */
        this.scenes = [];
        this.activeScene = null;

        /** "Zoom to fit" as whoever shows the drawing does it; null for the plain fit */
        this.fitToWindow = null;

        /** The view the file asks for (its Viewport element), in logical coordinates; null without one */
        this.viewport = null;

        /** Called when the view is moved or zoomed by the user (a pan, the wheel, two fingers) */
        this.viewChanged = null;

        /** Called when a figure is moved by the user */
        this.figuresMoved = null;

        /** Labels that can't be dragged: a drag on one moves the view */
        this.fixedLabels = new Set();
        this.isReading = false;
        this.isMoving = false;
        this.name = null;
        this.loadErrors = null;
        this.behaviorValue = null;

        /** The theme the drawing is shown under (the player's choice: two embeds on a page may differ) */
        this.theme = AppTheme.current;
        this.canvas = canvas;
        this.coordinateSystem = new CoordinateSystem(this);

        // the grid is the drawing's own: a new drawing starts without one, a file says
        this.coordinateGrid = new CartesianGrid();
        this.coordinateGrid.drawing = this;
        this.coordinateGrid.visible = false;
        this.figures.add(this.coordinateGrid);
        this.version = Settings.currentDrawingVersion;
    }

    /** The drawing's x- or y-axis as a line to build on: the same object for the drawing's whole life */
    getAxisLine(direction) {
        const index = direction === AxisDirection.Y ? 1 : 0;
        if (this.axisLines[index] == null) {
            const axis = new AxisLine(direction);
            axis.drawing = this;
            this.axisLines[index] = axis;
        }

        return this.axisLines[index];
    }

    /** The axis lines a click can take that aren't in the list: hit testing looks at them too */
    unlistedAxisLines() {
        if (this.coordinateGrid == null || !this.coordinateGrid.showsAxes) {
            return [];
        }

        const result = [];
        for (const direction of [AxisDirection.X, AxisDirection.Y]) {
            const axis = this.getAxisLine(direction);
            if (!this.figures.contains(axis)) {
                result.push(axis);
            }
        }

        return result;
    }

    /** The paper: a solid color or a gradient of the drawing's own, or the theme's paper */
    get background() {
        return this.getOwnBackground(this.theme) ?? new SolidColorBrush(this.theme.paper);
    }

    set background(value) {
        this.backgroundValue = value;
        this.applyBackground();
    }

    /** The paper the drawing chose (under the base theme), null for the theme's */
    get ownBackground() {
        return this.backgroundValue;
    }

    /** The paper the drawing has under the theme, null for the theme's own */
    getOwnBackground(theme) {
        const values = this.overrides.get(theme.name);
        if (values != null && values.has("background")) {
            return values.get("background");
        }

        return this.backgroundValue;
    }

    setOverride(theme, property, value) {
        let values = this.overrides.get(theme);
        if (values == null) {
            values = new Map();
            this.overrides.set(theme, values);
        }

        values.set(property, value);
        this.applyBackground();
    }

    /** Whether the paper is painted onto the canvas at all */
    paintsPaper = true;

    applyBackground() {
        this.coordinateGrid?.applyStyle();
        this.canvas?.invalidate?.();
    }

    /** The theme on screen changed: every figure draws again as the theme now says */
    refreshTheme() {
        if (this.canvas == null) {
            return;
        }

        this.applyBackground();
        for (const figure of this.figures.list) {
            figure.applyStyle();
        }
    }

    /** The scene nearest in shape to a room of this size; null without scenes */
    chooseScene(roomWidth, roomHeight) {
        if (this.scenes.length === 0 || roomWidth <= 0 || roomHeight <= 0) {
            return null;
        }

        const room = Math.log(roomWidth / roomHeight);
        return [...this.scenes].sort((a, b) => Math.abs(Math.log(a.width / a.height) - room) - Math.abs(Math.log(b.width / b.height) - room))[0];
    }

    /** Fits the scene into the canvas edge to edge and makes it the active one */
    showScene(scene) {
        this.activeScene = scene;
        this.coordinateSystem.fitScene(scene);
    }

    /** Zoom to fit, as the user asks for it: the layout the drawing was opened with, a scene if it has scenes, else everything visible */
    zoomToFit() {
        if (this.fitToWindow != null) {
            this.fitToWindow();
            return;
        }

        const scene = this.canvas != null ? this.chooseScene(this.canvas.width, this.canvas.height) : null;
        if (scene != null) {
            this.showScene(scene);
        } else {
            this.coordinateSystem.zoomExtend();
        }
    }

    /**
     * The paper as a renderer paints it: a gradient's relative start and end are relative to
     * the active scene, not to the canvas (placeBackground)
     */
    get paperForCanvas() {
        const brush = this.background;
        if (this.activeScene == null || this.coordinateSystem == null || !(brush instanceof LinearGradientBrush)) {
            return brush;
        }

        const scene = this.activeScene;
        const topLeft = this.coordinateSystem.toPhysical(new Point(scene.x, scene.bottom));
        const bottomRight = this.coordinateSystem.toPhysical(new Point(scene.right, scene.y));
        const place = point => new Point(
            topLeft.x + point.x * (bottomRight.x - topLeft.x),
            topLeft.y + point.y * (bottomRight.y - topLeft.y));
        const placed = new LinearGradientBrush(place(brush.startPoint), place(brush.endPoint), brush.gradientStops);
        placed.absolute = true;
        return placed;
    }

    get canvas() {
        return this.canvasValue;
    }

    set canvas(value) {
        if (this.canvasValue === value) {
            return;
        }

        if (this.canvasValue != null) {
            for (const figure of this.figures.list) {
                figure.onRemovingFromCanvas(this.canvasValue);
            }

            this.behavior = null;
        }

        this.canvasValue = value;
        if (value != null) {
            for (const figure of this.figures.list) {
                figure.onAddingToCanvas(value);
            }
        }
    }

    /** The canvas changed its size: the view keeps its middle, and whoever fits the drawing hears of it */
    onSizeChanged(previousWidth, previousHeight, newWidth, newHeight) {
        this.coordinateSystem.onSizeChanged(previousWidth, previousHeight, newWidth, newHeight);
        this.sizeChanged?.(previousWidth, previousHeight, newWidth, newHeight);
    }

    get behavior() {
        return this.behaviorValue;
    }

    set behavior(value) {
        if (this.behaviorValue === value) {
            return;
        }

        if (this.behaviorValue != null) {
            this.behaviorValue.stopping();
            this.behaviorValue.drawing = null;
        }

        this.behaviorValue = value;
        if (value != null) {
            value.drawing = this;
            value.started();
        }
    }

    recalculate() {
        for (const figure of this.figures.list) {
            figure.recalculateAndUpdateVisual();
        }

        this.canvas?.invalidate?.();
    }

    /** Draws every figure in the order of the layers (ZIndex), the list's order within a layer */
    render(renderer) {
        const figures = [...this.figures.list].sort((a, b) => a.zIndex - b.zIndex);
        for (const figure of figures) {
            if (figure.exists) {
                figure.render(renderer);
            }
        }
    }

    addFromXml(element) {
        const deserializer = new DrawingDeserializer();
        deserializer.readDrawing(this, element);
        this.loadErrors = deserializer.isSuccess ? null : deserializer.getErrorReport();
    }

    /** Makes the text labels the drawing has now, and its show/hide boxes, fixed: a drag on one moves the view */
    fixLabels() {
        this.fixedLabels.clear();
        for (const figure of this.figures.list) {
            if (figure instanceof Label || figure instanceof ShowHideControl) {
                this.fixedLabels.add(figure);
            }
        }
    }

    /**
     * Every figure's numbers, for the parity check against the desktop's --dump: name, kind,
     * whether it exists, and what defines it
     */
    dump() {
        const number = value => isValidValue(value) ? Number(value.toFixed(9)) : String(value);
        const point = p => [number(p.x), number(p.y)];
        const result = [];
        for (const figure of this.figures.list) {
            if (figure instanceof CartesianGrid) {
                continue;
            }

            const entry = { name: figure.name, kind: figure.typeName, exists: figure.exists, visible: figure.visible };
            if (figure.isPoint === true) {
                entry.coordinates = point(figure.coordinates);
            } else if (figure.isLine === true) {
                entry.p1 = point(figure.coordinates.p1);
                entry.p2 = point(figure.coordinates.p2);
            } else if (figure.isEllipse === true) {
                entry.center = point(figure.center);
                entry.semiMajor = number(figure.semiMajor);
                entry.semiMinor = number(figure.semiMinor);
            } else if (figure.isPolygonalChain === true) {
                entry.vertices = (figure.vertexCoordinates ?? []).map(point);
            } else if (figure instanceof LabelBase) {
                entry.text = figure.processedText;
            } else if (figure.isNumber === true) {
                entry.value = number(figure.value);
            }

            result.push(entry);
        }

        return result;
    }
}
