// Port of Main/Avalonia/DynamicGeometry/Serialization/DrawingDeserializer.cs: a drawing from
// its XML, read as far as it goes; what can't be read is left out with a line in words.

class DrawingDeserializer {
    constructor() {
        this.errors = [];
        this.beingRead = new Set();
        this.generateNewNames = false;
    }

    static openDrawing(canvas, text) {
        const drawing = new Drawing(canvas);
        const deserializer = new DrawingDeserializer();
        deserializer.readDrawing(drawing, text);
        if (!deserializer.isSuccess) {
            throw new Error(deserializer.getErrorReport());
        }

        return drawing;
    }

    /** The root element, or the text of a file */
    readDrawing(drawing, element) {
        if (typeof element === "string") {
            element = Xml.parse(element);
        }

        drawing.version = Xml.readDouble(element, "Version");
        this.readStyles(drawing, element);
        const figuresNode = Xml.element(element, "Figures");
        if (figuresNode != null) {
            // under the names the file gives them until all are in
            drawing.isReading = true;
            try {
                const figures = this.readFigures(figuresNode, drawing);
                for (const figure of figures) {
                    Actions.add(drawing, figure);
                }
            } finally {
                drawing.isReading = false;
            }
        }

        this.readViewport(drawing, element);
        DrawingDeserializer.readScenes(drawing, element);
        drawing.recalculate();

        // a drawing known to be from before the intersection order changed says so
        if (element.getAttribute("IntersectionOrder") === "Legacy") {
            for (const intersection of drawing.figures.list.filter(f => f instanceof IntersectionPoint)) {
                if (intersection.upgradeLegacyCircleAndLineOrder()) {
                    drawing.recalculate();
                }
            }
        }

        // Version 1: the offset of a label from what it labels is in pixels, not units of the plane
        if (drawing.version < 1) {
            for (const label of drawing.figures.list.filter(f => f instanceof LabelWithOffset)) {
                label.upgradeOffsetFromUnits();
            }

            drawing.version = Settings.currentDrawingVersion;
            drawing.recalculate();
        }
    }

    reportError(error) {
        this.errors.push(error);
    }

    get isSuccess() {
        return this.errors.length === 0;
    }

    getErrorReport() {
        return this.errors.join("\n");
    }

    readViewport(drawing, element) {
        const viewportNode = Xml.element(element, "Viewport");
        if (viewportNode == null) {
            return;
        }

        const minX = Xml.readDouble(viewportNode, "Left");
        const maxX = Xml.readDouble(viewportNode, "Right");
        const minY = Xml.readDouble(viewportNode, "Bottom");
        const maxY = Xml.readDouble(viewportNode, "Top");
        drawing.coordinateGrid.locked = Xml.readBool(viewportNode, "Locked", false);
        // a file that doesn't say has no grid
        drawing.coordinateGrid.visible = Xml.readBool(viewportNode, "Grid", false);
        drawing.coordinateGrid.showAxes = Xml.readBool(viewportNode, "Axes", true);
        drawing.coordinateSystem.gridStep = Xml.readDouble(viewportNode, "GridStep");
        drawing.viewport = new Rect(minX, minY, maxX - minX, maxY - minY);
        drawing.coordinateSystem.setViewport(minX, maxX, minY, maxY);
        const background = Xml.element(viewportNode, "Background");
        const gradient = background != null ? Xml.elements(background)[0] : null;
        if (gradient != null) {
            drawing.background = BrushSerializer.parseBrush(gradient);
        } else if (viewportNode.getAttribute("Color") != null) {
            drawing.background = new SolidColorBrush(ColorText.toColor(viewportNode.getAttribute("Color")));
        } else {
            drawing.background = null;
        }

        // the paper chosen for another theme: a child element named after the theme
        for (const themeNode of Xml.elements(viewportNode)) {
            if (AppTheme.byName(themeNode.localName) == null) {
                continue;
            }

            let paper = null;
            const themedBackground = Xml.element(themeNode, "Background");
            const themedGradient = themedBackground != null ? Xml.elements(themedBackground)[0] : null;
            if (themedGradient != null) {
                paper = BrushSerializer.parseBrush(themedGradient);
            } else if (themeNode.getAttribute("Color") != null) {
                paper = new SolidColorBrush(ColorText.toColor(themeNode.getAttribute("Color")));
            }

            drawing.setOverride(themeNode.localName, "background", paper);
        }
    }

    /** The suggested views, in the same Left/Top/Right/Bottom form as the viewport */
    static readScenes(drawing, element) {
        drawing.scenes = [];
        for (const sceneNode of Xml.elements(element, "Scene")) {
            const left = Xml.readDouble(sceneNode, "Left");
            const right = Xml.readDouble(sceneNode, "Right");
            const bottom = Xml.readDouble(sceneNode, "Bottom");
            const top = Xml.readDouble(sceneNode, "Top");
            if (right > left && top > bottom) {
                drawing.scenes.push(new Rect(left, bottom, right - left, top - bottom));
            }
        }
    }

    readStyles(drawing, element) {
        const stylesNode = Xml.element(element, "Styles");
        drawing.styleManager.clear();
        if (stylesNode == null) {
            drawing.styleManager.addDefaultStyles();
            return;
        }

        // a file has only the styles its figures use; the rest are the defaults
        const own = [];
        for (const styleNode of Xml.elements(stylesNode)) {
            const style = StyleReader.read(styleNode);
            if (style == null) {
                this.reportError("The style " + styleNode.getAttribute("Name") + " is left out: this version has no style of the kind " + styleNode.localName + ".");
                continue;
            }

            own.push(style);
        }

        drawing.styleManager.addWithDefaults(own);
    }

    readFigures(figuresNode, drawing, figures = new Map()) {
        const result = [];
        const nodeMap = new Map();
        for (const figureNode of Xml.elements(figuresNode)) {
            const name = figureNode.getAttribute("Name");
            if (name == null || name === "") {
                this.reportError(figureNode.localName + " without a name is left out.");
            } else if (nodeMap.has(name)) {
                this.reportError("Two figures are called " + name + ": only the first is read.");
            } else {
                nodeMap.set(name, figureNode);
            }
        }

        for (const figureName of nodeMap.keys()) {
            this.readFigure(figureName, figures, nodeMap, drawing, figure => result.push(figure));
        }

        return result;
    }

    readFigure(figureName, alreadyDeserializedFigures, nodeMap, drawing, callbackWhenCreated) {
        if (alreadyDeserializedFigures.has(figureName)) {
            return;
        }

        const figureNode = nodeMap.get(figureName);
        if (figureNode == null || this.beingRead.has(figureName)) {
            return;
        }

        this.beingRead.add(figureName);
        try {
            this.readFigureNode(figureName, figureNode, alreadyDeserializedFigures, nodeMap, drawing, callbackWhenCreated);
        } finally {
            this.beingRead.delete(figureName);
        }
    }

    readFigureNode(figureName, figureNode, alreadyDeserializedFigures, nodeMap, drawing, callbackWhenCreated) {
        const type = FigureTypes.find(figureNode.localName);
        if (type == null) {
            this.reportError(figureName + " is left out: this version has no figure of the kind " + figureNode.localName + ".");
            return;
        }

        // an axis is the drawing's own, never a second one
        if (figureNode.localName === "AxisLine" && drawing != null) {
            const axis = drawing.getAxisLine(AxisLine.readDirection(figureNode));
            alreadyDeserializedFigures.set(figureName, axis);
            if (!drawing.figures.contains(axis)) {
                callbackWhenCreated(axis);
            }

            return;
        }

        const dependencyNodes = Xml.elements(figureNode, "Dependency");
        const dependencyNames = dependencyNodes.map(e => e.getAttribute("Name"));
        for (const dependencyName of dependencyNames) {
            if (dependencyName != null) {
                this.readFigure(dependencyName, alreadyDeserializedFigures, nodeMap, drawing, callbackWhenCreated);
            }
        }

        const dependencies = [];
        for (let i = 0; i < dependencyNames.length; i++) {
            const dependencyName = dependencyNames[i];
            let existingDependency = dependencyName != null ? alreadyDeserializedFigures.get(dependencyName) : undefined;
            if (existingDependency === undefined) {
                this.reportError(figureName + " is left out: it is built on " + (dependencyName ?? "a figure without a name") + ", which could not be read.");
                return;
            }

            // a part of the figure (a vertex a regular polygon works out), not the figure
            const partName = dependencyNodes[i].getAttribute("Part");
            if (partName != null) {
                existingDependency = existingDependency.getPart != null ? existingDependency.getPart(partName) : null;
                if (existingDependency == null) {
                    this.reportError(figureName + " is left out: " + dependencyName + " has no part " + partName + ".");
                    return;
                }
            }

            dependencies.push(existingDependency);
        }

        const instance = new type();
        instance.drawing = drawing;
        instance.dependencies = dependencies;
        if (!this.generateNewNames) {
            instance.name = figureName;
        }

        if (drawing.figures.byName(instance.name) != null) {
            instance.visible = Xml.readBool(figureNode, "Visible", true);
            instance.generateNewNameIfNecessary(drawing);
        }

        alreadyDeserializedFigures.set(figureName, instance);
        try {
            instance.readXml(figureNode);
        } catch (error) {
            this.reportError(figureName + " was not read in full: " + error.message);
            callbackWhenCreated(instance);
            return;
        }

        callbackWhenCreated(instance);
    }
}
