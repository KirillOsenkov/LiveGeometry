// Port of Main/Avalonia/DynamicGeometry/Figures/FigureBase.cs and of the extension methods
// of Figures/IFigureExtensions.cs that act on one figure (dependsOn, registerWithDependencies,
// recalculateAllDependents, ...), which are methods here. Left out: the naming of new
// figures after their points and the rename waves (a file's figures have their names), the
// property grid, the name label (FigureLabel comes later), Clone.
//
// The "interfaces" a figure implements in C# (IPoint, ILine, ILinearFigure, IEllipse,
// ICircle, ILengthProvider, IAngleProvider, INumber, IShapeWithInterior, IMovable,
// IMovableParts) are getters that answer true: isPoint, isLine, isLinearFigure, isEllipse,
// isCircle, isLengthProvider, isAngleProvider, isNumber, isShapeWithInterior, isMovable,
// isMovableParts.

class FigureBase {
    constructor() {
        this.drawing = null;
        this.nameValue = null;
        this.mDependencies = [];
        this.mDependents = [];
        this.mExists = true;
        this.mVisible = true;
        this.isHitTestVisible = true;
        this.lockedValue = false;
        this.auxiliary = false;
        this.flipped = false;
        this.layerValue = 0;
        this.zValue = 0;
        this.styleValue = null;
        this.selectedValue = false;
        this.enabled = true;
    }

    /** The C# type name, which is the file's element name (FigureTypes.register sets it) */
    get typeName() {
        return this.constructor.typeName ?? this.constructor.name;
    }

    get name() {
        return this.nameValue;
    }

    set name(value) {
        if (value == null || value === "") {
            value = this.generateFigureName();
        }

        this.nameValue = value;
    }

    /** Numbered by the type: Segment1, Segment2... (GenerateNewName); points and sliders have their own */
    generateFigureName() {
        const className = this.typeName;
        for (let i = 1; ; i++) {
            const candidate = className + i;
            if (this.nameAvailable(candidate)) {
                return candidate;
            }
        }
    }

    nameAvailable(name) {
        if (this.drawing == null) {
            return true;
        }

        return !this.drawing.figures.list.some(f => f !== this && f.name === name);
    }

    /** A name while the figure has none, or one another figure of the drawing has (GenerateNewNameIfNecessary) */
    generateNewNameIfNecessary(drawing) {
        while (this.name == null || drawing.figures.list.some(f => f.name === this.name && f !== this)) {
            this.name = this.generateFigureName();
        }
    }

    onAddingToDrawing(drawing) {
        this.generateNewNameIfNecessary(drawing);
    }

    onRemovingFromDrawing(drawing) {
    }

    toString() {
        return this.name;
    }

    get selected() {
        return this.selectedValue;
    }

    set selected(value) {
        this.selectedValue = value;
    }

    get visible() {
        return this.mVisible;
    }

    set visible(value) {
        this.mVisible = value;
    }

    get locked() {
        return this.lockedValue;
    }

    set locked(value) {
        this.lockedValue = value;
    }

    /** The attributes every figure may have (FigureBase.ReadXml); a subclass reads its own after */
    readXml(element) {
        this.visible = Xml.readBool(element, "Visible", true);
        this.locked = Xml.readBool(element, "Locked", false);
        this.auxiliary = Xml.readBool(element, "Auxiliary", false);
        this.isHitTestVisible = Xml.readBool(element, "IsHitTestVisible", true);
        this.z = Xml.readInt(element, "Z", 0);
        const styleName = element.getAttribute("Style");
        if (styleName != null && this.drawing != null && this.drawing.styleManager != null) {
            const style = this.drawing.styleManager.get(styleName);
            if (style != null) {
                this.style = style;
            }
        }

        this.flipped = Xml.readBool(element, "Flipped", false);
    }

    get serializable() {
        return true;
    }

    /** The canvas the drawing is on: what gives the pixels their size */
    get canvas() {
        return this.drawing?.canvas ?? null;
    }

    /** The canvas's borders in logical coordinates */
    get canvasLogicalBorders() {
        return this.toLogicalPair(this.canvas.getBorderRectangle());
    }

    get dependencies() {
        return this.mDependencies;
    }

    set dependencies(value) {
        this.mDependencies = value == null ? [] : [...value];
        this.onDependenciesChanged();
    }

    onDependenciesChanged() {
    }

    get dependents() {
        return this.mDependents;
    }

    /** The layer the figure is drawn in, its kind's (ZOrder) */
    get layer() {
        return this.layerValue;
    }

    set layer(value) {
        this.layerValue = value;
        this.onZIndexChanged();
    }

    /** Where the figure is among those of its band of layers: 0 unless a file says otherwise (ZOrders) */
    get z() {
        return this.zValue;
    }

    set z(value) {
        this.zValue = value;
        this.onZIndexChanged();
    }

    /** What the layer and the Z come to: the order the figures are drawn and hit in */
    get zIndex() {
        return ZOrders.encode(this.layer, this.z);
    }

    /** The layer or the Z changed */
    onZIndexChanged() {
    }

    get exists() {
        return this.mExists;
    }

    set exists(value) {
        this.mExists = value;
    }

    /** The coordinates of the point that is the dependency at the index */
    point(index) {
        const dependency = index < this.mDependencies.length ? this.mDependencies[index] : null;
        if (dependency != null && dependency.isPoint === true) {
            return dependency.coordinates;
        }

        return this.missingPoint(index);
    }

    /** The coordinates of the line that is the dependency at the index */
    line(index) {
        const dependency = index < this.mDependencies.length ? this.mDependencies[index] : null;
        if (dependency != null && dependency.isLine === true) {
            return dependency.coordinates;
        }

        this.reportMissingDependency(index, "a line");
        return new PointPair(Point.infinite, Point.infinite);
    }

    /**
     * The figure asked for a point it is not built on: too few dependencies, or one of another
     * kind. A file may say that: while the file is read the figure is noted for the deserializer
     * to leave out (drawing.invalidFigures), made not to exist and given a point that is
     * nowhere. At any other time it is a bug, and throws.
     */
    missingPoint(index) {
        this.reportMissingDependency(index, "a point");
        return Point.infinite;
    }

    reportMissingDependency(index, what) {
        const count = this.mDependencies.length;
        const message = this.toString() + " is left out: it is built on " + count + " figure" + (count === 1 ? "" : "s")
            + " and needs " + what + " as the " + FigureBase.ordinal(index + 1) + ".";
        const drawing = this.drawing;
        if (drawing == null || !drawing.isReading) {
            throw new Error(message);
        }

        this.exists = false;
        if (!drawing.invalidFigures.some(invalid => invalid.figure === this)) {
            drawing.invalidFigures.push({ figure: this, message });
        }
    }

    static ordinal(number) {
        const rest = number % 100;
        const suffix = rest === 11 || rest === 12 || rest === 13
            ? "th"
            : number % 10 === 1 ? "st" : number % 10 === 2 ? "nd" : number % 10 === 3 ? "rd" : "th";
        return number + suffix;
    }

    get style() {
        return this.styleValue;
    }

    set style(value) {
        if (this.styleValue === value) {
            return;
        }

        this.styleValue = value;
        this.applyStyle();
    }

    ensureStyleAssigned() {
        if (this.style == null && this.drawing != null) {
            this.style = this.drawing.styleManager.assignDefaultStyle(this);
        }
    }

    onAddingToCanvas(newContainer) {
        this.ensureStyleAssigned();
    }

    /** The style resolved for the theme on screen is read again (a subclass caches what it draws with) */
    applyStyle() {
    }

    onRemovingFromCanvas(leavingContainer) {
    }

    updateExistence() {
        for (let i = 0; i < this.mDependencies.length; i++) {
            if (!this.mDependencies[i].exists) {
                this.exists = false;
                return;
            }
        }

        this.exists = true;
    }

    recalculate() {
    }

    /** Takes Coordinates or whatever other location information is current and updates the visual representation */
    updateVisual() {
    }

    /** Draws the figure (no counterpart in C#, where Avalonia draws the shapes) */
    render(renderer) {
    }

    /** The figure at the logical point, or null (what a figure's own hit test answers) */
    hitTest(point) {
        return null;
    }

    get center() {
        return new Point(0, 0);
    }

    // Coordinates

    get cursorTolerance() {
        return this.drawing.coordinateSystem.cursorTolerance;
    }

    toPhysical(point) {
        return this.drawing.coordinateSystem.toPhysical(point);
    }

    toPhysicalLength(logicalLength) {
        return this.drawing.coordinateSystem.toPhysicalLength(logicalLength);
    }

    toPhysicalPair(pointPair) {
        return this.drawing.coordinateSystem.toPhysicalPair(pointPair);
    }

    toLogical(pixel) {
        return this.drawing.coordinateSystem.toLogical(pixel);
    }

    toLogicalLength(pixelLength) {
        return this.drawing.coordinateSystem.toLogicalLength(pixelLength);
    }

    toLogicalPair(pointPair) {
        return this.drawing.coordinateSystem.toLogicalPair(pointPair);
    }

    // IFigureExtensions

    /** Whether the figure directly or indirectly depends on the possible dependency (a figure depends on itself) */
    dependsOn(possibleDependency) {
        if (this === possibleDependency) {
            return true;
        }

        if (this.dependencies.length === 0) {
            return false;
        }

        if (this.dependencies.includes(possibleDependency)) {
            return true;
        }

        for (const directDependency of this.dependencies) {
            if (directDependency.dependsOn(possibleDependency)) {
                return true;
            }
        }

        return false;
    }

    directlyDependsOn(possibleDependency) {
        return this === possibleDependency || this.dependencies.includes(possibleDependency);
    }

    /** Whether a tool or a tied value may take a length from the figure: a label that says no number gives none, nor the mark of an angle */
    givesLength() {
        return this.isLengthProvider === true && this.isAngleArc !== true && Label.givesNumber(this);
    }

    givesAngle() {
        return this.isAngleProvider === true && Label.givesNumber(this);
    }

    /** The figure and everything built on it, in dependency order, each worked out again */
    recalculateAllDependents() {
        const dependentsToRecalculate = DependencyAlgorithms.findDescendants(f => f.dependents, [this]);
        dependentsToRecalculate.reverse();
        for (const dependent of dependentsToRecalculate) {
            dependent.recalculateAndUpdateVisual();
        }
    }

    recalculateAndUpdateVisual() {
        if (this.drawing == null) {
            return;
        }

        this.updateExistence();
        this.recalculate();
        this.updateVisual();
    }

    registerWithDependencies() {
        this.addDependencies(this.dependencies);
    }

    unregisterFromDependencies() {
        this.removeDependencies(this.dependencies);
    }

    addDependencies(dependencies) {
        for (const dependency of dependencies) {
            dependency.dependents.push(this);
        }
    }

    removeDependencies(dependencies) {
        for (const dependency of dependencies) {
            if (dependency != null) {
                const index = dependency.dependents.indexOf(this);
                if (index >= 0) {
                    dependency.dependents.splice(index, 1);
                }
            }
        }
    }

    /** Adds a dependency at the index, registered, and works the figure out again (InsertDependencyCore) */
    insertDependencyCore(index, dependency) {
        this.mDependencies.splice(index, 0, dependency);
        dependency.dependents.push(this);
        this.onDependenciesChanged();
        this.recalculateAndUpdateVisual();
    }

    removeDependencyCore(index, dependency) {
        this.mDependencies.splice(index, 1);
        const at = dependency.dependents.indexOf(this);
        if (at >= 0) {
            dependency.dependents.splice(at, 1);
        }

        this.onDependenciesChanged();
        this.recalculateAndUpdateVisual();
    }
}

/** Dependencies.Exists(): whether every figure of the list exists */
function allExist(figures) {
    for (let i = 0; i < figures.length; i++) {
        if (!figures[i].exists) {
            return false;
        }
    }

    return true;
}

/** ToPoints: the coordinates of the points among the figures */
function toPoints(figures) {
    return figures.filter(f => f.isPoint === true).map(p => p.coordinates);
}
