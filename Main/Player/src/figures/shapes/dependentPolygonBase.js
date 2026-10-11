// Port of Main/Avalonia/DynamicGeometry/Figures/Shapes/DependentPolygonBase.cs: a figure with
// parts that other figures can be built on, though the parts are not figures of the
// drawing: the vertices and sides a regular polygon works out. A file refers to a part by
// its owner's name and the part's (<Dependency Name="p" Part="Vertex3" />).

/** A vertex the polygon works out */
class PolygonVertex extends PointBase {
    constructor(owner) {
        super();
        this.owner = owner;
    }

    get isFigurePart() {
        return true;
    }

    onAddingToDrawing(drawing) {
    }
}

/** A side between two vertices */
class PolygonSide extends Segment {
    constructor(owner) {
        super();
        this.owner = owner;
    }

    get isFigurePart() {
        return true;
    }

    onAddingToDrawing(drawing) {
    }
}

/** The filled inside: to the user, the polygon itself */
class InteriorPolygon extends Polygon {
    constructor(owner) {
        super();
        this.owner = owner;
    }

    get isFigurePart() {
        return true;
    }

    onAddingToDrawing(drawing) {
    }
}

/** What a selection means where a figure has parts */
const FigureParts = {
    /** What a click on the figure selects: a vertex or a side by itself, but the inside is the figure */
    selectionTarget(figure) {
        return figure.isFigurePart === true && !figure.owner.selectableParts.includes(figure) ? figure.owner : figure;
    },

    /** The figure of the drawing that the figure is, or is a part of */
    whole(figure) {
        return figure.isFigurePart === true ? figure.owner : figure;
    }
};

class DependentPolygonBase extends CompositeFigure {
    static VertexPart = "Vertex";
    static SidePart = "Side";
    static InteriorPart = "Interior";

    constructor() {
        super();
        this.vertices = [];
        this.sides = [];
        this.retiredVertices = [];
        this.retiredSides = [];
        this.layer = ZOrder.Polygons;
        this.polygon = new InteriorPolygon(this);
        this.children.push(this.polygon);
        this.isOnCanvas = false;
        this.partsUnregistered = true;
    }

    get isShapeWithInterior() {
        return true;
    }

    get isPolygonalChain() {
        return true;
    }

    get isFigureParts() {
        return true;
    }

    get selectableParts() {
        return [...this.vertices, ...this.sides];
    }

    /** Vertex2, Vertex3... (Vertex1 is the point the polygon is built on), Side1, Side2... and Interior */
    getPartName(part) {
        let index = this.vertices.indexOf(part);
        if (index >= 0) {
            return DependentPolygonBase.VertexPart + (index + 2);
        }

        index = this.sides.indexOf(part);
        if (index >= 0) {
            return DependentPolygonBase.SidePart + (index + 1);
        }

        return part === this.polygon ? DependentPolygonBase.InteriorPart : null;
    }

    getPart(partName) {
        if (partName === DependentPolygonBase.InteriorPart) {
            return this.polygon;
        }

        let index = DependentPolygonBase.tryGetIndex(partName, DependentPolygonBase.VertexPart);
        if (index != null) {
            index -= 2;
            return index >= 0 && index < this.vertices.length ? this.vertices[index] : null;
        }

        index = DependentPolygonBase.tryGetIndex(partName, DependentPolygonBase.SidePart);
        if (index != null) {
            index -= 1;
            return index >= 0 && index < this.sides.length ? this.sides[index] : null;
        }

        return null;
    }

    static tryGetIndex(partName, kind) {
        if (partName == null || !partName.startsWith(kind)) {
            return null;
        }

        const rest = partName.substring(kind.length);
        return /^-?\d+$/.test(rest) ? parseInt(rest, 10) : null;
    }

    /** Whether something that is not the polygon's own is built on the part */
    hasOutsideDependents(part) {
        return part.dependents.some(dependent => !this.children.includes(dependent));
    }

    onAddingToCanvas(newContainer) {
        super.onAddingToCanvas(newContainer);
        this.isOnCanvas = true;
    }

    onRemovingFromCanvas(leavingContainer) {
        super.onRemovingFromCanvas(leavingContainer);
        this.isOnCanvas = false;
    }

    /** Lists a part with what it is built on, if the polygon is in the drawing; else that waits until it is */
    registerPart(part) {
        if (!this.partsUnregistered) {
            part.registerWithDependencies();
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

    onAddingToDrawing(drawing) {
        super.onAddingToDrawing(drawing);
        if (this.partsUnregistered) {
            this.partsUnregistered = false;
            for (const part of this.children) {
                part.registerWithDependencies();
            }
        }
    }

    /** The style of the inside, which is the polygon's own */
    get style() {
        return this.polygon.style;
    }

    set style(value) {
        this.polygon.style = value;
    }

    /** The styles of the parts: <Sides Style>, <Vertices Style>, and <Part Name Style> for the odd ones */
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
        for (const side of this.sides) {
            apply(side, sidesStyle);
        }

        const verticesStyle = Xml.element(element, "Vertices")?.getAttribute("Style") ?? null;
        for (const vertex of this.vertices) {
            apply(vertex, verticesStyle);
        }

        for (const part of Xml.elements(element, "Part")) {
            apply(this.getPart(part.getAttribute("Name")), part.getAttribute("Style"));
        }
    }

    recreate(sideCount, recalculate = true) {
        this.adjustVerticesList(sideCount);
        this.adjustSides(sideCount);
        this.adjustPolygon();

        // a hidden polygon must hide its parts too
        if (!this.visible) {
            for (const part of this.children) {
                part.visible = false;
            }
        }

        if (recalculate) {
            this.recalculate();
        }
    }

    collectPolygonDependencies(collector) {
        for (const item of this.vertices) {
            collector(item);
        }

        collector(this);
    }

    adjustPolygon() {
        this.polygon.unregisterFromDependencies();
        const allVertices = [];
        this.collectPolygonDependencies(figure => allVertices.push(figure));
        this.polygon.dependencies = allVertices;
        this.registerPart(this.polygon);
    }

    get area() {
        return this.polygon.area;
    }

    get isPerimeterProvider() {
        return true;
    }

    get perimeter() {
        return this.polygon.perimeter;
    }

    get vertexCoordinates() {
        return this.vertices.map(v => v.coordinates);
    }

    adjustSides(sideCount) {
        if (this.sides.length < sideCount) {
            const needed = sideCount - this.sides.length;
            for (let i = 0; i < needed; i++) {
                this.addSide(sideCount);
            }
        } else if (this.sides.length > sideCount) {
            const extra = this.sides.length - sideCount;
            for (let i = 0; i < extra; i++) {
                this.removeSide();
            }
        }
    }

    removeSide() {
    }

    addSide(sideCount) {
    }

    adjustVerticesList(sideCount) {
    }

    removeVertex() {
        const vertex = this.vertices.pop();
        this.removePart(vertex);
        this.retiredVertices.push(vertex);
    }

    /** A part the count no longer needs leaves: out of its dependencies, out of the children */
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

    /** The style a new part takes: the one most of the others have, or none yet */
    static commonStyle(parts) {
        const counts = new Map();
        for (const part of parts) {
            if (part.style != null) {
                counts.set(part.style, (counts.get(part.style) ?? 0) + 1);
            }
        }

        let best = null;
        let bestCount = 0;
        for (const [style, count] of counts) {
            if (count > bestCount) {
                best = style;
                bestCount = count;
            }
        }

        return best;
    }

    addVertex() {
        const isNew = this.retiredVertices.length === 0;
        const vertex = isNew ? new PolygonVertex(this) : this.retiredVertices.pop();
        const common = isNew ? DependentPolygonBase.commonStyle(this.vertices) : null;
        vertex.dependencies = [this];
        vertex.visible = this.visible;
        this.registerPart(vertex);
        vertex.drawing = this.drawing;
        vertex.z = this.z;
        this.vertices.push(vertex);
        this.children.push(vertex);
        if (this.isOnCanvas) {
            vertex.onAddingToCanvas(this.drawing.canvas);
        }

        if (common != null) {
            vertex.style = common;
        } else if (isNew) {
            const style = this.drawing?.styleManager.getStyle(StyleManager.DependentPointStyleName);
            if (style != null) {
                vertex.style = style;
            }
        }
    }
}
