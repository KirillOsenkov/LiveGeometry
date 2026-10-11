// Port of Main/Avalonia/DynamicGeometry/Figures/Lines/Vector.cs: a segment with an arrowhead,
// an ILine like a segment. The shaft is a segment drawn from the start to the head's base,
// which a dash can break; the arrow draws only the head.

/** The segment from the start to the end, drawn from the start to the head */
class VectorShaft extends Segment {
    constructor(arrow) {
        super();
        this.arrow = arrow;
    }

    updateVisual() {
        if (!this.isShown || this.drawing == null) {
            this.screenLine = null;
            return;
        }

        const outline = this.arrow.measure();
        this.screenLine = new PointPair(outline.tail, outline.headBase);
    }

    /** The head's color, also for a style whose line is transparent (a polygon's style: its fill) */
    get stroke() {
        const stroke = super.stroke;
        if (stroke == null || this.style == null) {
            return stroke;
        }

        const brush = Arrow.getBrush(this.style, this.drawing);
        return { color: brush instanceof SolidColorBrush ? brush.color : stroke.color, width: stroke.width, dash: stroke.dash };
    }
}

class Vector extends CompositeFigure {
    constructor() {
        super();
        this.arrow = new Arrow();
        this.arrow.drawsShaft = false;
        this.arrow.layer = ZOrder.Vectors;
        this.line = new VectorShaft(this.arrow);
        this.line.layer = ZOrder.Vectors;
        this.arrow.dependencies = [this.line];
        this.children.push(this.line, this.arrow);
        this.layer = ZOrder.Vectors;
    }

    get isVector() {
        return true;
    }

    get isLine() {
        return true;
    }

    get isLinearFigure() {
        return true;
    }

    get isLengthProvider() {
        return true;
    }

    /** The arrow's style, which the shaft is drawn in too */
    get style() {
        return this.arrow.style;
    }

    set style(value) {
        this.arrow.style = value;
        this.line.style = value;
    }

    onDependenciesChanged() {
        this.line.dependencies = this.dependencies;
    }

    onAddingToCanvas(newContainer) {
        // left to itself the arrow would get the default polygon style; a vector is a line
        if (this.arrow.style == null && this.drawing != null) {
            this.style = this.drawing.styleManager.list.find(s => s.constructor === LineStyle) ?? null;
        }

        super.onAddingToCanvas(newContainer);
        this.arrow.ensureStyleAssigned();
        this.line.style = this.arrow.style;
    }

    /** The arrow is a filled polygon, so a hit on it has to land on the drawn pixels: the shaft gets the same room around it as a segment */
    hitTestWith(point, filter) {
        let result = this.arrow.hitTest(point) ?? this.line.hitTest(point);
        if (result != null) {
            result = this;
            if (!filter(result)) {
                result = null;
            }
        }

        return result;
    }

    /** Where the vector is, shown or not, as a segment answers */
    hitTest(point) {
        return this.hitTestWith(point, figure => true);
    }

    readXml(element) {
        // not the composite's: there are no children to read, the constructor made them
        this.visible = Xml.readBool(element, "Visible", true);
        this.locked = Xml.readBool(element, "Locked", false);
        this.isHitTestVisible = Xml.readBool(element, "IsHitTestVisible", true);
        const styleName = element.getAttribute("Style");
        if (styleName != null && this.drawing != null && this.drawing.styleManager != null) {
            const style = this.drawing.styleManager.get(styleName);
            if (style != null) {
                this.style = style;
            }
        }
    }

    get coordinates() {
        return this.line.coordinates;
    }

    get length() {
        return this.line.length;
    }

    get angle() {
        return GeometryMath.toDegrees(this.direction);
    }

    get direction() {
        return GeometryMath.getAngle(this.coordinates.p1, this.coordinates.p2);
    }

    getNearestParameterFromPoint(point) {
        return this.line.getNearestParameterFromPoint(point);
    }

    getPointFromParameter(parameter) {
        return this.line.getPointFromParameter(parameter);
    }

    getParameterDomain() {
        return this.line.getParameterDomain();
    }

    get center() {
        return this.line.center;
    }

    toString() {
        return this.name;
    }
}

FigureTypes.register("Vector", Vector);
