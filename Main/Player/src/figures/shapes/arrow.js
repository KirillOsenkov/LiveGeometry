// Port of Main/Avalonia/DynamicGeometry/Figures/Shapes/Arrow.cs: a polygon that is a line
// with a head, as wide as its style says, the head sized in pixels. Built on a line.

class Arrow extends Polygon {
    /** Of the head of a hairline arrow, in pixels; both grow with the line width */
    static HeadLength = 11;
    static HeadHalfWidth = 4;
    static HeadGrowth = 3;

    constructor() {
        super();

        /** Whether the arrow is shaft and head in one solid color (an axis), or only the head (a vector) */
        this.drawsShaft = true;
        this.screenPoints = null;
    }

    onDependenciesChanged() {
    }

    /** Where the arrow is, in pixels: tail, headBase, tip, across (the unit normal), halfShaft, halfHead */
    measure() {
        const parentLine = this.dependencies[0];
        const line = parentLine instanceof LineBase ? parentLine.onScreenCoordinates : parentLine.coordinates;
        const tail = this.toPhysical(line.p1);
        let tip = this.toPhysical(line.p2);
        const length = tail.distance(tip);
        const width = this.resolvedStyle != null && this.resolvedStyle.strokeWidth != null ? this.resolvedStyle.strokeWidth : 1;
        const headLength = Math.min(Arrow.HeadLength + Arrow.HeadGrowth * width, length);
        const along = length > 1e-9 ? tip.minus(tail).scale(1 / length) : new Point(1, 0);

        // the head points at the end point, it doesn't hide under it
        const endPoint = parentLine != null && parentLine.dependencies.length > 1 && parentLine.dependencies[1] instanceof PointBase
            ? parentLine.dependencies[1]
            : null;
        if (endPoint != null && endPoint.visible) {
            const pointRadius = endPoint.pointSize / 2;
            if (pointRadius > 0 && length > pointRadius + headLength) {
                tip = tip.minus(along.scale(pointRadius));
            }
        }

        return {
            tail,
            headBase: tip.minus(along.scale(headLength)),
            tip,
            across: new Point(-along.y, along.x),
            halfShaft: Math.max(width / 2, 0.5),
            halfHead: Arrow.HeadHalfWidth + Arrow.HeadGrowth / 2 * width
        };
    }

    updateVisual() {
        if (this.drawing == null) {
            return;
        }

        const outline = this.measure();
        const headBase = outline.headBase;
        const across = outline.across;
        const points = [
            headBase.plus(across.scale(outline.halfHead)),
            outline.tip,
            headBase.minus(across.scale(outline.halfHead))
        ];
        if (this.drawsShaft) {
            points.push(headBase.minus(across.scale(outline.halfShaft)));
            points.push(outline.tail.minus(across.scale(outline.halfShaft)));
            points.push(outline.tail.plus(across.scale(outline.halfShaft)));
            points.push(headBase.plus(across.scale(outline.halfShaft)));
        }

        this.screenPoints = points;

        // the polygon's hit testing works on logical vertices
        this.vertexCoordinates = points.map(p => this.toLogical(p));
    }

    render(renderer) {
        if (!this.isShown || this.screenPoints == null) {
            return;
        }

        renderer.drawPolygon(this.screenPoints, null, Arrow.getBrush(this.style, this.drawing), true);
    }

    /** The color an arrow of the style is drawn in, as the style looks under the drawing's theme */
    static getBrush(style, drawing) {
        const resolved = style?.resolve(AppTheme.of(drawing).name);
        if (resolved instanceof LineStyle && resolved.color.a > 0) {
            return new SolidColorBrush(resolved.color);
        }

        return (resolved instanceof ShapeStyle ? resolved.fill : null) ?? new SolidColorBrush(Color.black);
    }
}

FigureTypes.register("Arrow", Arrow);
