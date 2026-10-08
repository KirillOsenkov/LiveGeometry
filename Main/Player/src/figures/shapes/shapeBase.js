// Port of Main/Avalonia/DynamicGeometry/Figures/Shapes/ShapeBase.cs: a figure drawn as one
// shape. In C# the shape is an Avalonia control the style is applied to; here the figure
// keeps the style resolved for the theme on screen (resolvedStyle) and draws itself in
// render. Left out: ghosts (a hidden figure shown while selected), selection halos, the
// pixel hit test of a shape (HitTestShape: what used it hit-tests its geometry instead).

class ShapeBase extends FigureBase {
    constructor() {
        super();
        this.zIndex = this.defaultZOrder();
        this.resolvedStyle = null;
    }

    defaultZOrder() {
        return ZOrder.Figures;
    }

    get isMovable() {
        return true;
    }

    /** Implementation of IMovable.MoveTo: MoveToCore (set coordinates) and UpdateVisual */
    moveTo(newPosition) {
        this.moveToCore(newPosition);
        this.updateVisual();
    }

    allowMove() {
        return !this.locked && this.dependencies.length === 0;
    }

    moveToCore(newLocation) {
    }

    updateVisual() {
    }

    /** Whether the shape is on the canvas and wants its geometry kept up to date */
    get isShown() {
        return this.exists && this.visible;
    }

    /** The style as it looks under the theme on screen: what the figure draws with */
    applyStyle() {
        if (this.style == null) {
            return;
        }

        this.resolvedStyle = this.style.resolve(AppTheme.of(this.drawing).name);
        if (this.drawing != null) {
            this.updateVisual();
        }
    }

    /** The width the shape is stroked with (Shape.StrokeThickness) */
    get strokeThickness() {
        return this.resolvedStyle?.strokeWidth ?? 1;
    }

    /** The stroke as a renderer takes it: color, width, dash */
    get stroke() {
        const style = this.resolvedStyle;
        if (style == null || style.color == null) {
            return null;
        }

        const width = style.strokeWidth ?? 1;
        return { color: style.color, width, dash: style.getDashArray ? style.getDashArray(width) : null };
    }

    /** The fill as a renderer takes it, or null for none */
    get fill() {
        const style = this.resolvedStyle;
        return style != null && style.fillBrush !== undefined ? style.fillBrush : null;
    }
}
