// Port of Main/Avalonia/DynamicGeometry/Figures/Controls/ControlBase.cs: a figure that is a
// control (a label, a box), its coordinates the top-left corner, its size in pixels.

class ControlBase extends CoordinatesShapeBase {
    defaultZOrder() {
        return ZOrder.Controls;
    }

    hitTest(point) {
        if (this.rect.contains(point)) {
            return this;
        }

        return null;
    }

    applyStyle() {
        if (this.style == null) {
            return;
        }

        this.resolvedStyle = this.style.resolve(AppTheme.of(this.drawing).name);
        if (this.drawing != null) {
            this.updateVisual();
        }
    }

    /** The control's size in pixels, measured now */
    measureSize() {
        return new Size(0, 0);
    }

    /** The control's box in logical coordinates: from its top-left corner down and to the right */
    get rect() {
        const size = this.measureSize();
        const width = this.toLogicalLength(size.width);
        const height = this.toLogicalLength(size.height);
        return new Rect(this.coordinates.x, this.coordinates.y - height, width, height);
    }
}
