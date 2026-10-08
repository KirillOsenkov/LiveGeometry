// Port of Main/Avalonia/DynamicGeometry/Figures/Lists/CompositeFigure.cs: a figure made of
// figures of the library (a slider, a vector, a regular polygon).

class CompositeFigure extends FigureBase {
    constructor() {
        super();
        this.children = [];
    }

    clearChildren() {
        for (const child of [...this.children].reverse()) {
            this.removeChild(child);
        }
    }

    addChild(figure) {
        figure.registerWithDependencies();
        this.children.push(figure);
        if (this.drawing != null) {
            figure.drawing = this.drawing;
            figure.onAddingToDrawing(this.drawing);
            if (this.drawing.canvas != null) {
                figure.onAddingToCanvas(this.drawing.canvas);
            }
        }
    }

    removeChild(figure) {
        if (this.drawing != null) {
            if (this.drawing.canvas != null) {
                figure.onRemovingFromCanvas(this.drawing.canvas);
            }

            figure.onRemovingFromDrawing(this.drawing);
            figure.drawing = null;
        }

        const index = this.children.indexOf(figure);
        if (index >= 0) {
            this.children.splice(index, 1);
        }

        figure.unregisterFromDependencies();
    }

    onAddingToDrawing(drawing) {
        super.onAddingToDrawing(drawing);
        for (const item of this.children) {
            item.onAddingToDrawing(drawing);
        }
    }

    onRemovingFromDrawing(drawing) {
        super.onRemovingFromDrawing(drawing);
        for (const item of this.children) {
            item.onRemovingFromDrawing(drawing);
        }
    }

    onAddingToCanvas(newContainer) {
        for (const figure of this.children) {
            figure.onAddingToCanvas(newContainer);
        }

        super.onAddingToCanvas(newContainer);
    }

    onRemovingFromCanvas(leavingContainer) {
        super.onRemovingFromCanvas(leavingContainer);
        for (const figure of this.children) {
            figure.onRemovingFromCanvas(leavingContainer);
        }
    }

    recalculate() {
        for (const figure of this.children) {
            if (figure.exists) {
                figure.recalculate();
            }
        }
    }

    updateVisual() {
        if (!this.visible) {
            return;
        }

        for (const figure of this.children) {
            if (figure.exists) {
                figure.updateVisual();
            }
        }
    }

    render(renderer) {
        if (!this.visible || !this.exists) {
            return;
        }

        for (const figure of this.children) {
            if (figure.exists) {
                figure.render(renderer);
            }
        }
    }

    updateExistence() {
        super.updateExistence();
        for (const figure of this.children) {
            figure.updateExistence();
        }
    }

    get selected() {
        return super.selected || this.children.some(c => c.selected);
    }

    set selected(value) {
        super.selected = value;
        for (const item of this.children) {
            item.selected = value;
        }
    }

    get visible() {
        return this.mVisible;
    }

    set visible(value) {
        this.mVisible = value;
        for (const item of this.children) {
            item.visible = value;
        }
    }

    get drawing() {
        return this.drawingValue ?? null;
    }

    set drawing(value) {
        this.drawingValue = value;
        if (this.children != null) {
            for (const item of this.children) {
                item.drawing = value;
            }
        }
    }

    applyStyle() {
        for (const figure of this.children) {
            figure.applyStyle();
        }
    }

    /** The visible child with the topmost ZIndex at the point, or null */
    hitTestWith(point, filter) {
        let bestFoundSoFar = null;
        for (const item of this.children) {
            if (!filter(item)) {
                continue;
            }

            const found = item.hitTest(point);
            if (found != null) {
                if (bestFoundSoFar == null || bestFoundSoFar.zIndex <= found.zIndex) {
                    bestFoundSoFar = found;
                }
            }
        }

        return bestFoundSoFar;
    }

    hitTest(point) {
        return this.hitTestWith(point, f => f.visible);
    }
}
