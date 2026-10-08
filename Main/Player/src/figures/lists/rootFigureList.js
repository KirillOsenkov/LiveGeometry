// Port of Main/Avalonia/DynamicGeometry/Figures/Lists/RootFigureList.cs: the drawing's own
// list. Left out: Retire/Return (undo's), the settling of default names.

class RootFigureList extends FigureList {
    /** The figure of the list that the figure is, or is a part of (the knob of a slider); null if neither */
    findTopLevel(figure) {
        for (const item of this.list) {
            if (item === figure || (item instanceof CompositeFigure && containsRecursively(item.children, figure))) {
                return item;
            }
        }

        return null;
    }

    get hitTestCandidates() {
        return this.list.concat(this.drawing.unlistedAxisLines());
    }

    onItemAdded(item) {
        item.registerWithDependencies();
        item.onAddingToDrawing(this.drawing);
        if (this.drawing.canvas != null) {
            item.onAddingToCanvas(this.drawing.canvas);
            item.recalculateAndUpdateVisual();
        }
    }

    onItemRemoved(item) {
        item.onRemovingFromDrawing(this.drawing);
        if (this.drawing.canvas != null) {
            item.onRemovingFromCanvas(this.drawing.canvas);
        }

        item.unregisterFromDependencies();
    }

    /** Every dependency and dependent of a figure is in the drawing and registered both ways (CheckConsistency) */
    checkConsistency() {
        const figures = new Set();
        addRecursively(figures, this.list);
        for (const figure of this.list) {
            for (const dependency of figure.dependencies) {
                if (!figures.has(dependency)) {
                    throw new Error("Consistency check failed: dependency " + dependency + " of figure " + figure + " expected in the FigureList");
                }

                if (!dependency.dependents.includes(figure)) {
                    throw new Error("Consistency check failed: figure " + figure + " is not registered in the Dependents list of its dependency " + dependency);
                }
            }

            for (const dependent of figure.dependents) {
                if (!figures.has(dependent)) {
                    throw new Error("Consistency check failed: dependent " + dependent + " of figure " + figure + " expected in the FigureList");
                }

                if (!dependent.dependencies.includes(figure)) {
                    throw new Error("Consistency check failed: figure " + figure + " is not registered in the Dependencies list of its dependent " + dependent);
                }
            }
        }
    }
}

function containsRecursively(list, figure) {
    for (const item of list) {
        if (item === figure) {
            return true;
        }

        if (item instanceof CompositeFigure && containsRecursively(item.children, figure)) {
            return true;
        }
    }

    return false;
}

function addRecursively(figures, list) {
    for (const item of list) {
        figures.add(item);
        if (item instanceof CompositeFigure) {
            addRecursively(figures, item.children);
        }
    }
}
