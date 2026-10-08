// Port of Main/Avalonia/DynamicGeometry/Figures/Lists/FigureList.cs: any set of figures,
// with the hit testing. The list is `list`; the C# collection's own methods are here by name.

class FigureList {
    constructor(drawing) {
        this.drawing = drawing;
        this.list = [];
    }

    get length() {
        return this.list.length;
    }

    at(index) {
        return this.list[index];
    }

    [Symbol.iterator]() {
        return this.list[Symbol.iterator]();
    }

    /** The first figure of that name, in the order of the list, looking inside composites and not at them (the indexer) */
    byName(name) {
        return FigureList.findByName(this.list, name);
    }

    static findByName(figures, name) {
        for (let i = 0; i < figures.length; i++) {
            const figure = figures[i];
            const found = figure instanceof CompositeFigure
                ? FigureList.findByName(figure.children, name)
                : figure.name === name ? figure : null;
            if (found != null) {
                return found;
            }
        }

        return null;
    }

    containsName(name) {
        return this.list.some(f => f.name === name);
    }

    contains(figure) {
        return this.list.includes(figure);
    }

    indexOf(figure) {
        return this.list.indexOf(figure);
    }

    add(figure) {
        if (figure.drawing == null) {
            figure.drawing = this.drawing;
        }

        this.list.push(figure);
        this.onItemAdded(figure);
    }

    insert(index, figure) {
        if (figure.drawing == null) {
            figure.drawing = this.drawing;
        }

        this.list.splice(index, 0, figure);
        this.onItemAdded(figure);
    }

    remove(figure) {
        const index = this.list.indexOf(figure);
        if (index >= 0) {
            this.removeAt(index);
        }
    }

    removeAt(index) {
        const item = this.list[index];
        this.list.splice(index, 1);
        this.onItemRemoved(item);
    }

    onItemAdded(item) {
    }

    onItemRemoved(item) {
    }

    recalculate() {
        for (const figure of this.list) {
            figure.recalculate();
        }
    }

    updateVisual() {
        for (const figure of this.list) {
            if (figure.exists) {
                figure.updateVisual();
            }
        }
    }

    // HitTest

    /** What a hit test looks at: the figures of the list, and for a drawing's list also the axis lines that are not in it yet */
    get hitTestCandidates() {
        return this.list;
    }

    /** The figure with the topmost ZIndex at the point, or null */
    hitTest(point, filter = figure => figure.visible && figure.isHitTestVisible) {
        let bestFoundSoFar = null;
        for (const item of this.hitTestCandidates) {
            // a figure that doesn't exist right now is nowhere to be clicked
            if (!item.exists) {
                continue;
            }

            const found = item.hitTest(point);
            if (found != null && found.exists && filter(found)) {
                if (bestFoundSoFar == null || bestFoundSoFar.zIndex <= found.zIndex) {
                    // of two nearby points, pick the one which is closer to the hit point
                    if (bestFoundSoFar != null
                        && bestFoundSoFar.isPoint === true
                        && found.isPoint === true
                        && bestFoundSoFar.coordinates.distance(point) < found.coordinates.distance(point)) {
                        continue;
                    }

                    bestFoundSoFar = found;
                }
            }
        }

        return bestFoundSoFar;
    }

    /** Every figure at the point, the topmost ZIndex first, of two points the nearer, else the one later in the list */
    hitTestAll(point, filter) {
        const found = [];
        let index = 0;
        for (const item of this.hitTestCandidates) {
            index++;
            if (!item.exists) {
                continue;
            }

            const hit = item.hitTest(point);
            if (hit != null && hit.exists && filter(hit) && !found.some(f => f.figure === hit)) {
                found.push({ figure: hit, index });
            }
        }

        found.sort((a, b) => {
            if (b.figure.zIndex !== a.figure.zIndex) {
                return b.figure.zIndex - a.figure.zIndex;
            }

            const da = a.figure.isPoint === true ? a.figure.coordinates.distance(point) : 0;
            const db = b.figure.isPoint === true ? b.figure.coordinates.distance(point) : 0;
            if (da !== db) {
                return da - db;
            }

            return b.index - a.index;
        });
        return found.map(f => f.figure);
    }

    /** Every visible figure at the point, the newer first, then by ZIndex (stable) */
    hitTestMany(point) {
        const result = [];
        const candidates = [...this.hitTestCandidates].reverse();
        for (const item of candidates) {
            if (!item.exists) {
                continue;
            }

            const found = item.hitTest(point);
            if (found != null && found.exists && found.visible && found.isHitTestVisible) {
                result.push(found);
            }
        }

        // a stable sort: higher z first
        return result.map((figure, index) => ({ figure, index }))
            .sort((a, b) => b.figure.zIndex - a.figure.zIndex || a.index - b.index)
            .map(entry => entry.figure);
    }
}
