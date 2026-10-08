// Port of what the player needs of Main/Avalonia/DynamicGeometry/Actions/Actions.cs: every
// change of the drawing goes through here, which is where undo attaches when the player
// grows one. For now each action is done directly.

const Actions = {
    /** A figure into the drawing (AddFigureAction) */
    add(drawing, figure) {
        drawing.figures.add(figure);
    },

    /**
     * Moves the movables by the offset, then works out what is built on them, in the order
     * given (dependency order), with the drawing marked as moving (MoveAction)
     */
    move(drawing, movables, offset, toRecalculate) {
        for (const movable of movables) {
            movable.moveTo(movable.coordinates.plus(offset));
        }

        if (movables.some(movable => movable instanceof CoordinateSystem)) {
            drawing.viewChanged?.();
        } else {
            drawing.figuresMoved?.();
        }

        if (toRecalculate != null) {
            drawing.isMoving = true;
            try {
                for (const figure of toRecalculate) {
                    figure.recalculateAndUpdateVisual();
                }
            } finally {
                drawing.isMoving = false;
            }
        }

        drawing.canvas?.invalidate?.();
    }
};
