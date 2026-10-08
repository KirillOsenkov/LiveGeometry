// Port of Main/Avalonia/LiveGeometry/Gallery/GalleryDrawing.cs (the app's, not the
// library's): what the drawings of the gallery have in common, a caption - a heading and an
// explanation, two labels named "Title" and "Description", and maybe a third under them,
// "Hint", hidden until a show/hide box brings it up - pinned to the screen, on a plate,
// wrapped to its column. fit decides where it goes and how much room the figure gets, for
// the element the drawing plays in: a column at the right in a wide one, a strip at the
// bottom in a tall one. Without it a drawing exported from the app keeps the pin, offsets
// and width it had in the app's window, which a small element clips. Left out: hiding the
// text for a thumbnail.

const GalleryDrawing = {
    TitleName: "Title",
    DescriptionName: "Description",
    HintName: "Hint",

    /** The caption beside the figure is a column this wide (wider only for a heading that needs it) */
    CaptionWidth: 400,

    // a column narrower than this reads badly: the caption goes under the figure instead
    MinimumCaptionWidth: 240,

    // the caption's plate from the edge of the canvas (the text is a padding further in)
    TextMarginPixels: 16,

    // between the figure and the caption, and between the heading and the explanation:
    // their plates overlap by a padding, so the two texts are one padding apart
    GapPixels: 32,
    LineGapPixels: -8,

    // the text never squeezes the figure below this share of the canvas: of its width beside
    // the caption, of its height above it
    FigureShare: 0.4,
    StackedFigureShare: 1 / 3,

    // between two show/hide boxes in a row
    BoxRowGapPixels: 16,

    hasCaption(drawing) {
        return GalleryDrawing.findCaption(drawing) != null;
    },

    findCaption(drawing) {
        const title = drawing.figures.byName(GalleryDrawing.TitleName);
        const description = drawing.figures.byName(GalleryDrawing.DescriptionName);
        return title instanceof Label && description instanceof Label ? { title, description } : null;
    },

    /** A drawing about coordinates (a graph on the grid) has no bounds to fit: what it shows is the part of the plane its file says. Null for all other drawings. */
    getPlane(drawing) {
        return drawing.coordinateGrid.visible ? drawing.viewport : null;
    },

    /**
     * Zoom to fit, with the caption where it can't cover the figure. The text has a size in
     * pixels whatever the zoom, so the figure gets the canvas minus the text, and the zoom
     * follows from that directly. In an element too small for both, the figure keeps a share
     * of the canvas and the text runs off the bottom edge.
     */
    fit(drawing, plane, stackedFigureShare = GalleryDrawing.StackedFigureShare) {
        const coordinateSystem = drawing.coordinateSystem;
        const canvas = drawing.canvas;
        const canvasWidth = canvas != null ? canvas.width : 0;
        const canvasHeight = canvas != null ? canvas.height : 0;

        // nothing to fit into; refitted when it grows
        if (canvasWidth <= 0 || canvasHeight <= 0) {
            return;
        }

        const hasScene = drawing.scenes.length > 0;
        const caption = GalleryDrawing.findCaption(drawing);
        if (caption == null) {
            if (hasScene) {
                drawing.showScene(drawing.chooseScene(canvasWidth, canvasHeight));
            } else {
                coordinateSystem.zoomExtend(plane);
            }

            return;
        }

        const title = caption.title;
        const description = caption.description;
        const hintFigure = drawing.figures.byName(GalleryDrawing.HintName);
        const hint = hintFigure instanceof Label ? hintFigure : null;
        const isCaption = figure => figure === title || figure === description || figure === hint;

        // the geometry a show/hide box brings up gets room from the start; not its text,
        // which is sized in pixels and has no size in the plane to make room for
        const revealable = new Set();
        for (const figure of drawing.figures.list) {
            if (figure instanceof ShowHideControl) {
                for (const dependency of figure.dependencies) {
                    if (!(dependency instanceof ControlBase)) {
                        revealable.add(dependency);
                    }
                }
            }
        }

        const tryGetFigure = (withText = true) => {
            let bounds = coordinateSystem.tryGetContentBounds(
                figure => !isCaption(figure) && (withText || !(figure instanceof ControlBase)),
                figure => revealable.has(figure));
            if (plane != null) {
                bounds = bounds != null ? bounds.union(plane) : plane;
            }

            return bounds;
        };

        // the text of the figure keeps its size in pixels, so it is measured at a zoom the
        // element alone decides: the figure without its text across the whole canvas
        const bare = hasScene ? null : tryGetFigure(false);
        if (bare != null && (bare.width > 0 || bare.height > 0)) {
            const zoom = Math.min(
                bare.width > 0 ? canvasWidth / bare.width : Number.MAX_VALUE,
                bare.height > 0 ? canvasHeight / bare.height : Number.MAX_VALUE);
            coordinateSystem.setView(bare.center, CoordinateSystem.clampUnitLength(Math.min(zoom, CoordinateSystem.MaxFitUnitLength)));
        }

        // a drawing with scenes shows the scene, not its content (ground goes on forever)
        let figure = null;
        if (!hasScene) {
            figure = tryGetFigure();
            if (figure == null) {
                coordinateSystem.zoomExtend();
                return;
            }
        }

        const margin = CoordinateSystem.FitMarginPixels;
        const textMargin = GalleryDrawing.TextMarginPixels;
        const gap = GalleryDrawing.GapPixels;
        const lineGap = GalleryDrawing.LineGapPixels;

        // beside the figure when there is room for a column, else under it
        let isWide = canvasWidth >= canvasHeight;
        let column = 0;
        if (isWide) {
            title.wrapWidth = 0;
            title.backdrop = true;
            const available = canvasWidth - textMargin - gap - margin - GalleryDrawing.FigureShare * canvasWidth;
            column = Math.min(Math.max(GalleryDrawing.CaptionWidth, title.measureSize().width), available);
            if (column < GalleryDrawing.MinimumCaptionWidth) {
                isWide = false;
            }
        }

        // the hint is set off as another paragraph of the explanation would be, by about a line
        const hintGap = hint != null ? lineGap + TextMeasurer.fontSize(hint.font) : 0;
        let room;
        if (isWide) {
            GalleryDrawing.setCaption(LabelPin.TopRight, column, title, description, hint);
            const titleSize = title.measureSize();
            const descriptionSize = description.measureSize();
            const hintHeight = hint != null ? hintGap + hint.measureSize().height : 0;
            const textHeight = titleSize.height + lineGap + descriptionSize.height + hintHeight;
            const top = Math.max(textMargin, (canvasHeight - textHeight) / 2);
            title.pinOffset = new Point(textMargin, top);
            description.pinOffset = new Point(textMargin, top + titleSize.height + lineGap);
            if (hint != null) {
                hint.pinOffset = new Point(textMargin, top + titleSize.height + lineGap + descriptionSize.height + hintGap);
            }

            room = new Rect(margin, margin, canvasWidth - margin - gap - column - textMargin - margin, canvasHeight - 2 * margin);
        } else {
            GalleryDrawing.setCaption(LabelPin.BottomLeft, canvasWidth - 2 * textMargin, title, description, hint);
            const titleSize = title.measureSize();
            const descriptionSize = description.measureSize();
            const hintHeight = hint != null ? hintGap + hint.measureSize().height : 0;
            const textHeight = titleSize.height + lineGap + descriptionSize.height + hintHeight;
            const roomHeight = Math.max(canvasHeight - margin - gap - textHeight - textMargin, stackedFigureShare * canvasHeight);
            const textTop = margin + roomHeight + gap;

            // pinned at the bottom: an offset is from there to the label's lower edge
            title.pinOffset = new Point(textMargin, canvasHeight - textTop - titleSize.height);
            description.pinOffset = new Point(textMargin, canvasHeight - textTop - titleSize.height - lineGap - descriptionSize.height);
            if (hint != null) {
                hint.pinOffset = new Point(textMargin, canvasHeight - textTop - textHeight);
            }

            room = new Rect(margin, margin, canvasWidth - 2 * margin, roomHeight);
        }

        room = GalleryDrawing.layOutPinnedBoxes(drawing, room, figure, hasScene, margin);
        if (room.width <= 0 || room.height <= 0) {
            GalleryDrawing.updateCaption(title, description, hint);
            return;
        }

        // the scene nearest in shape to the room: landscape or portrait
        if (hasScene) {
            figure = drawing.chooseScene(room.width, room.height);
            drawing.activeScene = figure;
        }

        // the zoom that fills the room; a figure with no size keeps the zoom it has
        let unitLength = coordinateSystem.unitLength;
        if (figure.width > 0 || figure.height > 0) {
            unitLength = Math.min(
                figure.width > 0 ? room.width / figure.width : Number.MAX_VALUE,
                figure.height > 0 ? room.height / figure.height : Number.MAX_VALUE);
            if (!hasScene) {
                // a big point (an emoji) would stick out of the room by half its size
                unitLength = coordinateSystem.limitZoomByPointReach(unitLength, figure.center, room, f => !isCaption(f));
            }

            unitLength = Math.min(unitLength, CoordinateSystem.MaxFitUnitLength);
        }

        coordinateSystem.setView(figure.center, CoordinateSystem.clampUnitLength(unitLength), room.center);
        GalleryDrawing.updateCaption(title, description, hint);
    },

    /** The labels where their pins and offsets now say (a view change updates them too; a layout alone doesn't) */
    updateCaption(...labels) {
        for (const label of labels) {
            if (label != null) {
                label.updateVisual();
            }
        }
    },

    /**
     * The show/hide boxes pinned in the top left corner, laid out where they leave the
     * figure the most room: one under another beside it, or in a row over it. Returns the
     * room left for the figure.
     */
    layOutPinnedBoxes(drawing, room, figure, hasScene, margin) {
        const boxes = drawing.figures.list.filter(box => box instanceof ShowHideControl && box.pin === LabelPin.TopLeft && box.visible);
        if (boxes.length === 0) {
            return room;
        }

        const textMargin = GalleryDrawing.TextMarginPixels;
        const sizes = boxes.map(box => box.measureSize());
        const left = Math.max(room.x, textMargin + Math.max(...sizes.map(size => size.width)) + margin);
        const top = Math.max(room.y, textMargin + Math.max(...sizes.map(size => size.height)) + margin);
        const rowWidth = sizes.reduce((sum, size) => sum + size.width, 0) + GalleryDrawing.BoxRowGapPixels * (boxes.length - 1);
        const beside = new Rect(left, room.y, Math.max(0, room.right - left), room.height);
        const under = new Rect(room.x, top, room.width, Math.max(0, room.bottom - top));
        const zoom = candidate => {
            if (candidate.width <= 0 || candidate.height <= 0) {
                return 0;
            }

            const shown = hasScene ? drawing.chooseScene(candidate.width, candidate.height) : figure;
            return Math.min(candidate.width / shown.width, candidate.height / shown.height);
        };

        const inRow = textMargin + rowWidth <= room.right && zoom(under) > zoom(beside);
        let x = textMargin;
        let y = textMargin;
        for (let i = 0; i < boxes.length; i++) {
            boxes[i].pinOffset = new Point(x, y);
            boxes[i].updateVisual();
            if (inRow) {
                x += sizes[i].width + GalleryDrawing.BoxRowGapPixels;
            } else {
                y += sizes[i].height;
            }
        }

        return inRow ? under : beside;
    },

    setCaption(pin, width, ...labels) {
        for (const label of labels) {
            if (label != null) {
                label.pin = pin;
                label.wrapWidth = width;
                label.backdrop = true;
            }
        }
    }
};
