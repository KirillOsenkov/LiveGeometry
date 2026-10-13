// Port of Main/Avalonia/DynamicGeometry/Figures/ZOrder.cs: the layers figures are drawn in,
// and what a figure's layer and Z (0 unless the app's Bring to front or Send to back changed
// it, saved as Z) make of its zIndex.

const ZOrder = {
    Grid: 10,
    Axes: 20,
    Polygons: 30,
    Labels: 40,
    Figures: 50,
    Vectors: 55,
    SelectionHalos: 58,
    Handles: 59,
    Points: 60,
    PointLabels: 70,
    Controls: 80
};

const ZOrders = {
    /** How far a Z may go either way */
    Range: 999,

    /** One step of Z is this many zIndex values: room for every layer of the band */
    Stride: 1000,

    /**
     * The band whose Z orders figures of the layer, named by its lowest layer: the figures'
     * (Polygons to Vectors), the points' (Points), or null
     */
    band(layer) {
        if (layer >= ZOrder.Polygons && layer <= ZOrder.Vectors) {
            return ZOrder.Polygons;
        }

        if (layer === ZOrder.Points) {
            return ZOrder.Points;
        }

        return null;
    },

    /** Whether a figure of the layer can be brought to front and sent to back: it is in one of the bands */
    isMovable(layer) {
        return ZOrders.band(layer) != null;
    },

    /** Whether the Z of figures of the two layers orders them among each other: a Z is a place in its own band only */
    areInSameBand(layer, other) {
        return ZOrders.isMovable(layer) && ZOrders.band(layer) === ZOrders.band(other);
    },

    /**
     * The zIndex of the layer at the Z, bottom up: the layers below the figures' band as
     * they are, the figures' band ordered by Z and then by layer, the selection halos and the
     * handles, the points' band ordered by Z, the layers above it over all of it
     */
    encode(layer, z) {
        const span = ZOrders.Stride * (2 * ZOrders.Range + 1);
        const figureBand = ZOrders.Stride;
        const pointBand = figureBand + span + ZOrders.Stride;
        const place = (Math.min(ZOrders.Range, Math.max(-ZOrders.Range, z)) + ZOrders.Range) * ZOrders.Stride;
        if (layer < ZOrder.Polygons) {
            return layer;
        }

        if (layer <= ZOrder.Vectors) {
            return figureBand + place + layer;
        }

        if (layer < ZOrder.Points) {
            return figureBand + span + layer;
        }

        if (layer === ZOrder.Points) {
            return pointBand + place + layer;
        }

        return pointBand + span + layer;
    },

    /**
     * A part is drawn where its figure is: it takes the figure's Z when the two are in one
     * band, and keeps its own otherwise (a regular polygon's vertices, a slider's track)
     */
    passOn(figure, part) {
        if (ZOrders.areInSameBand(figure.layer, part.layer)) {
            part.z = figure.z;
        }
    }
};
