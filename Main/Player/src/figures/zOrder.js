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

    /** Whether a figure of the layer can be brought to front and sent to back */
    isMovable(layer) {
        return layer >= ZOrder.Polygons && layer <= ZOrder.Vectors;
    },

    /**
     * The zIndex of the layer at the Z: the layers below the band as they are, the band
     * above them ordered by Z and then by layer, the layers above the band over all of it
     */
    encode(layer, z) {
        const span = ZOrders.Stride * (2 * ZOrders.Range + 1);
        if (layer < ZOrder.Polygons) {
            return layer;
        }

        if (layer > ZOrder.Vectors) {
            return ZOrders.Stride + span + layer;
        }

        const clamped = Math.min(ZOrders.Range, Math.max(-ZOrders.Range, z));
        return ZOrders.Stride + (clamped + ZOrders.Range) * ZOrders.Stride + layer;
    }
};
