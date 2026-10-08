// Port of Main/Avalonia/DynamicGeometry/Figures/Controls/LabelPin.cs

/** Where a label (or a show/hide box) is held: nowhere, or a corner of the canvas it keeps its distance from in pixels */
const LabelPin = {
    None: "None",
    TopLeft: "TopLeft",
    TopRight: "TopRight",
    BottomLeft: "BottomLeft",
    BottomRight: "BottomRight"
};

/** The pixels of a pin, for whatever can be pinned */
const Pinning = {
    /** The top-left corner of something of this size whose pinned corner sits offset pixels inward from the same corner of the canvas */
    topLeft(pin, offset, size, canvas) {
        const x = pin === LabelPin.TopLeft || pin === LabelPin.BottomLeft
            ? offset.x
            : canvas.x - offset.x - size.width;
        const y = pin === LabelPin.TopLeft || pin === LabelPin.TopRight
            ? offset.y
            : canvas.y - offset.y - size.height;
        return new Point(x, y);
    },

    /** The offset that puts something of this size at this top-left corner */
    offsetFrom(pin, topLeft, size, canvas) {
        const x = pin === LabelPin.TopLeft || pin === LabelPin.BottomLeft
            ? topLeft.x
            : canvas.x - topLeft.x - size.width;
        const y = pin === LabelPin.TopLeft || pin === LabelPin.TopRight
            ? topLeft.y
            : canvas.y - topLeft.y - size.height;
        return new Point(x, y);
    },

    /** The pin a file names, None for anything else */
    parse(name) {
        return name != null && Object.values(LabelPin).includes(name) ? name : LabelPin.None;
    }
};
