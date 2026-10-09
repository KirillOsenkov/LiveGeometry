// Port of Main/Avalonia/DynamicGeometry/Settings.cs: what the player needs of it.

const Settings = {
    defaultUnitLength: 48,
    scaleTextWithDrawing: false,

    /** What a saved drawing says it is; 1 = label offsets in pixels */
    currentDrawingVersion: 1,

    /** How many decimals a number is shown with (a label's DecimalsToShow by default) */
    displayDecimals: 2,

    /** Whether a point on an ellipse keeps its absolute angle or one relative to the ellipse */
    pointsOnEllipticalsUseAbsoluteAngle: true,

    /** In pixels: how far from a figure a click still takes it (10 for a finger) */
    cursorTolerance: 5,

    autoLabelPoints: false,

    pointAlphabet: "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
};
