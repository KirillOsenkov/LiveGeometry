// Port of Main/Avalonia/DynamicGeometry/Figures/AngleSweep.cs: which of the two angles
// between two rays out of a point a figure means (an angle's mark and number, a bisector,
// an arc with its sector and segment). With θ the counterclockwise angle from the first ray
// to the second, the regions measure θ and 2π - θ: Counterclockwise and Clockwise are fixed
// choices, Smaller and Larger trade places at 180°, where the counterclockwise one is taken.

const AngleSweep = Object.freeze({
    Counterclockwise: "Counterclockwise",
    Clockwise: "Clockwise",
    Smaller: "Smaller",
    Larger: "Larger",

    /** Whether the region goes clockwise from the first ray to the second, given the counterclockwise angle between them (0 to 2π) */
    isClockwise(sweep, counterclockwise) {
        switch (sweep) {
            case AngleSweep.Clockwise:
                return true;
            case AngleSweep.Smaller:
                return counterclockwise > Math.PI;
            case AngleSweep.Larger:
                return counterclockwise < Math.PI;
            default:
                return false;
        }
    },

    /** The measure of the region, 0 to 2π; two rays along one line measure 0 whichever way round */
    measure(sweep, counterclockwise) {
        return AngleSweep.isClockwise(sweep, counterclockwise) && counterclockwise > 0
            ? 2 * Math.PI - counterclockwise
            : counterclockwise;
    },

    /** The attribute's value read back by name; the given default when it is missing or unknown */
    read(element, defaultSweep) {
        const text = element.getAttribute("Sweep");
        return text != null && AngleSweep[text] != null && typeof AngleSweep[text] === "string" ? AngleSweep[text] : defaultSweep;
    }
});
