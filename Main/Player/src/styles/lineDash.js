// Port of Main/Avalonia/DynamicGeometry/Styles/LineDash.cs

/** The same five the original DG had, in the same order. The names are what goes into drawings. */
const LineDash = {
    Solid: "Solid",
    Dash: "Dash",
    Dot: "Dot",
    DashDot: "DashDot",
    DashDotDot: "DashDotDot"
};

const LineDashes = {
    /** Dash and gap lengths in pixels */
    getPattern(dash) {
        switch (dash) {
            case LineDash.Dash:
                return [8, 5];
            case LineDash.Dot:
                return [2, 4];
            case LineDash.DashDot:
                return [9, 4, 2, 4];
            case LineDash.DashDotDot:
                return [9, 4, 2, 4, 2, 4];
            default:
                return null;
        }
    },

    /**
     * The dash lengths in pixels (Avalonia counts them in stroke widths, a canvas in pixels;
     * what GetDashArray gives times the width): up to 3 px wide as given, beyond that they
     * grow with the line. Null for solid.
     */
    getDashArray(dash, strokeThickness) {
        const pattern = LineDashes.getPattern(dash);
        if (pattern == null) {
            return null;
        }

        const unit = Math.max(strokeThickness, 0.1);
        const stretch = Math.max(1, unit / 3);
        return pattern.map(length => length * stretch);
    },

    /** The style of a line in a .dgf file of the original DG */
    fromVB6DrawStyle(drawStyle) {
        return drawStyle >= 1 && drawStyle <= 4 ? Object.values(LineDash)[drawStyle] : LineDash.Solid;
    }
};
