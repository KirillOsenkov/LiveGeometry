// Port of Main/Avalonia/DynamicGeometry/Figures/Points/PointMarker.cs: the visual of a
// point, a PointShape filling its size, centered on the point, or a character (an emoji)
// instead. Here it is what a renderer draws (Renderer.drawMarker), worked out as a path.

const PointMarker = {
    /** How far the corners of each polygon reach, in radii of the circle, so they look about as big */
    Reach: {
        Triangle: 1.3,
        Square: 1.2,
        Diamond: 1.25,
        Pentagon: 1.15,
        Hexagon: 1.1
    },

    /** The corners of the marker's polygon around the center, in pixels; null for a circle */
    getPolygon(kind, center, size, strokeThickness) {
        const radius = Math.max(0, size / 2 - strokeThickness / 2);
        const reach = PointMarker.Reach[kind];
        switch (kind) {
            case PointShape.Triangle:
                return PointMarker.regularPolygon(center, radius * reach, 3, -90);
            case PointShape.Square:
                return PointMarker.regularPolygon(center, radius * reach, 4, 45);
            case PointShape.Diamond:
                return PointMarker.regularPolygon(center, radius * reach, 4, -90);
            case PointShape.Pentagon:
                return PointMarker.regularPolygon(center, radius * reach, 5, -90);
            case PointShape.Hexagon:
                return PointMarker.regularPolygon(center, radius * reach, 6, 0);
            default:
                return null;
        }
    },

    /** The radius of the circle marker, in pixels */
    getRadius(size, strokeThickness) {
        return Math.max(0, size / 2 - strokeThickness / 2);
    },

    /** startAngle: where the first corner is, in degrees clockwise from the right (y down) */
    regularPolygon(center, radius, sides, startAngle) {
        const corners = [];
        for (let i = 0; i < sides; i++) {
            const angle = (startAngle + 360 * i / sides) * Math.PI / 180;
            corners.push(new Point(center.x + radius * Math.cos(angle), center.y + radius * Math.sin(angle)));
        }

        return corners;
    }
};
