// Port of Main/Avalonia/DynamicGeometry/Figures/Lines/DirectionConstraint.cs, as far as the
// player goes: the direction of a segment, a ray or a line through two points, which an
// expression reads (AB.Direction). Typing, fixing and freeing it are the editor's.

const DirectionConstraint = {
    /** The direction of the line from its first point to its second, in degrees from 0 to 360 */
    of(line) {
        // a full turn but for rounding is no turn
        const angle = GeometryMath.toDegrees(GeometryMath.getAngle(line.p1, line.p2));
        return 360 - angle < 1e-9 ? 0 : angle;
    }
};
