// Port of Main/Avalonia/DynamicGeometry/Figures/Lines/SegmentDecoration.cs: the school marks on
// a segment (ticks, chevrons, a wave), drawn at its middle in its own stroke, sized in pixels

const SegmentDecoration = {
    None: "None",
    OneTick: "OneTick",
    TwoTicks: "TwoTicks",
    ThreeTicks: "ThreeTicks",
    OneArrow: "OneArrow",
    TwoArrows: "TwoArrows",
    ThreeArrows: "ThreeArrows",
    Wave: "Wave"
};

const SegmentDecorationMark = {
    TickLength: 10,
    Spacing: 4,
    ArrowLength: 7,
    ArrowHalfWidth: 5,
    WaveHalfLength: 12.5,
    WaveHeight: 5.5,

    /** start, end: the segment's ends in pixels (chevrons point to the end); stroke: the segment's */
    render(renderer, start, end, stroke, decoration) {
        const along = RightAngleMark.direction(start, end);
        if (decoration === SegmentDecoration.None || along == null || stroke == null) {
            return;
        }

        const middle = new Point((start.x + end.x) / 2, (start.y + end.y) / 2);
        const roundStroke = { color: stroke.color, width: stroke.width, dash: null, cap: "round", join: "round" };
        for (const polyline of SegmentDecorationMark.createGeometry(decoration, middle, along, stroke.width)) {
            renderer.drawPolyline(polyline, roundStroke);
        }
    },

    /** The mark's lines, in pixels: at the middle of a segment running along the unit vector, sized for a stroke of the given width */
    createGeometry(decoration, middle, along, thickness) {
        const across = new Point(-along.y, along.x);
        const figures = [];
        const grow = Math.max(0, thickness - 1);
        const spacing = SegmentDecorationMark.Spacing + thickness;
        const ticks = { OneTick: 1, TwoTicks: 2, ThreeTicks: 3 }[decoration];
        const arrows = { OneArrow: 1, TwoArrows: 2, ThreeArrows: 3 }[decoration];
        if (ticks != null) {
            const half = SegmentDecorationMark.TickLength / 2 + grow / 2;
            for (let i = 0; i < ticks; i++) {
                const center = middle.plus(along.scale((i - (ticks - 1) / 2) * spacing));
                figures.push([center.minus(across.scale(half)), center.plus(across.scale(half))]);
            }
        } else if (arrows != null) {
            const length = SegmentDecorationMark.ArrowLength + grow;
            const half = SegmentDecorationMark.ArrowHalfWidth + grow / 2;
            for (let i = 0; i < arrows; i++) {
                const tip = middle.plus(along.scale((i - (arrows - 1) / 2) * spacing + length / 2));
                const back = tip.minus(along.scale(length));
                figures.push([back.plus(across.scale(half)), tip, back.minus(across.scale(half))]);
            }
        } else if (decoration === SegmentDecoration.Wave) {
            const halfLength = SegmentDecorationMark.WaveHalfLength + grow;
            const height = SegmentDecorationMark.WaveHeight + grow / 2;
            const steps = 16;
            const points = [];
            for (let i = 0; i <= steps; i++) {
                const t = -1 + 2 * i / steps;
                points.push(middle.plus(along.scale(t * halfLength)).plus(across.scale(height * Math.sin(Math.PI * t))));
            }

            figures.push(points);
        }

        return figures;
    }
};
