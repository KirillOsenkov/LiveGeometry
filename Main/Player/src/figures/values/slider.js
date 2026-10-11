// Port of Main/Avalonia/DynamicGeometry/Figures/Values/Slider.cs: an adjustable number with
// a handle, one composite whose parts are figures of the library: a free point anchor, a
// knob kept on the horizontal through it, the segment track, a caption "a = 2.00".

class Slider extends CompositeFigure {
    constructor() {
        super();

        // the knob's direction: an angle provider saying 0 degrees; not a child, it has no shape
        this.horizontal = new NumberFigure();
        this.horizontal.value = 0;

        // the parts are looked up by name only through the composite, so their names must not
        // be ones an expression could say
        this.anchor = new FreePoint();
        this.anchor.name = "slider anchor";
        this.knob = new SliderKnob(this);
        this.knob.name = "slider knob";
        this.minimum = 0;
        this.maximum = Infinity;
        this.knob.setSources(this.anchor, null, this.horizontal);
        this.track = new Segment();
        this.track.name = "slider track";
        this.track.dependencies = [this.anchor, this.knob];
        this.caption = new SliderCaption(this);
        this.caption.name = "slider caption";
        this.caption.dependencies = [this.anchor, this.knob];
        this.addChild(this.anchor);
        this.addChild(this.knob);
        this.addChild(this.track);
        this.addChild(this.caption);
        this.wholeHandle = new SliderWholeHandle(this);

        // of the figures under the cursor the topmost wins: the knob over a line crossing it
        this.zIndex = ZOrder.Points;
    }

    get isSlider() {
        return true;
    }

    get isNumber() {
        return true;
    }

    get isLengthProvider() {
        return true;
    }

    get isAngleProvider() {
        return true;
    }

    get isMovableParts() {
        return true;
    }

    // Value and place

    /** The number the slider stands for: the minimum at the anchor, one more per unit of track, up to the maximum. Set it and the knob moves; anything built on the slider follows. */
    get value() {
        return this.minimum + this.knob.distance;
    }

    set value(value) {
        const origin = this.anchor.coordinates;
        const clamped = Math.max(this.minimum, Math.min(this.maximum, value));
        this.knob.moveToCore(new Point(origin.x + clamped - this.minimum, origin.y));
        this.onChanged();
    }

    /** How far the knob may go from the anchor */
    get span() {
        return this.maximum - this.minimum;
    }

    /** Where the anchor is; the knob keeps its distance */
    get position() {
        return this.anchor.coordinates;
    }

    set position(value) {
        this.anchor.moveToCore(value);
        this.onChanged();
    }

    onChanged() {
        if (this.drawing != null) {
            this.recalculateAllDependents();
        }
    }

    get decimals() {
        return this.caption.decimalsToShow;
    }

    set decimals(value) {
        this.caption.decimalsToShow = value;
    }

    get length() {
        return this.value;
    }

    /** Angle providers speak radians; the slider itself is in degrees, like a Number */
    get angle() {
        return GeometryMath.toRadians(this.value);
    }

    get center() {
        return this.track.center;
    }

    /** a, b, c: what an expression says, next to points A, B, C */
    generateFigureName() {
        const letters = "abcdfghkmnpqrstuvwz";
        for (let i = 0; ; i++) {
            for (const letter of letters) {
                const candidate = i === 0 ? letter : letter + i;
                if (this.nameAvailable(candidate)) {
                    return candidate;
                }
            }
        }
    }

    toString() {
        return this.name;
    }

    get name() {
        return super.name;
    }

    set name(value) {
        super.name = value;
        if (this.drawing != null) {
            this.caption.updateVisual();
        }
    }

    /** The track's line style stands for the slider's */
    get style() {
        return this.track.style;
    }

    set style(value) {
        this.track.style = value;
    }

    onAddingToCanvas(newContainer) {
        // before the parts take the styles of their kinds
        this.ensureStyleAssigned();
        super.onAddingToCanvas(newContainer);
    }

    // Hit testing and dragging

    /** A hit on any part is a hit on the slider */
    hitTestWith(point, filter) {
        return this.findPart(point) != null && filter(this) ? this : null;
    }

    hitTest(point) {
        return this.hitTestWith(point, f => f.visible);
    }

    /** The visible part under the point; the knob first, it sits on the track and on the anchor at zero */
    findPart(point) {
        for (const part of [this.knob, this.anchor, this.caption, this.track]) {
            if (part.visible && part.hitTest(point) != null) {
                return part;
            }
        }

        return null;
    }

    /** The knob changes the value, the anchor and everything else move the slider */
    findMovablePart(point) {
        const part = this.findPart(point);
        if (part === this.knob) {
            return this.knob;
        }

        if (part === this.anchor) {
            return this.anchor;
        }

        return this.wholeHandle;
    }

    get wholePart() {
        return this.wholeHandle;
    }

    readXml(element) {
        super.readXml(element);
        if (element.hasAttribute("Decimals")) {
            this.decimals = Math.trunc(Xml.readDouble(element, "Decimals"));
        }

        // the range before the value, which is clamped to it
        if (element.hasAttribute("Minimum")) {
            this.minimum = Xml.readDouble(element, "Minimum");
        }

        if (element.hasAttribute("Maximum")) {
            this.maximum = Xml.readDouble(element, "Maximum");
        }

        this.anchor.moveToCore(new Point(Xml.readDouble(element, "X"), Xml.readDouble(element, "Y")));
        this.value = Xml.readDouble(element, "Value");
    }
}

/** Dragging the track or the caption: the anchor moves by as much as the cursor, without jumping under it */
class SliderWholeHandle {
    constructor(slider) {
        this.slider = slider;
    }

    get isMovable() {
        return true;
    }

    get coordinates() {
        return this.slider.anchor.coordinates;
    }

    allowMove() {
        return !this.slider.locked;
    }

    moveTo(position) {
        this.slider.anchor.moveTo(position);
    }
}

/** The knob: a point sliding along the horizontal through the anchor, never to its left, nor past the slider's span */
class SliderKnob extends TranslatedPoint {
    constructor(slider) {
        super();
        this.slider = slider;
    }

    moveToCore(newPosition) {
        const source = this.source;
        if (source != null) {
            const left = source.coordinates.x;
            const right = left + this.slider.span;
            if (newPosition.x < left) {
                newPosition = new Point(left, newPosition.y);
            } else if (newPosition.x > right) {
                newPosition = new Point(right, newPosition.y);
            }
        }

        super.moveToCore(newPosition);
    }
}

/** "a = 2.00" just above the anchor, left-aligned with it. Not draggable on its own. */
class SliderCaption extends LabelWithOffset {
    static gap = 1;

    constructor(slider) {
        super();
        this.slider = slider;
    }

    get anchor() {
        return this.slider.anchor.coordinates;
    }

    allowMove() {
        return false;
    }

    updateVisual() {
        if (this.drawing == null) {
            return;
        }

        this.setText(NameDisplay.format(this.slider.name) + " = " + NumberFormat.toString(GeometryMath.round(this.slider.value, this.decimalsToShow)));
        const size = this.measureSize();

        // left-aligned with the anchor, and above the bigger of the two points
        const anchorRadius = this.slider.anchor.pointSize / 2;
        const knobWidth = this.slider.knob.pointSize;
        const pointRadius = isValidValue(knobWidth) ? Math.max(anchorRadius, knobWidth / 2) : anchorRadius;
        this.offset = new Point(-anchorRadius, -(size.height + pointRadius + SliderCaption.gap));
        super.updateVisual();
    }
}

FigureTypes.register("Slider", Slider);
