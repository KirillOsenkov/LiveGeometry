// Port of Main/Avalonia/DynamicGeometry/Figures/Controls/LabelBase.cs: text on the paper. Its
// [...] parts are expressions, compiled when the text is set, whose values the text shows.
// The TextBlock of the C# is a text layout the canvas measures (textLayout). Left out: the
// 300 ms throttle of the shown text while figures move (the render loop coalesces frames),
// the rename of expressions.

class LabelBase extends ControlBase {
    /** What an expression that has no value shows */
    static UndefinedText = "undefined";

    static squareBrackets = /\[[^\[\]]*\]/g;

    constructor() {
        super();
        this.text = "";
        this.processedText = "";
        this.shouldProcessText = false;
        this.textChunks = null;
        this.embeddedExpressions = null;
        this.decimalsToShowValue = Settings.displayDecimals;
        this.textLayout = null;
        this.padding = 0;
        this.wrapWidthValue = 0;
    }

    get isLabel() {
        return true;
    }

    defaultZOrder() {
        return ZOrder.Labels;
    }

    get textValue() {
        return this.text;
    }

    /** The text as typed (Text); setting it compiles the [...] parts */
    setText(value) {
        this.text = value ?? "";
        if (this.shouldProcessText) {
            this.processText();
        } else {
            this.processedText = this.text;
            this.textLayout = null;
        }
    }

    /** The font the text is drawn in, from the style on screen */
    get font() {
        const style = this.resolvedStyle;
        return style != null && style.toCanvasFont != null ? style.toCanvasFont() : Fonts.family;
    }

    get textColor() {
        return this.resolvedStyle?.color ?? Color.black;
    }

    /** The text laid out at the font and wrap width on screen (the TextBlock's layout), measured by the canvas */
    getTextLayout() {
        const canvas = this.canvas;
        if (canvas == null) {
            return null;
        }

        const font = this.font;
        const text = this.processedText ?? "";
        if (this.textLayout == null || this.textLayout.text !== text || this.textLayout.font !== font || this.textLayout.wrapWidth !== this.wrapWidthValue) {
            this.textLayout = canvas.measureText(text, font, this.wrapWidthValue);
        }

        return this.textLayout;
    }

    /** The size the text takes right now, padding included, laid out here and now */
    measureSize() {
        const layout = this.getTextLayout();
        if (layout == null) {
            return new Size(0, 0);
        }

        return new Size(layout.width + 2 * this.padding, layout.height + 2 * this.padding);
    }

    processText() {
        if (!this.shouldProcessText) {
            return;
        }

        const text = this.text;
        this.unregisterFromDependencies();
        this.mDependencies = [];
        this.embeddedExpressions = [];
        this.textChunks = [];
        let chunkStart = 0;
        for (const match of text.matchAll(LabelBase.squareBrackets)) {
            const chunkEnd = match.index;
            this.textChunks.push(chunkEnd > chunkStart ? text.substring(chunkStart, chunkEnd) : "");
            this.processMatch(match[0]);
            chunkStart = match.index + match[0].length;
        }

        this.textChunks.push(text.length > chunkStart ? text.substring(chunkStart) : "");
        this.onDependenciesChanged();
        this.registerWithDependencies();
        this.recalculate();
    }

    recalculate() {
        if (!this.shouldProcessText) {
            return;
        }

        if (this.text === "") {
            this.processedText = "";
            this.textLayout = null;
            return;
        }

        if (this.textChunks == null || this.embeddedExpressions == null) {
            this.processText();
        }

        let result = "";
        for (let i = 0; i < this.textChunks.length; i++) {
            if (i !== 0) {
                const compileResult = this.embeddedExpressions[i - 1];
                if (compileResult.isSuccess) {
                    const value = compileResult.expression();
                    result += !isValidValue(value) ? LabelBase.UndefinedText : this.formatNumber(value);
                } else {
                    result += compileResult.toString();
                }
            }

            result += this.textChunks[i];
        }

        this.processedText = result;
    }

    /** A number of a text label's [...] part, with every decimal it shows, zeros at the end included (3.70, 9.00) */
    formatNumber(value) {
        return NumberFormat.toFixed(GeometryMath.round(value, this.decimalsToShow), this.decimalsToShow);
    }

    rebindExpressions() {
        if (this.shouldProcessText && this.text !== "") {
            this.processText();
        }
    }

    processMatch(match) {
        if (match.length < 3) {
            const error = new CompileResult();
            error.addError("Empty expression");
            this.embeddedExpressions.push(error);
            return;
        }

        const expression = match.substring(1, match.length - 1);
        const compileResult = Compiler.instance.compileExpression(this.drawing, expression, figure => !figure.dependsOn(this));
        this.embeddedExpressions.push(compileResult);
        if (compileResult.isSuccess) {
            // once each
            for (const dependency of compileResult.dependencies) {
                if (!this.mDependencies.includes(dependency)) {
                    this.mDependencies.push(dependency);
                }
            }
        }
    }

    /** The value of a label that is one expression and nothing else ("[AB / 3]"), as it is and not as the label shows it; null for any other label */
    get exactValue() {
        if (!this.shouldProcessText
            || this.textChunks == null
            || this.embeddedExpressions == null
            || this.embeddedExpressions.length !== 1
            || this.textChunks.length !== 2
            || this.textChunks[0].trim() !== ""
            || this.textChunks[1].trim() !== ""
            || !this.embeddedExpressions[0].isSuccess) {
            return null;
        }

        const value = this.embeddedExpressions[0].expression();
        return isValidValue(value) ? value : NaN;
    }

    get decimalsToShow() {
        return this.decimalsToShowValue;
    }

    set decimalsToShow(value) {
        if (value >= 0 && value <= 10) {
            this.decimalsToShowValue = value;
            this.updateVisual();
        }
    }

    readXml(element) {
        super.readXml(element);
        // files from before the attribute existed: the default, not 0
        if (element.hasAttribute("DecimalsToShow")) {
            this.decimalsToShow = Math.trunc(Xml.readDouble(element, "DecimalsToShow"));
        }
    }

    render(renderer) {
        if (!this.isShown) {
            return;
        }

        const layout = this.getTextLayout();
        if (layout == null) {
            return;
        }

        const topLeft = this.toPhysical(this.coordinates);
        if (!topLeft.exists()) {
            return;
        }

        renderer.drawText(layout, topLeft, this.font, this.textColor, this.padding, this.backdropBrush, this.resolvedStyle?.underline === true);
    }

    /** The plate behind the text, if any (Label.Backdrop) */
    get backdropBrush() {
        return null;
    }
}
