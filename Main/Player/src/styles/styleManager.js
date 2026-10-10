// Port of Main/Avalonia/DynamicGeometry/Styles/StyleManager.cs: the styles of a drawing,
// the defaults a new drawing starts with (which a file names without carrying), and the
// default style of each kind of figure. Left out: creating, renaming and deleting styles
// (the editor's), the picker's IsOffered.

class StyleManager {
    // The names of the default styles: a figure whose style is the default of its kind is
    // written without one, and a file names a default it doesn't carry
    static FreePointStyleName = "FreePoint";
    static PointOnFigureStyleName = "PointOnFigure";
    static IntersectionPointStyleName = "IntersectionPoint";
    static MidpointStyleName = "Midpoint";
    static DependentPointStyleName = "DependentPoint";
    static LineStyleName = "Line";
    static SliderTrackStyleName = "SliderTrack";
    static HandleStyleName = "Handle";
    static ShapeStyleName = "Shape";
    static OutlinedShapeStyleName = "OutlinedShape";
    static TextStyleName = "Text";
    static HeadingStyleName = "Heading";
    static HyperlinkStyleName = "Hyperlink";
    static GalleryTitleStyleName = "GalleryTitle";
    static GalleryTextStyleName = "GalleryText";
    static GalleryLocusStyleName = "GalleryLocus";

    /** What files from before 2026-09-29 call the dependent point style */
    static LegacyDependentPointStyleName = "DependentPointStyle";

    static ThinStrokeWidth = 1.5;
    static ThickStrokeWidth = 2.5;

    /** The kinds of style a file may hold, by element name */
    static StyleTypes = [LineStyle, ShapeStyle, PointStyle, TextStyle];

    constructor(drawing) {
        this.list = [];
        this.aliases = new Map();
        this.numDefaultStyles = 0;
        this.addDefaultStyles();
        this.drawing = drawing;
    }

    onStyleAdded(style) {
        if (style.name === "") {
            style.name = this.generateUniqueName();
        }

        style.styleManager = this;
    }

    generateUniqueName() {
        let n = 1;
        while (this.get(String(n)) != null) {
            n++;
        }

        return String(n);
    }

    /** The style of that name (the indexer); null when there is none */
    get(name) {
        name = StyleManager.canonicalName(name);
        const target = this.aliases.get(name);
        if (target != null) {
            name = target;
        }

        for (const style of this.list) {
            if (style.name === name) {
                return style;
            }
        }

        return null;
    }

    /** The name a style goes by now, for the names older files use */
    static canonicalName(name) {
        return name === StyleManager.LegacyDependentPointStyleName ? StyleManager.DependentPointStyleName : name;
    }

    getAllStyles() {
        return this.list;
    }

    /**
     * The styles of a new drawing. Their colors are the theme's (the paper group of
     * AppTheme): the base value from the Light theme, an override from every other.
     */
    addDefaultStyles() {
        const freePointStyle = StyleManager.themedPoint(StyleManager.FreePointStyleName, 10, theme => theme.freePointFill);
        const pointOnFigureStyle = StyleManager.themedPoint(StyleManager.PointOnFigureStyleName, 10, theme => theme.pointOnFigureFill);
        const intersectionPointStyle = StyleManager.themedPoint(StyleManager.IntersectionPointStyleName, 8, theme => theme.intersectionPointFill);
        const midpointStyle = StyleManager.themedPoint(StyleManager.MidpointStyleName, 8, theme => theme.midpointFill);
        const dependentPointStyle = StyleManager.themedPoint(StyleManager.DependentPointStyleName, 8, theme => theme.dependentPointFill);
        const pointStyles = [
            dependentPointStyle,
            StyleManager.huePoint("RedPoint", "#FF7B7B", "#F07272"),
            midpointStyle,
            freePointStyle,
            pointOnFigureStyle,
            intersectionPointStyle,
            StyleManager.huePoint("BluePoint", "#7EA8F8", "#6E9EF5"),
            StyleManager.huePoint("PurplePoint", "#B98CF5", "#A77EF0")
        ];
        pointStyles.push(...StyleHue.All.map(StyleManager.bead));

        // the handles of a Bezier path: small gray squares
        const handleStyle = StyleManager.themedPoint(StyleManager.HandleStyleName, 7, theme => theme.dependentPointFill);
        handleStyle.shape = PointShape.Square;
        pointStyles.push(handleStyle);

        // Lines: thin, thick, and dashed. Gray is the theme's: the line every new figure gets
        const lineStyle = new LineStyle();
        lineStyle.name = StyleManager.LineStyleName;
        lineStyle.bindToTheme("color", theme => theme.line);
        const thickLineStyle = new LineStyle();
        thickLineStyle.name = "ThickLine";
        thickLineStyle.strokeWidth = StyleManager.ThickStrokeWidth;
        thickLineStyle.bindToTheme("color", theme => AppTheme.withAlpha(theme.ink, 230));
        const dashedLineStyle = StyleManager.hueLine("DashedLine", StyleHue.Gray, 1.25, LineDash.Dash);
        const lineStyles = [lineStyle];
        lineStyles.push(...StyleHue.Colors.map(hue => StyleManager.hueLine(hue.name + "Line", hue, StyleManager.ThinStrokeWidth, LineDash.Solid)));
        lineStyles.push(thickLineStyle);
        lineStyles.push(...StyleHue.Colors.map(hue => StyleManager.hueLine("Thick" + hue.name + "Line", hue, StyleManager.ThickStrokeWidth, LineDash.Solid)));
        lineStyles.push(dashedLineStyle);
        lineStyles.push(...StyleHue.Colors.map(hue => StyleManager.hueLine("Dashed" + hue.name + "Line", hue, StyleManager.ThinStrokeWidth, LineDash.Dash)));

        // a bar rather than a line: what the knob of a slider runs along
        const sliderTrackStyle = new LineStyle();
        sliderTrackStyle.name = StyleManager.SliderTrackStyleName;
        sliderTrackStyle.strokeWidth = 6;
        sliderTrackStyle.bindToTheme("color", theme => theme.sliderTrack);
        lineStyles.push(sliderTrackStyle);

        // Shapes: outlined with a flat fill, outlined with a gradient, and gradients without an outline
        const shapeStyle = new ShapeStyle();
        shapeStyle.name = StyleManager.ShapeStyleName;
        shapeStyle.color = Color.transparent;
        shapeStyle.bindToTheme("fill", theme => new SolidColorBrush(theme.shapeFill));
        const greenShapeStyle = new ShapeStyle();
        greenShapeStyle.name = "GreenShape";
        greenShapeStyle.color = Color.transparent;
        greenShapeStyle.fill = new SolidColorBrush(Color.fromArgb(100, 200, 255, 200));
        greenShapeStyle.setOverride(AppTheme.Dark.name, "fill", new SolidColorBrush(Color.fromArgb(100, 128, 200, 128)));
        const shapeStyles = [];
        shapeStyles.push(...StyleHue.All.map(hue => StyleManager.hueShape(
            hue === StyleHue.Gray ? StyleManager.OutlinedShapeStyleName : hue.name + "Outline",
            hue,
            true,
            false)));
        shapeStyles.push(...StyleHue.All.map(hue => StyleManager.hueShape("Gradient" + hue.name + "Outline", hue, true, true)));
        shapeStyles.push(...StyleHue.All.map(hue =>
            hue === StyleHue.Brown ? shapeStyle
                : hue === StyleHue.Green ? greenShapeStyle
                    : StyleManager.hueShape(hue.name + "Shape", hue, false, true)));

        const hyperLinkStyle = StyleManager.themedText(StyleManager.HyperlinkStyleName, 18);
        const textStyle = StyleManager.themedText(StyleManager.TextStyleName, 18);
        const headerStyle = StyleManager.themedText(StyleManager.HeadingStyleName, 40);

        // the caption of a drawing of the gallery: a heading in the splash's blue, the
        // explanation in the chrome's text color, the locus of a "drag to here" ring
        const galleryTitleStyle = new TextStyle();
        galleryTitleStyle.name = StyleManager.GalleryTitleStyleName;
        galleryTitleStyle.fontSize = 30;
        galleryTitleStyle.color = Color.fromRgb(0x1F, 0x4E, 0x8C);
        galleryTitleStyle.fontFamily = "Segoe UI";
        galleryTitleStyle.bold = true;
        galleryTitleStyle.setOverride(AppTheme.Dark.name, "color", Color.fromRgb(0x9C, 0xC4, 0xF0));
        const galleryTextStyle = new TextStyle();
        galleryTextStyle.name = StyleManager.GalleryTextStyleName;
        galleryTextStyle.fontSize = 16;
        galleryTextStyle.fontFamily = "Segoe UI";
        galleryTextStyle.bindToTheme("color", theme => theme.text);
        const galleryLocusStyle = new LineStyle();
        galleryLocusStyle.name = StyleManager.GalleryLocusStyleName;
        galleryLocusStyle.color = Color.fromRgb(0xE0, 0x36, 0x2B);
        galleryLocusStyle.strokeWidth = 2.5;

        const newStyles = [
            ...pointStyles,
            ...lineStyles,
            ...shapeStyles,
            textStyle,
            headerStyle,
            hyperLinkStyle,
            galleryTitleStyle,
            galleryTextStyle,
            galleryLocusStyle
        ];
        for (const style of newStyles) {
            this.add(style);
        }

        this.numDefaultStyles = newStyles.length;
    }

    /** A point style filled with a theme color, rimmed with the theme's ink */
    static themedPoint(name, size, fill) {
        const style = new PointStyle();
        style.name = name;
        style.size = size;
        style.bindToTheme("fill", theme => new SolidColorBrush(fill(theme)));
        style.bindToTheme("color", theme => AppTheme.withAlpha(theme.ink, 100));
        return style;
    }

    /** A point style of a color of its own on each paper, rimmed with the theme's ink, as small as the constructed ones */
    static huePoint(name, fill, darkFill) {
        const style = new PointStyle();
        style.name = name;
        style.size = 8;
        style.fill = new SolidColorBrush(Color.parse(fill));
        style.setOverride(AppTheme.Dark.name, "fill", new SolidColorBrush(Color.parse(darkFill)));
        style.bindToTheme("color", theme => AppTheme.withAlpha(theme.ink, 100));
        return style;
    }

    /** A big point with a highlight, for the points that matter */
    static bead(hue) {
        const style = new PointStyle();
        style.name = hue.name + "Bead";
        style.size = 12;
        style.fill = hue.beadFill;
        style.color = hue.beadRim;
        style.setOverride(AppTheme.Dark.name, "fill", hue.darkBeadFill);
        style.setOverride(AppTheme.Dark.name, "color", hue.darkBeadRim);
        return style;
    }

    static hueLine(name, hue, strokeWidth, dash) {
        const style = new LineStyle();
        style.name = name;
        style.color = hue.stroke;
        style.strokeWidth = strokeWidth;
        style.dash = dash;
        style.setOverride(AppTheme.Dark.name, "color", hue.darkStroke);
        return style;
    }

    /** A shape filled with the hue's gradient or with its lighter end alone, outlined in the hue or not at all */
    static hueShape(name, hue, outlined, gradient) {
        const style = new ShapeStyle();
        style.name = name;
        style.color = outlined ? hue.stroke : Color.transparent;
        style.strokeWidth = outlined ? StyleManager.ThinStrokeWidth : 1;
        style.fill = gradient ? hue.fill : hue.solidFill;
        style.setOverride(AppTheme.Dark.name, "fill", gradient ? hue.darkFill : hue.darkSolidFill);
        if (outlined) {
            style.setOverride(AppTheme.Dark.name, "color", hue.darkStroke);
        }

        return style;
    }

    /** A text style in the theme's ink */
    static themedText(name, fontSize) {
        const style = new TextStyle();
        style.name = name;
        style.fontSize = fontSize;
        style.fontFamily = "Segoe UI";
        style.bindToTheme("color", theme => theme.ink);
        return style;
    }

    getStyle(name) {
        return this.get(name);
    }

    /**
     * The style a new figure gets: points go by their kind, a slider gets its track's;
     * everything else the first style of the default name that fits (GetDefaultStyle)
     */
    assignDefaultStyle(figure) {
        const byKind = figure.isPoint === true ? this.getStyle(StyleManager.getDefaultPointStyleName(figure))
            : figure.isSlider === true ? this.getStyle(StyleManager.SliderTrackStyleName)
                : null;
        if (byKind != null && byKind.constructor.supportsFigure(figure)) {
            return byKind;
        }

        return this.getDefaultStyle(StyleManager.styledFigure(figure));
    }

    /** The figure whose styles the figure takes: itself, but for a Bezier path, whose style is its inside's */
    static styledFigure(figure) {
        return figure.isBezierPath === true && figure.interior != null ? figure.interior : figure;
    }

    /** The outlined shape for a figure drawn as an outline around an inside, else the line, or for a shape the default fill, else the first style that fits */
    getDefaultStyle(figure) {
        const supportedStyles = this.getSupportedStyles(figure);
        if (StyleManager.isOutlinedShape(figure)) {
            const outlined = supportedStyles.find(style => style.name === StyleManager.OutlinedShapeStyleName);
            if (outlined != null) {
                return outlined;
            }
        }

        return supportedStyles.find(style => style.name === StyleManager.LineStyleName || style.name === StyleManager.ShapeStyleName)
            ?? supportedStyles[0]
            ?? null;
    }

    /** An arc with an inside (a sector, a circular segment) draws its own outline around a fill; not a circle, whose default is the line, nor a polygon, whose outline is its side segments */
    static isOutlinedShape(figure) {
        return figure.isArc === true && figure.isShapeWithInterior === true;
    }

    getSupportedStyles(figure) {
        return this.list.filter(style => style.constructor.supportsFigure(figure));
    }

    static getDefaultPointStyleName(point) {
        if (point.isBezierPathHandle === true) {
            return StyleManager.HandleStyleName;
        }

        // PointOnFigure is a FreePoint, so it goes first
        if (point instanceof PointOnFigure) {
            return StyleManager.PointOnFigureStyleName;
        }

        if (point instanceof FreePoint) {
            return StyleManager.FreePointStyleName;
        }

        if (point instanceof IntersectionPoint) {
            return StyleManager.IntersectionPointStyleName;
        }

        if (point instanceof MidPoint) {
            return StyleManager.MidpointStyleName;
        }

        // draggable along its circle or line, like a point on a figure
        if (point instanceof TranslatedPoint && point.hasFreedom) {
            return StyleManager.PointOnFigureStyleName;
        }

        return StyleManager.DependentPointStyleName;
    }

    clear() {
        this.list = [];
    }

    add(style) {
        this.list.push(style);
        this.onStyleAdded(style);
    }

    /** A fresh set of the styles a new drawing starts with, names included */
    static createDefaultStyles() {
        return new StyleManager(null).list;
    }

    /**
     * What the defaults looked like in files from before the named defaults (the phone and
     * CD drawings carry numbered copies): a file style that looks like one of these is taken
     * for the default it stands for, which follows the theme.
     */
    static legacyDefaults() {
        const legacyPoint = fill => {
            const style = new PointStyle();
            style.size = 10;
            style.fill = new SolidColorBrush(fill);
            style.color = Color.fromRgb(0, 0, 0);
            style.strokeWidth = 1;
            return style;
        };
        const legacyText = fontSize => {
            const style = new TextStyle();
            style.fontSize = fontSize;
            style.color = Color.fromRgb(0, 0, 0);
            style.fontFamily = "Segoe UI";
            return style;
        };
        return [
            [legacyPoint(Color.fromRgb(0xFF, 0xFF, 0x00)), StyleManager.FreePointStyleName],
            [legacyPoint(Color.fromRgb(0x00, 0xFF, 0x00)), StyleManager.PointOnFigureStyleName],
            [legacyPoint(Color.fromRgb(0xC0, 0xC0, 0xC0)), StyleManager.DependentPointStyleName],
            [legacyText(18), StyleManager.TextStyleName],
            [legacyText(40), StyleManager.HeadingStyleName]
        ];
    }

    static legacySignatures = null;

    static getLegacySignatures() {
        if (StyleManager.legacySignatures == null) {
            StyleManager.legacySignatures = StyleManager.legacyDefaults()
                .map(([prototype, name]) => ({ type: prototype.constructor, signature: prototype.getBaseSignature(), name }));
        }

        return StyleManager.legacySignatures;
    }

    /** The name of the default the style looks like in Light, as one of the legacy defaults; null when none */
    static findLegacyDefault(style) {
        if (style.overrides.size !== 0 || style.saysThemes) {
            return null;
        }

        let signature = null;
        for (const legacy of StyleManager.getLegacySignatures()) {
            if (legacy.type !== style.constructor) {
                continue;
            }

            signature = signature ?? style.getBaseSignature();
            if (signature === legacy.signature) {
                return legacy.name;
            }
        }

        return null;
    }

    static looksLikeInLight(style, other) {
        return style.constructor === other.constructor
            && style.overrides.size === 0
            && !style.saysThemes
            && style.getBaseSignature() === other.getBaseSignature();
    }

    /**
     * A drawing's own styles (from a file, which carries only the ones its figures use)
     * completed with the default styles it lacks by name, in the order of a new drawing.
     * A file's copy of a default that looks the same in Light is dropped for the default
     * itself, which follows the theme. So is a copy under another name of what a default
     * used to look like (legacyDefaults): the figures find the default under the old name.
     */
    addWithDefaults(own) {
        const taken = new Set();
        this.aliases.clear();
        for (const style of own) {
            if (style.name === StyleManager.LegacyDependentPointStyleName) {
                style.name = StyleManager.DependentPointStyleName;
            }
        }

        for (const defaultStyle of StyleManager.createDefaultStyles()) {
            let replacement = own.find(s => s.name === defaultStyle.name) ?? null;
            if (replacement != null) {
                taken.add(replacement);
                if (StyleManager.looksLikeInLight(replacement, defaultStyle)) {
                    replacement = null;
                }
            }

            this.add(replacement ?? defaultStyle);
        }

        for (const style of own) {
            if (taken.has(style)) {
                continue;
            }

            const legacyName = StyleManager.findLegacyDefault(style);
            if (legacyName != null) {
                this.aliases.set(style.name, legacyName);
            } else {
                this.add(style);
            }
        }
    }
}
