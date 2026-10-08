// Port of Main/Avalonia/DynamicGeometry/Styles/FigureStyle.cs and IFigureStyle.cs: a style
// has its values (how it looks under the base theme, Light) and may hold overrides for
// another theme. Whoever draws with it resolves it first (resolve).
//
// Each kind of style lists its properties (`properties`, in the order the C# reflection
// gives them: own properties first, then the base class's), with their types, which is what
// the serializer reads and the signature writes.

class FigureStyle {
    constructor() {
        this.name = "";
        this.styleManager = null;

        /** By theme name, the properties that differ under that theme, with their values */
        this.overrides = new Map();

        /** The file says how the style looks under other themes, if only that it looks no different */
        this.saysThemes = false;
        this.themeBindings = null;
    }

    /** The properties a file reads and writes: [name, type], type one of color, double, bool, enum, string, brush, fontFamily */
    static get properties() {
        return [];
    }

    /** Whether a style of this kind can be the figure's */
    static supportsFigure(figure) {
        return false;
    }

    /** The property values a fresh style of this kind has (what a file leaves out) */
    static fresh() {
        return new this();
    }

    get typeName() {
        return this.constructor.typeName;
    }

    setOverride(theme, property, value) {
        let values = this.overrides.get(theme);
        if (values == null) {
            values = new Map();
            this.overrides.set(theme, values);
        }

        values.set(property, value);
        this.onPropertyChanged(property);
    }

    removeOverride(theme, property) {
        const values = this.overrides.get(theme);
        if (values != null && values.delete(property)) {
            if (values.size === 0) {
                this.overrides.delete(theme);
            }

            this.onPropertyChanged(property);
        }
    }

    clearOverrides(theme) {
        this.overrides.delete(theme);
    }

    /** The property takes its value from every theme's colors: a default style's ink or fill */
    bindToTheme(property, value) {
        this.themeBindings = this.themeBindings ?? new Map();
        this.themeBindings.set(property, value);
        for (const theme of AppTheme.All) {
            const themeValue = value(theme);
            if (AppTheme.isBase(theme)) {
                this[property] = themeValue;
            } else {
                this.setOverride(theme.name, property, themeValue);
            }
        }
    }

    /** The style as it looks under the theme on screen */
    resolve(theme = AppTheme.current.name) {
        const values = this.overrides.get(theme);
        if (values == null || values.size === 0) {
            return this;
        }

        const result = this.clone();
        result.name = this.name;
        for (const [property, value] of values) {
            result[property] = value;
        }

        return result;
    }

    /** A copy with the same values and overrides, no name and no ties to the theme */
    clone() {
        const result = new this.constructor();
        for (const [property] of this.constructor.properties) {
            result[property] = this[property];
        }

        result.copyPrivateValues(this);
        result.name = "";
        for (const [theme, values] of this.overrides) {
            result.overrides.set(theme, new Map(values));
        }

        return result;
    }

    /** What a subclass keeps beside its listed properties (a point style's second size) */
    copyPrivateValues(source) {
    }

    onPropertyChanged(propertyName) {
        this.listeners?.forEach(listener => listener(propertyName));
    }

    addListener(listener) {
        this.listeners = this.listeners ?? [];
        this.listeners.push(listener);
    }

    removeListener(listener) {
        const index = this.listeners?.indexOf(listener) ?? -1;
        if (index >= 0) {
            this.listeners.splice(index, 1);
        }
    }

    /** The values, then each theme's overrides: two styles that look the same under every theme have the same signature */
    getSignature() {
        let result = this.getBaseSignature();
        const themes = [...this.overrides.keys()].sort();
        for (const theme of themes) {
            const values = this.overrides.get(theme);
            for (const [property] of this.constructor.properties) {
                if (values.has(property)) {
                    result += " " + theme + "." + property + "=" + StyleSerializer.writeValue(values.get(property));
                }
            }
        }

        return result;
    }

    /** The values under the base theme alone: how the style looks in Light */
    getBaseSignature() {
        const values = [];
        for (const [property] of this.constructor.properties) {
            if (property === "name") {
                continue;
            }

            const written = StyleSerializer.writeValue(this[property]);
            if (written != null) {
                values.push(written);
            }
        }

        return values.join(" ");
    }

    toString() {
        return this.name;
    }
}

/** How a property's value is written (a signature) and read (a file), by its type */
const StyleSerializer = {
    writeValue(value) {
        if (value == null) {
            return null;
        }

        if (value instanceof Color) {
            return ColorText.toArgbHex(value);
        }

        if (value instanceof SolidColorBrush) {
            return ColorText.toArgbHex(value.color);
        }

        if (value instanceof LinearGradientBrush) {
            return value.toString();
        }

        if (typeof value === "boolean") {
            return value ? "true" : "false";
        }

        if (typeof value === "number") {
            return NumberFormat.toString(value);
        }

        return String(value);
    },

    /** The value of an attribute, by the property's type */
    readValue(type, text) {
        switch (type) {
            case "color":
                return ColorText.toColor(text);
            case "double":
                return Xml.parseDouble(text);
            case "bool":
                return text.trim().toLowerCase() === "true";
            case "brush":
                return new SolidColorBrush(ColorText.toColor(text));
            default:
                return text;
        }
    },

    /** The value of a child element: a gradient brush (BrushSerializer.Read(XElement)) */
    readElementValue(type, element) {
        if (type === "brush") {
            const first = Xml.elements(element)[0];
            return first != null ? BrushSerializer.parseBrush(first) : null;
        }

        return element.textContent;
    }
};
