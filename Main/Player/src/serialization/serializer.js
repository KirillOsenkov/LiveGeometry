// Port of what the player needs of Main/Avalonia/DynamicGeometry/Serialization/Serializer.cs:
// brushes from a file (BrushSerializer) and a style element read into a style
// (ComplexTypeSerializer over the style's properties).

const BrushSerializer = {
    /** A SolidColorBrush or LinearGradientBrush element, as WPF's XAML writer wrote them; anything else is black */
    parseBrush(element) {
        switch (element.localName) {
            case "SolidColorBrush":
                return new SolidColorBrush(BrushSerializer.readColor(element.getAttribute("Color"), Color.black));
            case "LinearGradientBrush": {
                const brush = new LinearGradientBrush();
                brush.startPoint = BrushSerializer.readRelativePoint(element.getAttribute("StartPoint"), 0, 0);
                brush.endPoint = BrushSerializer.readRelativePoint(element.getAttribute("EndPoint"), 1, 1);
                for (const stop of element.getElementsByTagName("GradientStop")) {
                    brush.gradientStops.push(new GradientStop(
                        BrushSerializer.readColor(stop.getAttribute("Color"), Color.black),
                        Xml.parseDouble(stop.getAttribute("Offset") ?? "0")));
                }

                return brush;
            }

            default:
                return new SolidColorBrush(Color.black);
        }
    },

    readColor(text, fallback) {
        return text == null || text === "" ? fallback : ColorText.toColor(text);
    },

    readRelativePoint(text, defaultX, defaultY) {
        if (text != null && text !== "") {
            const parts = text.split(",");
            if (parts.length === 2) {
                const x = Number(parts[0]);
                const y = Number(parts[1]);
                if (!Number.isNaN(x) && !Number.isNaN(y)) {
                    return new Point(x, y);
                }
            }
        }

        return new Point(defaultX, defaultY);
    }
};

const StyleReader = {
    /** A style element of a file: an instance of the kind the element is named after, with the attributes and child elements it has */
    read(styleNode) {
        const type = StyleManager.StyleTypes.find(t => t.typeName === styleNode.localName);
        if (type == null) {
            return null;
        }

        const style = new type();
        for (const [property, propertyType] of type.properties) {
            const attributeName = StyleReader.attributeName(property);
            const attribute = styleNode.getAttribute(attributeName);
            if (attribute != null) {
                style[property] = StyleSerializer.readValue(propertyType, attribute);
                continue;
            }

            const subElement = Xml.element(styleNode, attributeName);
            if (subElement != null) {
                const value = StyleSerializer.readElementValue(propertyType, subElement);
                if (value != null) {
                    style[property] = value;
                }
            }
        }

        // what differs under another theme: a child element named after the theme
        for (const themeNode of Xml.elements(styleNode)) {
            const theme = themeNode.localName;
            if (AppTheme.byName(theme) == null) {
                continue;
            }

            // even an empty one: the style is as under the base theme, on purpose
            style.saysThemes = true;
            for (const attribute of themeNode.attributes) {
                const property = StyleReader.propertyOf(type, attribute.localName);
                if (property != null) {
                    style.setOverride(theme, property[0], StyleSerializer.readValue(property[1], attribute.value));
                }
            }

            for (const propertyNode of Xml.elements(themeNode)) {
                const property = StyleReader.propertyOf(type, propertyNode.localName);
                if (property != null) {
                    const value = StyleSerializer.readElementValue(property[1], propertyNode);
                    if (value != null) {
                        style.setOverride(theme, property[0], value);
                    }
                }
            }
        }

        return style;
    },

    /** The attribute a property is written as: the C# name, capitalized */
    attributeName(property) {
        return property[0].toUpperCase() + property.substring(1);
    },

    propertyOf(type, attributeName) {
        return type.properties.find(([property]) => StyleReader.attributeName(property) === attributeName) ?? null;
    }
};
