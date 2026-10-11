// Reading attributes of a drawing's XML, as the ReadDouble/ReadBool/ReadString extensions
// of Utilities.cs do: invariant culture, a missing attribute reads as 0, false (or the given
// default) and null.

const Xml = {
    parse(text) {
        // an XML declaration must come first: the text inside an embed's script tag starts
        // with the line break after the tag, and a file may start with a byte order mark
        const document = new DOMParser().parseFromString(text.replace(/^﻿/, "").trimStart(), "application/xml");
        const error = document.querySelector("parsererror");
        if (error != null) {
            throw new Error("The file is not well-formed XML: " + error.textContent);
        }

        return document.documentElement;
    },

    /** double.TryParse with NumberStyles.Float and the invariant culture; 0 when it fails */
    readDouble(element, attributeName) {
        const text = element.getAttribute(attributeName);
        if (text == null) {
            return 0;
        }

        return Xml.parseDouble(text);
    },

    parseDouble(text) {
        const trimmed = text.trim();
        if (/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(trimmed)) {
            return Number(trimmed);
        }

        if (trimmed === "NaN") {
            return NaN;
        }

        if (trimmed === "Infinity" || trimmed === "∞") {
            return Infinity;
        }

        if (trimmed === "-Infinity" || trimmed === "-∞") {
            return -Infinity;
        }

        return 0;
    },

    /** A whole number; the default when the attribute is missing or is no number (ReadInt) */
    readInt(element, attributeName, defaultValue) {
        const text = element.getAttribute(attributeName);
        if (text == null || !/^\s*[+-]?\d+\s*$/.test(text)) {
            return defaultValue;
        }

        return parseInt(text, 10);
    },

    /** bool.Parse of the attribute, the default when it is missing */
    readBool(element, attributeName, defaultValue) {
        const text = element.getAttribute(attributeName);
        if (text == null) {
            return defaultValue;
        }

        return text.trim().toLowerCase() === "true";
    },

    readString(element, attributeName) {
        return element.getAttribute(attributeName);
    },

    /** The child elements, as Elements() gives them (text nodes left out) */
    elements(element, localName = null) {
        const result = [];
        for (const child of element.children) {
            if (localName == null || child.localName === localName) {
                result.push(child);
            }
        }

        return result;
    },

    /** Element(name): the first child element of that name, or null */
    element(element, localName) {
        for (const child of element.children) {
            if (child.localName === localName) {
                return child;
            }
        }

        return null;
    }
};
