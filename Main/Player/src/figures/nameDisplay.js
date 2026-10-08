// Port of Main/Avalonia/DynamicGeometry/Figures/NameDisplay.cs: how a name is drawn, A_1 as
// A₁ and A1 as A₁ too. Only the text on the screen changes.

const NameDisplay = {
    Plain: "0123456789aeoxhklmnpst",
    Subscript: "₀₁₂₃₄₅₆₇₈₉ₐₑₒₓₕₖₗₘₙₚₛₜ",

    format(name) {
        if (name == null || name === "") {
            return name;
        }

        if (name.indexOf("_") < 0) {
            return NameDisplay.formatPointNames(name) ?? NameDisplay.formatTrailingDigits(name);
        }

        let result = "";
        let i = 0;
        while (i < name.length) {
            const c = name[i];
            if (c !== "_" || i === 0) {
                result += c;
                i++;
                continue;
            }

            const start = i + 1;
            let next;
            let part;
            if (start < name.length && name[start] === "{") {
                const close = name.indexOf("}", start);
                if (close < 0) {
                    result += c;
                    i++;
                    continue;
                }

                part = name.substring(start + 1, close);
                next = close + 1;
            } else {
                let end = start;
                while (end < name.length && NameDisplay.isLetterOrDigit(name[end])) {
                    end++;
                }

                part = name.substring(start, end);
                next = end;
            }

            const subscript = NameDisplay.toSubscript(part);
            result += subscript ?? name.substring(i, next);
            i = next;
        }

        return result;
    },

    isLetterOrDigit(c) {
        return /[\p{L}\p{Nd}]/u.test(c);
    },

    isDigit(c) {
        return c >= "0" && c <= "9";
    },

    isUpper(c) {
        return /\p{Lu}/u.test(c);
    },

    /** G1H1IJ1, a figure named after its points, as G₁H₁IJ₁; null when the name is not capital letters, each with its primes and digits */
    formatPointNames(name) {
        let result = "";
        let i = 0;
        while (i < name.length) {
            if (!NameDisplay.isUpper(name[i])) {
                return null;
            }

            result += name[i];
            i++;
            while (i < name.length && name[i] === "'") {
                result += name[i];
                i++;
            }

            const start = i;
            while (i < name.length && NameDisplay.isDigit(name[i])) {
                i++;
            }

            if (i > start) {
                result += NameDisplay.toSubscript(name.substring(start, i));
            }
        }

        return result;
    },

    /** A1 as A₁; a name that is all digits, or ends in none, stays */
    formatTrailingDigits(name) {
        let start = name.length;
        while (start > 0 && NameDisplay.isDigit(name[start - 1])) {
            start--;
        }

        if (start === 0 || start === name.length) {
            return name;
        }

        return name.substring(0, start) + NameDisplay.toSubscript(name.substring(start));
    },

    /** The text in subscript characters, or null when a character has none */
    toSubscript(text) {
        if (text.length === 0) {
            return null;
        }

        let result = "";
        for (const c of text) {
            const index = NameDisplay.Plain.indexOf(c);
            if (index < 0) {
                return null;
            }

            result += NameDisplay.Subscript[index];
        }

        return result;
    }
};
