// Port of Main/Avalonia/LiveGeometry/Figures/Controls/LineBreaks.cs: where the lines of a
// wrapped label may break. A space next to an operator (a word that is nothing but
// operators: = + − × ≤ ...) is shown as a no-break space, so a short expression stays on one
// line ("AP = 8 − r" broke as "AP = 8 −" and "r."), and a long chain (longer than LongChain
// characters) may still break before its second and later relations. For what is shown
// only: the text keeps the spaces it was typed with.

class LineBreaks {
    static NoBreakSpace = String.fromCharCode(0xA0);

    /** The characters of an operator, a word of which a space next to it doesn't break */
    static Operators = "=≠≈<>≤≥+−-±×·÷*/";

    /** The characters of a relation, before which a long chain may break */
    static Relations = "=≠≈<>≤≥";

    /** How many characters a chain of words joined by operators may have and still never break */
    static LongChain = 24;

    static keepExpressionsTogether(text) {
        if (text == null || text === "" || !text.includes(" ")) {
            return text;
        }

        const words = LineBreaks.findWords(text);
        if (words.length < 2) {
            return text;
        }

        let characters = null;
        let chainStart = 0;
        while (chainStart < words.length) {
            let chainEnd = chainStart;
            while (chainEnd + 1 < words.length && LineBreaks.joins(text, words[chainEnd], words[chainEnd + 1])) {
                chainEnd++;
            }

            const isLong = words[chainEnd].end - words[chainStart].start > LineBreaks.LongChain;
            let hadRelation = LineBreaks.isRelation(text, words[chainStart]);
            for (let i = chainStart + 1; i <= chainEnd; i++) {
                const isRelation = LineBreaks.isRelation(text, words[i]);
                if (!(isLong && isRelation && hadRelation)) {
                    characters ??= text.split("");
                    for (let j = words[i - 1].end; j < words[i].start; j++) {
                        characters[j] = LineBreaks.NoBreakSpace;
                    }
                }

                hadRelation ||= isRelation;
            }

            chainStart = chainEnd + 1;
        }

        return characters != null ? characters.join("") : text;
    }

    /** The runs of characters between spaces and line breaks, each { start, end } (in UTF-16 units, as the C#) */
    static findWords(text) {
        const words = [];
        let start = -1;
        for (let i = 0; i <= text.length; i++) {
            const isSeparator = i === text.length || /\s/.test(text[i]) && text[i] !== LineBreaks.NoBreakSpace;
            if (!isSeparator && start < 0) {
                start = i;
            } else if (isSeparator && start >= 0) {
                words.push({ start, end: i });
                start = -1;
            }
        }

        return words;
    }

    /** Whether the spaces between two words don't break: only spaces between them, and one of them an operator */
    static joins(text, left, right) {
        for (let i = left.end; i < right.start; i++) {
            if (text[i] !== " ") {
                return false;
            }
        }

        return LineBreaks.isMadeOf(text, left, LineBreaks.Operators) || LineBreaks.isMadeOf(text, right, LineBreaks.Operators);
    }

    static isRelation(text, word) {
        if (!LineBreaks.isMadeOf(text, word, LineBreaks.Operators)) {
            return false;
        }

        for (let i = word.start; i < word.end; i++) {
            if (LineBreaks.Relations.includes(text[i])) {
                return true;
            }
        }

        return false;
    }

    static isMadeOf(text, word, characters) {
        for (let i = word.start; i < word.end; i++) {
            if (!characters.includes(text[i])) {
                return false;
            }
        }

        return true;
    }
}
