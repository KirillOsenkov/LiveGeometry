// Port of Main/Avalonia/LiveGeometry/Expressions/Parser/Scanner.cs

class Scanner {
    constructor(text) {
        this.text = text;
        this.current = 0;
        this.length = 0;
        this.result = new ScanResult();
    }

    static scan(expression) {
        const scanner = new Scanner(expression);
        scanner.scanAll();
        return scanner.result;
    }

    scanAll() {
        const text = this.text;
        this.current = 0;
        this.length = text.length;
        while (this.current < this.length) {
            const currentChar = text[this.current];
            switch (currentChar) {
                case "+":
                    this.addTokenOfType(TokenType.Plus);
                    break;
                // (the signs a caption writes, also when one is pasted in)
                case "-":
                case "−":
                // (what a page or a word processor has for a minus: the hyphen, the
                // non-breaking hyphen, the figure dash, the en dash)
                case "\u{2010}":
                case "\u{2011}":
                case "\u{2012}":
                case "\u{2013}":
                    this.addTokenOfType(TokenType.Minus);
                    break;
                case "*":
                case "×":
                case "·":
                // (the dot operator, the bullet operator, the asterisk operator, the bullet)
                case "\u{22C5}":
                case "\u{2219}":
                case "\u{2217}":
                case "\u{2022}":
                    this.addTokenOfType(TokenType.Multiply);
                    break;
                case "/":
                case "÷":
                // (the division slash, the fraction slash)
                case "\u{2215}":
                case "\u{2044}":
                    this.addTokenOfType(TokenType.Divide);
                    break;
                // √x, ∛x, as in class
                case "√":
                case "∛":
                    this.addTokenOfType(TokenType.Root);
                    break;
                case "°":
                    this.addTokenOfType(TokenType.Degree);
                    break;
                case "|":
                    this.addTokenOfType(TokenType.Pipe);
                    break;
                // ∠ABC, the angle at B, as in class
                case "∠":
                    this.addTokenOfType(TokenType.Angle);
                    break;
                case "<":
                    this.addTokenOfType(this.nextChar() === "=" ? TokenType.LessOrEqual : TokenType.Less, this.nextChar() === "=" ? 2 : 1);
                    break;
                case ">":
                    this.addTokenOfType(this.nextChar() === "=" ? TokenType.GreaterOrEqual : TokenType.Greater, this.nextChar() === "=" ? 2 : 1);
                    break;
                case "≤":
                case "⩽":
                    this.addTokenOfType(TokenType.LessOrEqual);
                    break;
                case "≥":
                case "⩾":
                    this.addTokenOfType(TokenType.GreaterOrEqual);
                    break;
                // = as in class, == as programmers and GeoGebra write it
                case "=":
                    this.addTokenOfType(TokenType.Equal, this.nextChar() === "=" ? 2 : 1);
                    break;
                case "≠":
                    this.addTokenOfType(TokenType.NotEqual);
                    break;
                case "!":
                    if (this.nextChar() !== "=") {
                        this.reportError("Invalid character");
                        return;
                    }

                    this.addTokenOfType(TokenType.NotEqual, 2);
                    break;
                case "^":
                    this.addTokenOfType(TokenType.Power);
                    break;
                case "(":
                    this.addTokenOfType(TokenType.OpenParen);
                    break;
                case ")":
                    this.addTokenOfType(TokenType.CloseParen);
                    break;
                case ".":
                    if (Scanner.isDigit(this.nextChar())) {
                        const literalStart = this.current;
                        this.current++;
                        this.scanNumericLiteralAfterPeriod(literalStart, "0.");
                    } else {
                        this.addTokenOfType(TokenType.Dot);
                    }

                    break;
                case ",":
                    this.addTokenOfType(TokenType.Comma);
                    break;
                case " ":
                    this.current++;
                    break;
                case "`":
                    if (!this.scanQuotedName()) {
                        return;
                    }

                    break;
                default:
                    if (Scanner.isDigit(currentChar)) {
                        this.scanNumericLiteral();
                    } else if (Scanner.isLetter(currentChar)) {
                        this.scanIdentifier();
                    } else if (Scanner.isSpace(currentChar)) {
                        this.current++;
                    } else if (Scanner.fractions.has(currentChar)) {
                        this.result.tokens.push(new Token(Scanner.fractions.get(currentChar), this.current, TokenType.Fraction, 1));
                        this.current++;
                    } else if (!this.scanSuperscript()) {
                        this.reportError("Invalid character");
                        return;
                    }

                    break;
            }
        }
    }

    /** What a pasted text has between its words besides the space: a tab, a line break, the non-breaking and the thin spaces of a page, the zero-width ones nobody sees */
    static isSpace(ch) {
        return ch === " "
            || (ch >= "\t" && ch <= "\r")
            || ch === "\u{0085}"
            || ch === "\u{00A0}"
            || ch === "\u{1680}"
            || (ch >= "\u{2000}" && ch <= "\u{200B}")
            || ch === "\u{2028}"
            || ch === "\u{2029}"
            || ch === "\u{202F}"
            || ch === "\u{205F}"
            || ch === "\u{2060}"
            || ch === "\u{3000}"
            || ch === "\u{FEFF}";
    }

    /** The digit a raised one writes (² is 2), or -1 */
    static superscriptDigit(ch) {
        return "⁰¹²³⁴⁵⁶⁷⁸⁹".indexOf(ch);
    }

    /**
     * A raised whole number, x² or x⁻¹: one token that reads as the digits it writes (2, -1)
     * and takes the raised ones' place in the text. False when there is none here.
     */
    scanSuperscript() {
        const text = this.text;
        const start = this.current;
        let digits = "";
        let index = this.current;
        if (text[index] === "⁻") {
            digits = "-";
            index++;
        }

        while (index < this.length && Scanner.superscriptDigit(text[index]) >= 0) {
            digits += Scanner.superscriptDigit(text[index]);
            index++;
        }

        if (digits === "" || digits === "-") {
            return false;
        }

        this.current = index;
        this.result.tokens.push(new Token(digits, start, TokenType.Superscript, this.current - start));
        return true;
    }

    reportError(error) {
        this.result.errors.push(new CompileError(error, this.current));
    }

    static isDigit(ch) {
        return ch >= "0" && ch <= "9";
    }

    static isLetter(ch) {
        return ch === "_" || /\p{L}/u.test(ch);
    }

    /** A prime after a name: A', A'' - the image of A, as in class */
    static isPrime(ch) {
        return ch === "'" || ch === "′" || ch === "″";
    }

    /** Whether an expression can say the name as it is: a letter (or _), then letters, digits and primes; any other it says in backticks */
    static isName(name) {
        if (name == null || name === "" || !Scanner.isLetter(name[0])) {
            return false;
        }

        for (let i = 1; i < name.length; i++) {
            if (!Scanner.isLetter(name[i]) && !Scanner.isDigit(name[i]) && !Scanner.isPrime(name[i])) {
                return false;
            }
        }

        return true;
    }

    /**
     * A name in backticks, from the opening one: what is between them, a doubled backtick
     * standing for one. False, with the error said, when it is empty or not closed.
     */
    scanQuotedName() {
        const text = this.text;
        const start = this.current;
        let name = "";
        this.current++;
        while (this.current < this.length) {
            if (text[this.current] === "`") {
                if (this.nextChar() !== "`") {
                    this.current++;
                    if (name === "") {
                        this.current = start;
                        this.reportError("A name in backticks can't be empty");
                        return false;
                    }

                    this.result.tokens.push(new Token(name, start, TokenType.Identifier, this.current - start, true));
                    return true;
                }

                this.current++;
            }

            name += text[this.current];
            this.current++;
        }

        this.current = start;
        this.reportError("A name in backticks needs a closing backtick");
        return false;
    }

    scanNumericLiteral() {
        const start = this.current;
        while (true) {
            this.current++;
            if (this.current === this.length) {
                this.addNumericLiteral(start);
                return;
            }

            const currentChar = this.text[this.current];
            if (currentChar === ".") {
                if (Scanner.isDigit(this.nextChar())) {
                    this.current++;
                    this.scanNumericLiteralAfterPeriod(start, this.text.substring(start, this.current));
                    return;
                }

                this.reportError("Decimal period must be followed by fractional part (a digit)");
                this.current = this.length;
                return;
            }

            if (!Scanner.isDigit(currentChar)) {
                this.addNumericLiteral(start);
                return;
            }
        }
    }

    /** literalStart: where the number begins in the text; beforePeriod: the number up to its digits after the period, as read ("12.", "0." for .5) */
    scanNumericLiteralAfterPeriod(literalStart, beforePeriod) {
        const start = this.current;
        while (true) {
            this.current++;
            if (this.current === this.length || !Scanner.isDigit(this.text[this.current])) {
                this.addNumericLiteral(literalStart, beforePeriod + this.text.substring(start, this.current));
                return;
            }
        }
    }

    /** The number from start up to here, read as literal: its token takes its place in the text, also where it reads otherwise (.5 is 0.5) */
    addNumericLiteral(start, literal = this.text.substring(start, this.current)) {
        this.result.tokens.push(new Token(literal, start, TokenType.NumericLiteral, this.current - start));
    }

    nextChar() {
        const index = this.current + 1;
        return index < this.length ? this.text[index] : "\0";
    }

    scanIdentifier() {
        const start = this.current;
        while (true) {
            this.current++;
            if (this.current === this.length) {
                this.addIdentifier(start);
                return;
            }

            const currentChar = this.text[this.current];
            if (!Scanner.isDigit(currentChar) && !Scanner.isLetter(currentChar) && !Scanner.isPrime(currentChar)) {
                this.addIdentifier(start);
                return;
            }
        }
    }

    addIdentifier(start) {
        this.result.tokens.push(new Token(this.text.substring(start, this.current), start, TokenType.Identifier));
    }

    addTokenOfType(tokenType, length = 1) {
        this.result.tokens.push(new Token(this.text.substring(this.current, this.current + length), this.current, tokenType));
        this.current += length;
    }
}

/** The division a fraction's character writes (½ is 1/2) */
Scanner.fractions = new Map([
    ["½", "1/2"], ["⅓", "1/3"], ["⅔", "2/3"], ["¼", "1/4"], ["¾", "3/4"], ["⅕", "1/5"], ["⅖", "2/5"], ["⅗", "3/5"], ["⅘", "4/5"],
    ["⅙", "1/6"], ["⅚", "5/6"], ["⅐", "1/7"], ["⅛", "1/8"], ["⅜", "3/8"], ["⅝", "5/8"], ["⅞", "7/8"], ["⅑", "1/9"], ["⅒", "1/10"]
]);
