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
                    this.addTokenOfType(TokenType.Minus);
                    break;
                case "*":
                case "×":
                case "·":
                    this.addTokenOfType(TokenType.Multiply);
                    break;
                case "/":
                case "÷":
                    this.addTokenOfType(TokenType.Divide);
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
                    this.addTokenOfType(TokenType.LessOrEqual);
                    break;
                case "≥":
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
                    } else {
                        this.reportError("Invalid character");
                        return;
                    }

                    break;
            }
        }
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
