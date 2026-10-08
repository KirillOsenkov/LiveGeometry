// Port of Main/Avalonia/DynamicGeometry/Expressions/Parser/Scanner.cs

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
                case "-":
                    this.addTokenOfType(TokenType.Minus);
                    break;
                case "*":
                    this.addTokenOfType(TokenType.Multiply);
                    break;
                case "/":
                    this.addTokenOfType(TokenType.Divide);
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
                        this.current++;
                        this.scanNumericLiteralAfterPeriod("0.");
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

    /** Whether an expression can say the name: a letter (or _), then letters, digits and primes */
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
                    this.scanNumericLiteralAfterPeriod(this.text.substring(start, this.current));
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

    scanNumericLiteralAfterPeriod(beforePeriod) {
        const start = this.current;
        while (true) {
            this.current++;
            if (this.current === this.length) {
                this.addNumericLiteral(start, beforePeriod);
                return;
            }

            if (!Scanner.isDigit(this.text[this.current])) {
                this.addNumericLiteral(start, beforePeriod);
                return;
            }
        }
    }

    addNumericLiteral(start, prefix = "") {
        const token = new Token(prefix + this.text.substring(start, this.current), start - prefix.length, TokenType.NumericLiteral);
        this.result.tokens.push(token);
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

    addTokenOfType(tokenType) {
        this.result.tokens.push(new Token(this.text[this.current], this.current, tokenType));
        this.current++;
    }
}
