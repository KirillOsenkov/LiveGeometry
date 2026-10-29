// Port of Main/Avalonia/LiveGeometry/Expressions/Parser/Parser.cs: a precedence-climbing
// parser. A minus in front takes everything up to the next + - * /, powers included; ^ is
// right-associative; a number right before a name or a parenthesis multiplies it (2x);
// comparisons come last, ° follows a term and |x| is between bars.

class Parser {
    constructor(tokens) {
        this.result = new ParseResult();
        this.tokens = tokens;
        this.current = 0;
        this.length = tokens.length;
    }

    static parse(expression) {
        const scanResult = Scanner.scan(expression);
        if (scanResult.errors.length > 0) {
            return new ParseResult(scanResult.errors);
        }

        const parser = new Parser(scanResult.tokens);
        parser.parseAll();
        return parser.result;
    }

    get currentToken() {
        return this.eof ? null : this.tokens[this.current];
    }

    get currentTokenKind() {
        return this.eof ? TokenType.Unknown : this.currentToken.kind;
    }

    get next() {
        return this.current + 1 >= this.length ? null : this.tokens[this.current + 1];
    }

    parseAll() {
        const expression = this.parseExpression(0);
        if (this.current < this.length) {
            this.reportError("Unexpected expression ending: " + this.tokens[this.current].text);
        } else {
            this.result.root = expression;
        }
    }

    parseExpression(precedence = 0) {
        let leftOperand = null;
        if (this.current === this.length) {
            this.reportError("Expression expected");
            return null;
        }

        if (this.currentToken.kind === TokenType.Minus) {
            this.advanceToken();
            if (this.current === this.length) {
                this.reportError("Expected an expression after a negation sign -");
                return null;
            }

            // the minus takes what follows up to the next + - * /, powers included: -x^2 is -(x^2)
            const unaryOperand = this.parseExpression(Parser.getPrecedence(NodeType.Multiplication));
            leftOperand = new SyntaxNode(NodeType.Negation, [unaryOperand]);
        } else {
            leftOperand = this.parseTerm(precedence);
        }

        while (true) {
            if (this.current === this.length) {
                break;
            }

            // a number right before a name or a parenthesis multiplies it: 2x, 3sin(x), 2(x + 1)
            const implicitMultiplication = this.isImplicitMultiplication();
            if (!implicitMultiplication && !Parser.isOperatorToken(this.currentToken.kind)) {
                break;
            }

            const operationType = implicitMultiplication ? NodeType.Multiplication : Parser.getExpressionType(this.currentToken.kind);
            const newPrecedence = Parser.getPrecedence(operationType);
            if (newPrecedence < precedence) {
                break;
            }

            if (newPrecedence === precedence && !Parser.isRightAssociative(operationType)) {
                break;
            }

            if (!implicitMultiplication) {
                this.advanceToken();
            }

            const rightOperand = this.parseExpression(newPrecedence);
            leftOperand = new SyntaxNode(operationType, [leftOperand, rightOperand]);
        }

        return leftOperand;
    }

    /** A number just read, and a name or a parenthesis right after it */
    isImplicitMultiplication() {
        return this.current > 0
            && this.tokens[this.current - 1].kind === TokenType.NumericLiteral
            && (this.currentToken.kind === TokenType.Identifier || this.currentToken.kind === TokenType.OpenParen || this.currentToken.kind === TokenType.Angle);
    }

    get eof() {
        return this.current >= this.length;
    }

    reportError(error) {
        let position = this.length - 1;
        if (!this.eof) {
            position = this.currentToken.start;
        }

        this.result.errors.push(new CompileError(error, position));
    }

    advanceToken() {
        this.current++;
    }

    static isOperatorToken(tokenType) {
        switch (Parser.getExpressionType(tokenType)) {
            case NodeType.Addition:
            case NodeType.Subtraction:
            case NodeType.Multiplication:
            case NodeType.Division:
            case NodeType.Power:
                return true;
            default:
                return Parser.isComparison(Parser.getExpressionType(tokenType));
        }
    }

    /** A term and the degree signs after it: 30°, angle(A, B, C)°, (a + b)° */
    parseTerm(precedence) {
        let term = this.parsePrimary(precedence);
        while (term != null && this.currentTokenKind === TokenType.Degree) {
            term = new SyntaxNode(NodeType.Degrees, [term]);
            this.advanceToken();
        }

        return term;
    }

    parsePrimary(precedence) {
        let result = null;
        switch (this.currentToken.kind) {
            case TokenType.NumericLiteral:
                result = this.constant();
                this.advanceToken();
                break;
            case TokenType.Identifier: {
                const next = this.next;
                if (next == null) {
                    result = this.variable();
                    this.advanceToken();
                    break;
                }

                if (next.kind === TokenType.Dot) {
                    return this.parsePropertyAccess();
                }

                // (a name in backticks is a figure's, never a function's)
                if (next.kind === TokenType.OpenParen && !this.currentToken.isQuoted) {
                    return this.parseFunctionCall();
                }

                result = this.variable();
                this.advanceToken();
                break;
            }

            case TokenType.OpenParen:
                this.advanceToken();
                result = this.parseExpression(0);
                this.expectToken(TokenType.CloseParen);
                break;
            case TokenType.Pipe:
                // |x|: what is between the bars, which end it as a ) ends a parenthesis
                this.advanceToken();
                result = new SyntaxNode(NodeType.AbsoluteValue, [this.parseExpression(0)]);
                this.expectToken(TokenType.Pipe);
                break;
            case TokenType.Angle:
                return this.parseAngle();
            default:
                this.reportError("Expected a number, a variable, a function or a parenthesized expression");
                break;
        }

        return result;
    }

    parseFunctionCall() {
        const result = new SyntaxNode(NodeType.FunctionCall, this.currentToken);
        this.advanceToken();
        this.expectToken(TokenType.OpenParen);
        result.children.push(this.parseExpression());
        while (this.currentTokenKind === TokenType.Comma) {
            this.advanceToken();
            result.children.push(this.parseExpression());
        }

        this.expectToken(TokenType.CloseParen);
        return result;
    }

    /** ∠ABC: a call of angle (the sign is one of its names) with the three points' names in one, as angle(ABC); ∠(A, B, C) is angle(A, B, C) */
    parseAngle() {
        if (this.next?.kind === TokenType.OpenParen) {
            return this.parseFunctionCall();
        }

        const result = new SyntaxNode(NodeType.FunctionCall, this.currentToken);
        this.advanceToken();
        if (this.currentTokenKind !== TokenType.Identifier) {
            this.reportError("Expected the names of three points after ∠: ∠ABC");
            return null;
        }

        result.children.push(this.variable());
        this.advanceToken();
        return result;
    }

    expectToken(tokenType) {
        if (this.current === this.length || this.currentToken.kind !== tokenType) {
            this.reportError("Expected " + tokenType);
            return;
        }

        this.advanceToken();
    }

    parsePropertyAccess() {
        const left = this.variable();
        this.expectToken(TokenType.Identifier);
        this.expectToken(TokenType.Dot);
        const right = this.variable();
        this.expectToken(TokenType.Identifier);
        return new SyntaxNode(NodeType.PropertyAccess, [left, right]);
    }

    variable() {
        if (this.eof) {
            this.reportError("Variable expected");
            return null;
        }

        return new SyntaxNode(NodeType.Variable, this.currentToken);
    }

    constant() {
        if (this.eof) {
            this.reportError("Constant expected");
            return null;
        }

        return new SyntaxNode(NodeType.Constant, this.currentToken);
    }

    static isRightAssociative(operation) {
        return operation === NodeType.Power;
    }

    static getExpressionType(tokenType) {
        switch (tokenType) {
            case TokenType.NumericLiteral:
                return NodeType.Constant;
            case TokenType.Identifier:
                return NodeType.Variable;
            case TokenType.Plus:
                return NodeType.Addition;
            case TokenType.Minus:
                return NodeType.Subtraction;
            case TokenType.Multiply:
                return NodeType.Multiplication;
            case TokenType.Divide:
                return NodeType.Division;
            case TokenType.Power:
                return NodeType.Power;
            case TokenType.Less:
                return NodeType.Less;
            case TokenType.LessOrEqual:
                return NodeType.LessOrEqual;
            case TokenType.Greater:
                return NodeType.Greater;
            case TokenType.GreaterOrEqual:
                return NodeType.GreaterOrEqual;
            case TokenType.Equal:
                return NodeType.Equal;
            case TokenType.NotEqual:
                return NodeType.NotEqual;
            default:
                return NodeType.Unknown;
        }
    }

    /** A comparison: what if(...) asks, 1 when it holds and 0 when it doesn't */
    static isComparison(nodeType) {
        switch (nodeType) {
            case NodeType.Less:
            case NodeType.LessOrEqual:
            case NodeType.Greater:
            case NodeType.GreaterOrEqual:
            case NodeType.Equal:
            case NodeType.NotEqual:
                return true;
            default:
                return false;
        }
    }

    static getPrecedence(nodeType) {
        if (Parser.isComparison(nodeType)) {
            return 1;
        }

        switch (nodeType) {
            case NodeType.Constant:
                return 0;
            case NodeType.Addition:
            case NodeType.Subtraction:
                return 2;
            case NodeType.Multiplication:
            case NodeType.Division:
                return 3;
            case NodeType.Power:
                return 4;
            case NodeType.Negation:
                return 5;
            case NodeType.PropertyAccess:
                return 6;
            case NodeType.FunctionCall:
                return 7;
            default:
                return 0;
        }
    }
}
