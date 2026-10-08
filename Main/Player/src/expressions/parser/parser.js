// Port of Main/Avalonia/DynamicGeometry/Expressions/Parser/Parser.cs: a precedence-climbing
// parser. A minus in front takes everything up to the next + - * /, powers included; ^ is
// right-associative; there is no implicit multiplication.

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

            if (!Parser.isOperatorToken(this.currentToken.kind)) {
                break;
            }

            const operationType = Parser.getExpressionType(this.currentToken.kind);
            const newPrecedence = Parser.getPrecedence(operationType);
            if (newPrecedence < precedence) {
                break;
            }

            if (newPrecedence === precedence && !Parser.isRightAssociative(operationType)) {
                break;
            }

            this.advanceToken();
            const rightOperand = this.parseExpression(newPrecedence);
            leftOperand = new SyntaxNode(operationType, [leftOperand, rightOperand]);
        }

        return leftOperand;
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
        switch (tokenType) {
            case TokenType.Plus:
            case TokenType.Minus:
            case TokenType.Multiply:
            case TokenType.Divide:
            case TokenType.Power:
                return true;
            default:
                return false;
        }
    }

    parseTerm(precedence) {
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

                if (next.kind === TokenType.OpenParen) {
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
            default:
                return NodeType.Unknown;
        }
    }

    static getPrecedence(nodeType) {
        switch (nodeType) {
            case NodeType.Constant:
                return 0;
            case NodeType.Addition:
            case NodeType.Subtraction:
                return 1;
            case NodeType.Multiplication:
            case NodeType.Division:
                return 2;
            case NodeType.Power:
                return 3;
            case NodeType.Negation:
                return 4;
            case NodeType.PropertyAccess:
                return 5;
            case NodeType.FunctionCall:
                return 6;
            default:
                return 0;
        }
    }
}
