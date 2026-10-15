// Port of Main/Avalonia/LiveGeometry/Expressions/Parser/Token.cs, TokenType.cs,
// ScanResult.cs, ParseResult.cs and CompileError.cs

const TokenType = {
    NumericLiteral: "NumericLiteral",
    Identifier: "Identifier",
    Plus: "Plus",
    Minus: "Minus",
    Multiply: "Multiply",
    Divide: "Divide",
    Power: "Power",
    OpenParen: "OpenParen",
    CloseParen: "CloseParen",
    Comma: "Comma",
    Dot: "Dot",
    Unknown: "Unknown"
};

class Token {
    /**
     * length: how much of the expression's text the token takes, more than its text for a
     * name in backticks; isQuoted: a name in backticks (`Drag me!`), the figure called
     * exactly that
     */
    constructor(text, start, kind, length = text.length, isQuoted = false) {
        this.text = text;
        this.start = start;
        this.kind = kind;
        this.length = length;
        this.isQuoted = isQuoted;
    }
}

class CompileError {
    constructor(text, position = 0) {
        this.text = text;
        this.position = position;
    }
}

class ScanResult {
    constructor() {
        this.tokens = [];
        this.errors = [];
    }
}

class ParseResult {
    constructor(errors = []) {
        this.root = null;
        this.errors = errors;
    }
}
