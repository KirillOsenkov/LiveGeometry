// Port of Main/Avalonia/DynamicGeometry/Expressions/Parser/Token.cs, TokenType.cs,
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
    constructor(text, start, kind) {
        this.text = text;
        this.start = start;
        this.kind = kind;
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
