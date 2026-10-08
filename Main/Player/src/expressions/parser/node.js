// Port of Main/Avalonia/DynamicGeometry/Expressions/Parser/Node.cs and NodeType.cs. The class
// is SyntaxNode here: Node is the DOM's.

const NodeType = {
    Unknown: "Unknown",
    Negation: "Negation",
    Addition: "Addition",
    Subtraction: "Subtraction",
    Multiplication: "Multiplication",
    Division: "Division",
    Power: "Power",
    PropertyAccess: "PropertyAccess",
    FunctionCall: "FunctionCall",
    Constant: "Constant",
    Variable: "Variable"
};

class SyntaxNode {
    /** A node of a kind with children, or with a token */
    constructor(kind, childrenOrToken) {
        this.kind = kind;
        if (childrenOrToken instanceof Token) {
            this.token = childrenOrToken;
            this.children = [];
        } else {
            this.token = null;
            this.children = childrenOrToken ?? [];
        }
    }
}
