// Port of Main/Avalonia/LiveGeometry/Expressions/CompileResult.cs and of the ExpressionClass
// of ExpressionClassification.cs: what a piece of an expression's text is, for its color in a
// number's caption that says its formula (the app colors its expression boxes the same way).

const ExpressionClass = Object.freeze({
    /** A figure by its name: A in A.X, the points of distance(A, B) and of AB, a slider, a number */
    Figure: "Figure",

    /** What is read off a figure: the X of A.X, the Length of AB.Length */
    Property: "Property",

    /** A function: sin, sqrt, distance, angle, log_2 */
    Function: "Function",

    /** A number as typed: 3, 0.5 */
    Number: "Number",

    /** What the language gives a meaning of its own: pi, π, e, the x of a function, the degree sign */
    Reserved: "Reserved"
});

class CompileResult {
    constructor() {
        /** A function of x (compileFunction) */
        this.function = null;

        /** A function of nothing giving the value (compileExpression) */
        this.expression = null;
        this.dependencies = [];
        this.errors = [];

        /**
         * What the pieces of the text were read as, in the order of the text: each name as
         * the binder resolved it, each number as typed, { start, length, kind }. Operators,
         * parentheses and commas are left out. Whole only when the text compiled.
         */
        this.classification = [];
    }

    get isSuccess() {
        return this.errors.length === 0 && (this.expression != null || this.function != null);
    }

    classify(token, kind) {
        this.classification.push({ start: token.start, length: token.length, kind });
    }

    addError(error) {
        this.errors.push(new CompileError(error));
    }

    addBindError(figureName) {
        this.addError("Could not find figure with name '" + figureName + "'");
    }

    addPropertyNotFoundError(figure, propertyName) {
        this.addError("Could not find property '" + propertyName + "' on figure '" + figure.name + "'");
    }

    addMethodNotFoundError(functionName) {
        this.addError("There is no function '" + functionName + "'");
    }

    /** What to show under the box the text came from: null when it compiled */
    getErrorText(whenEmpty = "Type a number or an expression.") {
        if (this.isSuccess) {
            return null;
        }

        return this.errors.length === 0 ? whenEmpty : this.toString();
    }

    toString() {
        return this.errors.map(error => error.text).join("\n");
    }

    addDependencyCycleError(figureName) {
        this.addError("Using figure '" + figureName + "' will create a cycle. Circular dependencies are not allowed.");
    }

    addUnknownIdentifierError(text) {
        this.addError("Unknown identifier: '" + text + "'");
    }

    addFigureIsNotAPointError(name) {
        this.addError("Figure '" + name + "' is not a point.");
    }

    addIncorrectNumberOfArgumentsError(functionName, actualNumberOfArguments) {
        this.addError("'" + functionName + "' has " + actualNumberOfArguments + (actualNumberOfArguments === 1 ? " argument" : " arguments") + " here, which is not how it is written");
    }
}
