// Port of Main/Avalonia/LiveGeometry/Expressions/CompileResult.cs

class CompileResult {
    constructor() {
        /** A function of x (compileFunction) */
        this.function = null;

        /** A function of nothing giving the value (compileExpression) */
        this.expression = null;
        this.dependencies = [];
        this.errors = [];
    }

    get isSuccess() {
        return this.errors.length === 0 && (this.expression != null || this.function != null);
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
