// Port of Main/Avalonia/DynamicGeometry/Expressions/DrawingExpression.cs: an expression owned
// by a figure (a point by coordinates, a line or circle by equation), which depends on the
// figures its expressions name and on nothing else.

class DrawingExpression {
    constructor(parent, name = "", expressionText = "") {
        this.parentFigure = parent;
        this.name = name;
        this.textValue = expressionText;
        this.isValid = false;
        this.mValue = null;
        this.dependencies = null;
    }

    get text() {
        return this.textValue;
    }

    set text(value) {
        this.textValue = value ?? "";
        this.mValue = null;
    }

    /** The compiled expression, a function of nothing; null while it doesn't compile */
    get value() {
        if (this.mValue == null) {
            this.recalculate();
        }

        return this.mValue;
    }

    recalculate() {
        const parent = this.parentFigure;
        const result = Compiler.instance.compileExpression(parent.drawing, this.text, f => !f.dependsOn(parent));
        this.isValid = result.isSuccess;
        if (!this.isValid) {
            return;
        }

        this.mValue = result.expression;
        this.dependencies = result.dependencies;

        // the figure depends on what its expressions name, all of them, listed in the order of the expressions
        const named = [];
        const expressions = parent.isExpressionOwner === true ? parent.expressions : [this];
        const figures = parent.drawing.figures;

        // a figure being read is asked for its place before what it names is in the drawing:
        // its dependencies are then what the file says until all of them compile
        if (!figures.contains(parent)
            && expressions.some(expression => expression != null
                && expression !== this
                && expression.dependencies == null
                && expression.text !== "")) {
            return;
        }

        parent.unregisterFromDependencies();
        for (const expression of expressions) {
            if (expression == null || expression.dependencies == null) {
                continue;
            }

            for (const dependency of expression.dependencies) {
                if ((expression === this || containsRecursively(figures.list, dependency)) && !named.includes(dependency)) {
                    named.push(dependency);
                }
            }
        }

        parent.dependencies = named;

        // only when the parent is already in the drawing: RegisterWithDependencies is called when it is added
        if (figures.contains(parent)) {
            parent.registerWithDependencies();
            if (!parent.drawing.isReading) {
                parent.recalculateAllDependents();
            }
        }
    }

    /** Compiles the text again */
    rebind() {
        if (this.text !== "") {
            this.recalculate();
        }
    }

    toString() {
        return this.text;
    }
}
