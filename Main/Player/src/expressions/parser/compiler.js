// Port of Main/Avalonia/DynamicGeometry/Expressions/Parser/Compiler.cs: text to a delegate.
// Only the OwnTree strategy (the bound tree evaluating itself) exists here.

class Compiler {
    static instance = new Compiler();

    compileFunction(drawing, functionText, isFigureAllowed = null) {
        const result = new CompileResult();
        if (functionText == null || functionText === "") {
            return result;
        }

        const ast = Compiler.parse(functionText, result);
        if (result.errors.length > 0) {
            return result;
        }

        const builder = new ExpressionTreeBuilder();
        builder.setContext(drawing, isFigureAllowed);
        try {
            const bound = builder.createFunction(ast, result);
            if (bound == null || result.errors.length > 0) {
                return result;
            }

            result.function = bound.toFunctionDelegate();
        } catch (error) {
            result.addError(error.message);
        }

        return result;
    }

    compileExpression(drawing, expressionText, isFigureAllowed) {
        const result = new CompileResult();
        if (expressionText == null || expressionText === "") {
            return result;
        }

        const ast = Compiler.parse(expressionText, result);
        if (result.errors.length > 0) {
            return result;
        }

        const builder = new ExpressionTreeBuilder();
        builder.setContext(drawing, isFigureAllowed);
        try {
            const bound = builder.createExpression(ast, result);
            if (bound == null || result.errors.length > 0) {
                return result;
            }

            result.expression = bound.toExpressionDelegate();
        } catch (error) {
            result.addError(error.message);
        }

        return result;
    }

    static parse(text, result) {
        const ast = Parser.parse(text);
        if (ast.errors.length > 0) {
            result.errors.push(...ast.errors);
        }

        return ast.root;
    }
}
