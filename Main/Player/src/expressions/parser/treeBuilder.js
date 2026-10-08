// Port of Main/Avalonia/DynamicGeometry/Expressions/Parser/TreeBuilder.cs: binds the parser's
// tree against a drawing into a BoundExpression. An error is said in the CompileResult and
// the tree is null; nothing a user can type may throw.

class ExpressionTreeBuilder {
    constructor() {
        this.binder = new Binder();
        this.status = null;
    }

    createFunction(root, status) {
        this.status = status;
        this.binder.registerParameter("x");
        return this.createExpressionCore(root);
    }

    createExpression(root, status) {
        this.status = status;
        return this.createExpressionCore(root);
    }

    createExpressionCore(root) {
        if (root == null) {
            return null;
        }

        switch (root.kind) {
            case NodeType.Negation:
                return this.createUnaryExpression(root);
            case NodeType.Addition:
            case NodeType.Subtraction:
            case NodeType.Multiplication:
            case NodeType.Division:
            case NodeType.Power:
                return this.createBinaryExpression(root);
            case NodeType.Variable:
                return this.createIdentifierExpression(root);
            case NodeType.Constant:
                // the language writes decimals with a point, whatever the user's culture
                return new BoundConstant(Number(root.token.text));
            case NodeType.FunctionCall:
                return this.createCallExpression(root);
            case NodeType.PropertyAccess:
                return this.createPropertyAccessExpression(root);
            default:
                return null;
        }
    }

    createUnaryExpression(root) {
        const operand = this.createExpressionCore(root.children[0]);
        if (operand == null) {
            return null;
        }

        return new BoundNegation(operand);
    }

    createIdentifierExpression(root) {
        const text = root.token.text;

        // a number called exactly this comes before the reading as two points
        const exact = this.binder.resolveExactNumber(text);
        if (exact != null) {
            return this.createNumberExpression(exact, text);
        }

        // pi is π whatever the drawing has
        if (text === "pi" || text === "e") {
            return this.binder.resolve(text);
        }

        const resolveTwoPoints = this.resolveTwoPoints(text);
        if (resolveTwoPoints != null) {
            return resolveTwoPoints;
        }

        const parameter = this.binder.resolve(text);
        if (parameter != null) {
            return parameter;
        }

        // a Number in the drawing, by its name: the expression then depends on it
        const number = this.binder.resolveFigure(text);
        if (number != null && number.isNumber === true) {
            return this.createNumberExpression(number, text);
        }

        this.status.addUnknownIdentifierError(text);
        return null;
    }

    createNumberExpression(number, text) {
        if (!this.binder.isFigureAllowed(number)) {
            this.status.addDependencyCycleError(text);
            return null;
        }

        this.status.dependencies.push(number);
        return new BoundNumber(number);
    }

    resolveTwoPoints(twoPoints) {
        const drawing = this.binder.drawing;
        if (drawing == null) {
            return null;
        }

        const lower = twoPoints.toLowerCase();
        let longestPrefix = "";
        let longestSuffix = "";
        for (const name of this.binder.pointNames) {
            if (name == null || name === "") {
                continue;
            }

            const lowerName = name.toLowerCase();
            if (lower.startsWith(lowerName) && name.length > longestPrefix.length) {
                longestPrefix = name;
            }

            if (lower.endsWith(lowerName) && name.length > longestSuffix.length) {
                longestSuffix = name;
            }
        }

        if (longestPrefix.length + longestSuffix.length === twoPoints.length) {
            const point1 = drawing.figures.byName(longestPrefix);
            const point2 = drawing.figures.byName(longestSuffix);
            if (!(point1 instanceof PointBase)) {
                this.status.addFigureIsNotAPointError(longestPrefix);
                return null;
            }

            if (!(point2 instanceof PointBase)) {
                this.status.addFigureIsNotAPointError(longestSuffix);
                return null;
            }

            if (!this.binder.isFigureAllowed(point1)) {
                this.status.addDependencyCycleError(longestPrefix);
                return null;
            }

            if (!this.binder.isFigureAllowed(point2)) {
                this.status.addDependencyCycleError(longestSuffix);
                return null;
            }

            this.status.dependencies.push(point1);
            this.status.dependencies.push(point2);
            return new BoundDistance(point1, point2);
        }

        return null;
    }

    createPropertyAccessExpression(root) {
        const figureName = root.children[0].token.text;
        const propertyName = root.children[1].token.text;
        const figure = this.binder.resolveFigure(figureName);
        if (figure == null) {
            this.status.addUnknownIdentifierError(figureName);
            return null;
        }

        if (!this.binder.isFigureAllowed(figure)) {
            this.status.addDependencyCycleError(figureName);
            return null;
        }

        const property = ExpressionReflection.findProperty(figure, propertyName);
        if (property == null) {
            this.status.addPropertyNotFoundError(figure, propertyName);
            return null;
        }

        if (!this.asNumber(figure[property], figureName + "." + propertyName)) {
            return null;
        }

        this.status.dependencies.push(figure);
        return new BoundProperty(figure, property);
    }

    createCallExpression(root) {
        const functionName = root.token.text;
        const args = root.children;
        const method = this.binder.resolveMethod(functionName, args.length);
        if (method == null) {
            this.status.addMethodNotFoundError(functionName);
            return null;
        }

        // a function of numbers: each argument is an expression; anything else takes points
        if (Binder.takesNumbers(method, args.length)) {
            const values = [];
            for (const node of args) {
                const value = node == null ? null : this.createExpressionCore(node);
                if (value == null) {
                    return null;
                }

                values.push(value);
            }

            return new BoundCall(method, values);
        }

        return this.createPointFunctionCallExpression(method, args);
    }

    /** What the language calculates with is a double; anything else is an error, said here */
    asNumber(value, what) {
        if (ExpressionReflection.isNumberValue(value)) {
            return true;
        }

        this.status.addError("'" + what + "' is not a number");
        return false;
    }

    createPointFunctionCallExpression(method, args) {
        if (!method.takesArray && method.parameterCount !== args.length) {
            this.status.addIncorrectNumberOfArgumentsError(method, args.length);
            return null;
        }

        const points = [];
        for (const node of args) {
            if (node == null || node.token == null || node.token.kind !== TokenType.Identifier) {
                this.status.addError("'" + method.name + "' takes the names of points");
                return null;
            }

            const point = this.resolvePoint(node.token.text);
            if (point == null) {
                return null;
            }

            points.push(point);
        }

        return new BoundPointCall(method, points);
    }

    resolvePoint(pointName) {
        const figure = this.binder.resolveFigure(pointName);
        if (figure == null) {
            this.status.addUnknownIdentifierError(pointName);
            return null;
        }

        if (figure.isPoint !== true) {
            this.status.addFigureIsNotAPointError(pointName);
            return null;
        }

        if (!this.binder.isFigureAllowed(figure)) {
            this.status.addDependencyCycleError(pointName);
            return null;
        }

        this.status.dependencies.push(figure);
        return figure;
    }

    createBinaryExpression(node) {
        const left = this.createExpressionCore(node.children[0]);
        const right = this.createExpressionCore(node.children[1]);
        if (left == null || right == null) {
            return null;
        }

        switch (node.kind) {
            case NodeType.Addition:
                return new BoundBinary(BoundOperator.Add, left, right);
            case NodeType.Subtraction:
                return new BoundBinary(BoundOperator.Subtract, left, right);
            case NodeType.Multiplication:
                return new BoundBinary(BoundOperator.Multiply, left, right);
            case NodeType.Division:
                return new BoundBinary(BoundOperator.Divide, left, right);
            case NodeType.Power:
                return new BoundBinary(BoundOperator.Power, left, right);
            default:
                return null;
        }
    }

    setContext(drawing, isFigureAllowed) {
        this.binder.drawing = drawing;
        this.binder.figureAllowed = isFigureAllowed;
    }
}
