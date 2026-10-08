// Port of Main/Avalonia/DynamicGeometry/Expressions/BoundExpression.cs: an expression with
// its names resolved, evaluating itself (the OwnTree strategy, the only one here). Every
// value is a double.

class BoundExpression {
    /** parameter: the x of a function; a plain expression ignores it */
    evaluate(parameter) {
        return NaN;
    }

    toExpressionDelegate() {
        return () => this.evaluate(0);
    }

    toFunctionDelegate() {
        return parameter => this.evaluate(parameter);
    }
}

class BoundConstant extends BoundExpression {
    constructor(value) {
        super();
        this.value = value;
    }

    evaluate(parameter) {
        return this.value;
    }
}

/** The x of a function */
class BoundParameter extends BoundExpression {
    evaluate(parameter) {
        return parameter;
    }
}

/** A number of the drawing (a slider, a Number) named in the expression: its value */
class BoundNumber extends BoundExpression {
    constructor(number) {
        super();
        this.number = number;
    }

    evaluate(parameter) {
        return this.number.value;
    }
}

/** Two points' names run together (AB): their distance */
class BoundDistance extends BoundExpression {
    constructor(first, second) {
        super();
        this.first = first;
        this.second = second;
    }

    evaluate(parameter) {
        return this.first.coordinates.distance(this.second.coordinates);
    }
}

/** A property of a figure (A.X, AB.Length, a.Value): a number */
class BoundProperty extends BoundExpression {
    constructor(figure, property) {
        super();
        this.figure = figure;
        this.property = property;
    }

    evaluate(parameter) {
        const value = this.figure[this.property];
        return typeof value === "number" ? value : NaN;
    }
}

/** A function of numbers (sin, max, atan2, round): each argument is an expression */
class BoundCall extends BoundExpression {
    constructor(method, args) {
        super();
        this.method = method;
        this.arguments = args;
    }

    evaluate(parameter) {
        const values = this.arguments.map(argument => argument.evaluate(parameter));
        return this.method.invoke(...values);
    }
}

/** A function of points (dist, ang, area), called with the names of points */
class BoundPointCall extends BoundExpression {
    constructor(method, points) {
        super();
        this.method = method;
        this.points = points;
    }

    evaluate(parameter) {
        if (this.method.takesArray) {
            return this.method.invoke(this.points);
        }

        return this.method.invoke(...this.points.map(point => point.coordinates));
    }
}

class BoundNegation extends BoundExpression {
    constructor(operand) {
        super();
        this.operand = operand;
    }

    evaluate(parameter) {
        return -this.operand.evaluate(parameter);
    }
}

const BoundOperator = {
    Add: "Add",
    Subtract: "Subtract",
    Multiply: "Multiply",
    Divide: "Divide",
    Power: "Power"
};

class BoundBinary extends BoundExpression {
    constructor(operator, left, right) {
        super();
        this.operator = operator;
        this.left = left;
        this.right = right;
    }

    evaluate(parameter) {
        const left = this.left.evaluate(parameter);
        const right = this.right.evaluate(parameter);
        switch (this.operator) {
            case BoundOperator.Add:
                return left + right;
            case BoundOperator.Subtract:
                return left - right;
            case BoundOperator.Multiply:
                return left * right;
            case BoundOperator.Divide:
                return left / right;
            default:
                return Math.pow(left, right);
        }
    }
}
