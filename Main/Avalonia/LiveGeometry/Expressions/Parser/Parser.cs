using System.Collections.Generic;

namespace LiveGeometry;

public class Parser
{
    private ParseResult result;
    private IList<Token> tokens;
    private int current;
    private int length;

    private Parser(IList<Token> tokens)
    {
        result = new ParseResult();
        this.tokens = tokens;
        length = tokens.Count;
    }

    public static ParseResult Parse(string expression)
    {
        var scanResult = Scanner.Scan(expression);
        if (scanResult.Errors.Count > 0)
        {
            return new ParseResult(scanResult.Errors);
        }

        var parser = new Parser(scanResult.Tokens);
        parser.Parse();

        return parser.result;
    }

    private Token Current
    {
        get
        {
            if (EOF)
            {
                return null;
            }

            return tokens[current];
        }
    }

    private TokenType CurrentTokenKind
    {
        get
        {
            if (EOF)
            {
                return TokenType.Unknown;
            }

            return Current.Kind;
        }
    }

    private Token Next
    {
        get
        {
            if (current + 1 >= length)
            {
                return null;
            }

            return tokens[current + 1];
        }
    }

    private void Parse()
    {
        Node expression = ParseExpression(0);
        if (current < length)
        {
            ReportError("Unexpected expression ending: " + tokens[current].Text);
        }
        else
        {
            result.Root = expression;
        }
    }

    private Node ParseExpression(int precedence)
    {
        Node leftOperand = null;

        if (current == length)
        {
            ReportError("Expression expected");
            return null;
        }

        if (Current.Kind == TokenType.Minus)
        {
            AdvanceToken();
            if (current == length)
            {
                ReportError("Expected an expression after a negation sign -");
                return null;
            }
            // The minus takes what follows up to the next + - * /, powers included: -x^2
            // is -(x^2), as it is written in class. (It bound tighter than the power, so
            // -x^2 was (-x)^2: the parabola y = -x^2 opened upward.)
            var unaryOperand = ParseExpression(GetPrecedence(NodeType.Multiplication));
            leftOperand = Negation(unaryOperand);
        }
        else
        {
            leftOperand = ParseTerm(precedence);
        }

        while (true)
        {
            if (current == length)
            {
                break;
            }

            // A number right before a name or a parenthesis multiplies it, as in class: 2x,
            // 3sin(x), 2(x + 1), 2pi. Never two names (AB is the distance of A and B), nor
            // a name before a parenthesis (a function's call).
            bool implicitMultiplication = IsImplicitMultiplication();
            if (!implicitMultiplication && !IsOperatorToken(Current.Kind))
            {
                break;
            }

            var operationType = implicitMultiplication ? NodeType.Multiplication : GetExpressionType(Current.Kind);
            var newPrecedence = GetPrecedence(operationType);

            if (newPrecedence < precedence)
            {
                break;
            }

            if (newPrecedence == precedence && !IsRightAssociative(operationType))
            {
                break;
            }

            if (!implicitMultiplication)
            {
                AdvanceToken();
            }

            var rightOperand = ParseExpression(newPrecedence);
            leftOperand = BinaryOperator(operationType, leftOperand, rightOperand);
        }

        return leftOperand;
    }

    /// <summary>A number just read, and a name or a parenthesis right after it</summary>
    private bool IsImplicitMultiplication()
    {
        return current > 0
            && tokens[current - 1].Kind == TokenType.NumericLiteral
            && (Current.Kind == TokenType.Identifier || Current.Kind == TokenType.OpenParen);
    }

    private bool EOF => current >= length;

    private void ReportError(string error)
    {
        int position = length - 1;
        if (!EOF)
        {
            position = Current.Start;
        }
        result.Errors.Add(new CompileError() { Text = error, Position = position });
    }

    private Node BinaryOperator(NodeType operationType, Node leftOperand, Node rightOperand)
    {
        return new Node(operationType, leftOperand, rightOperand);
    }

    private void AdvanceToken()
    {
        current++;
    }

    private bool IsOperatorToken(TokenType tokenType)
    {
        return GetExpressionType(tokenType) switch
        {
            NodeType.Addition
                or NodeType.Subtraction
                or NodeType.Multiplication
                or NodeType.Division
                or NodeType.Power
                or NodeType.Less
                or NodeType.LessOrEqual
                or NodeType.Greater
                or NodeType.GreaterOrEqual
                or NodeType.Equal
                or NodeType.NotEqual => true,
            _ => false
        };
    }

    /// <summary>A term and the degree signs after it: 30°, angle(A, B, C)°, (a + b)°</summary>
    private Node ParseTerm(int precedence)
    {
        var term = ParsePrimary(precedence);
        while (term != null && CurrentTokenKind == TokenType.Degree)
        {
            // (the sign is the node's token, to be colored as pi is)
            var degrees = new Node(NodeType.Degrees, Current);
            degrees.Children.Add(term);
            term = degrees;
            AdvanceToken();
        }

        return term;
    }

    private Node ParsePrimary(int precedence)
    {
        Node result = null;

        switch (Current.Kind)
        {
            case TokenType.NumericLiteral:
                result = Constant();
                AdvanceToken();
                break;
            case TokenType.Identifier:
                var next = Next;
                if (next == null)
                {
                    result = Variable();
                    AdvanceToken();
                    break;
                }

                if (next.Kind == TokenType.Dot)
                {
                    return ParsePropertyAccess();
                }

                // (a name in backticks is a figure's, never a function's)
                if (next.Kind == TokenType.OpenParen && !Current.IsQuoted)
                {
                    return ParseFunctionCall();
                }

                result = Variable();
                AdvanceToken();
                break;
            case TokenType.OpenParen:
                AdvanceToken();
                result = ParseExpression(0);
                AdvanceToken(TokenType.CloseParen);
                break;
            case TokenType.Pipe:
                // |x|: what is between the bars, which end it as a ) ends a parenthesis
                AdvanceToken();
                result = new Node(NodeType.AbsoluteValue, ParseExpression(0));
                AdvanceToken(TokenType.Pipe);
                break;
            default:
                ReportError("Expected a number, a variable, a function or a parenthesized expression");
                break;
        }

        return result;
    }

    private Node ParseFunctionCall()
    {
        var result = new Node(NodeType.FunctionCall, Current);
        AdvanceToken();
        AdvanceToken(TokenType.OpenParen);
        AddArgument(result, ParseExpression());
        while (CurrentTokenKind == TokenType.Comma)
        {
            AdvanceToken();
            AddArgument(result, ParseExpression());
        }
        AdvanceToken(TokenType.CloseParen);
        return result;
    }

    private void AddArgument(Node functionCall, Node argument)
    {
        functionCall.Children.Add(argument);
    }

    private Node ParseExpression()
    {
        return ParseExpression(0);
    }

    private void AdvanceToken(TokenType tokenType)
    {
        if (current == length || Current.Kind != tokenType)
        {
            ReportError("Expected " + (tokenType == TokenType.Pipe ? "| to close |" : tokenType.ToString()));
            return;
        }

        AdvanceToken();
    }

    private Node ParsePropertyAccess()
    {
        var left = Variable();
        AdvanceToken(TokenType.Identifier);
        AdvanceToken(TokenType.Dot);
        var right = Variable();
        AdvanceToken(TokenType.Identifier);
        return PropertyAccess(left, right);
    }

    private Node PropertyAccess(Node left, Node right)
    {
        return new Node(NodeType.PropertyAccess, left, right);
    }

    private Node Variable()
    {
        if (EOF)
        {
            ReportError("Variable expected");
            return null;
        }

        return new Node(NodeType.Variable, Current);
    }

    private Node Constant()
    {
        if (EOF)
        {
            ReportError("Constant expected");
            return null;
        }

        return new Node(NodeType.Constant, Current);
    }

    private Node Negation(Node unaryOperand)
    {
        return new Node(NodeType.Negation, unaryOperand);
    }

    public static bool IsRightAssociative(NodeType operation)
    {
        return operation == NodeType.Power;
    }

    public static NodeType GetExpressionType(TokenType tokenType)
    {
        switch (tokenType)
        {
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
            case TokenType.Less:
                return NodeType.Less;
            case TokenType.LessOrEqual:
                return NodeType.LessOrEqual;
            case TokenType.Greater:
                return NodeType.Greater;
            case TokenType.GreaterOrEqual:
                return NodeType.GreaterOrEqual;
            case TokenType.Equal:
                return NodeType.Equal;
            case TokenType.NotEqual:
                return NodeType.NotEqual;
            default:
                return NodeType.Unknown;
        }
    }

    /// <summary>A comparison: what if(...) asks, 1 when it holds and 0 when it doesn't</summary>
    public static bool IsComparison(NodeType nodeType)
    {
        return nodeType is NodeType.Less
            or NodeType.LessOrEqual
            or NodeType.Greater
            or NodeType.GreaterOrEqual
            or NodeType.Equal
            or NodeType.NotEqual;
    }

    public static int GetPrecedence(NodeType nodeType)
    {
        switch (nodeType)
        {
            case NodeType.Constant:
                return 0;
            case NodeType.Less:
            case NodeType.LessOrEqual:
            case NodeType.Greater:
            case NodeType.GreaterOrEqual:
            case NodeType.Equal:
            case NodeType.NotEqual:
                return 1;
            case NodeType.Addition:
            case NodeType.Subtraction:
                return 2;
            case NodeType.Multiplication:
            case NodeType.Division:
                return 3;
            case NodeType.Power:
                return 4;
            case NodeType.Negation:
                return 5;
            case NodeType.PropertyAccess:
                return 6;
            case NodeType.FunctionCall:
                return 7;
        }

        return 0;
    }
}
