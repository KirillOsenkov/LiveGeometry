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

    /// <summary>A number just read (½ is one), and a name, a parenthesis or a root right after it</summary>
    private bool IsImplicitMultiplication()
    {
        return current > 0
            && tokens[current - 1].Kind is TokenType.NumericLiteral or TokenType.Fraction
            && Current.Kind is TokenType.Identifier or TokenType.OpenParen or TokenType.Angle or TokenType.Root;
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

    /// <summary>
    /// A term and what is written after it: degree signs (30°, angle(A, B, C)°, (a + b)°)
    /// and raised powers (x², AB², (a + b)²)
    /// </summary>
    private Node ParseTerm(int precedence)
    {
        var term = ParsePrimary(precedence);
        while (term != null)
        {
            if (CurrentTokenKind == TokenType.Degree)
            {
                // (the sign is the node's token, to be colored as pi is)
                var degrees = new Node(NodeType.Degrees, Current);
                degrees.Children.Add(term);
                term = degrees;
            }
            else if (CurrentTokenKind == TokenType.Superscript)
            {
                // x² is x^2 with the 2 right on the x: 5π² is five of π², -x² is -(x²),
                // 2^x² is 2^(x²). (The raised digits are the exponent's token, to be
                // colored as a number is.)
                term = new Node(NodeType.Power, term, new Node(NodeType.Constant, Current));
            }
            else
            {
                break;
            }

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

                // 2½ is two and a half: the fraction right after the number, nothing between
                if (CurrentTokenKind == TokenType.Fraction && Current.Start == tokens[current - 1].Start + tokens[current - 1].Length)
                {
                    result = new Node(NodeType.Addition, result, new Node(NodeType.Fraction, Current));
                    AdvanceToken();
                }

                break;
            case TokenType.Fraction:
                result = new Node(NodeType.Fraction, Current);
                AdvanceToken();
                break;
            case TokenType.Root:
                return ParseRoot();
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

                // sin²(x): the power on the function's name
                if (next.Kind == TokenType.Superscript
                    && !Current.IsQuoted
                    && current + 2 < length
                    && tokens[current + 2].Kind == TokenType.OpenParen)
                {
                    return ParsePoweredCall(next, powerTokenCount: 1);
                }

                // sin^2(x), sin^-1(x): the same typed with a caret
                if (IsCaretPoweredCall(out var caretPower, out int caretTokenCount))
                {
                    return ParsePoweredCall(caretPower, caretTokenCount);
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
            case TokenType.Angle:
                return ParseAngle();
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
        ParseArguments(result);
        return result;
    }

    /// <summary>
    /// sin²(x), sin⁻¹(x): a call with a power written on the function's name, as in class.
    /// The call and the power, for the binder to read: sin²(x) is (sin(x))², and sin⁻¹ is
    /// arcsin, not one over the sine.
    /// </summary>
    /// <param name="power">The power as a number's token: the raised digits, or what follows the caret</param>
    /// <param name="powerTokenCount">How many tokens the power takes between the name and the parenthesis</param>
    private Node ParsePoweredCall(Token power, int powerTokenCount)
    {
        var call = new Node(NodeType.FunctionCall, Current);
        current += 1 + powerTokenCount;
        ParseArguments(call);
        return new Node(NodeType.PoweredCall, call, new Node(NodeType.Constant, power));
    }

    /// <summary>
    /// sin^2(x), sin^-1(x): the name of a function here, then a caret, a number (a minus
    /// before it or not) and a parenthesis. Only for a name the language has a function by:
    /// a^2(x + 1) is a product, a squared times the sum. The power comes out as one token,
    /// the minus in it, which takes the place of both in the text.
    /// </summary>
    private bool IsCaretPoweredCall(out Token power, out int tokenCount)
    {
        power = null;
        tokenCount = 0;
        if (Current.IsQuoted
            || Next?.Kind != TokenType.Power
            || (FunctionCatalog.Find(Current.Text) == null && ExpressionReflection.LogBaseOf(Current.Text) == null))
        {
            return false;
        }

        int number = current + 2;
        Token minus = null;
        if (number < length && tokens[number].Kind == TokenType.Minus)
        {
            minus = tokens[number];
            number++;
        }

        if (number + 1 >= length
            || tokens[number].Kind != TokenType.NumericLiteral
            || tokens[number + 1].Kind != TokenType.OpenParen)
        {
            return false;
        }

        power = tokens[number];
        if (minus != null)
        {
            power = new Token(
                "-" + power.Text,
                minus.Start,
                power.Start + power.Length - minus.Start,
                TokenType.NumericLiteral,
                isQuoted: false);
        }

        tokenCount = number - current;
        return true;
    }

    /// <summary>The arguments of a call, from its ( to its )</summary>
    private void ParseArguments(Node functionCall)
    {
        AdvanceToken(TokenType.OpenParen);
        AddArgument(functionCall, ParseExpression());
        while (CurrentTokenKind == TokenType.Comma)
        {
            AdvanceToken();
            AddArgument(functionCall, ParseExpression());
        }

        AdvanceToken(TokenType.CloseParen);
    }

    /// <summary>
    /// ∠ABC: a call of angle (the sign is one of its names) with the three points' names run
    /// together into one, as angle(ABC) is; ∠(A, B, C) is angle(A, B, C)
    /// </summary>
    private Node ParseAngle()
    {
        if (Next?.Kind == TokenType.OpenParen)
        {
            return ParseFunctionCall();
        }

        var result = new Node(NodeType.FunctionCall, Current);
        AdvanceToken();
        if (CurrentTokenKind != TokenType.Identifier)
        {
            ReportError("Expected the names of three points after ∠: ∠ABC");
            return null;
        }

        AddArgument(result, Variable());
        AdvanceToken();
        return result;
    }

    /// <summary>
    /// √x, ∛x: a call of sqrt or cbrt (the sign is one of its names). √(x + 1) is
    /// sqrt(x + 1), what the parentheses hold; without them the root is of what follows up
    /// to the next + - * /, powers included: √x² is √(x²), and 2√3 is two of √3.
    /// </summary>
    private Node ParseRoot()
    {
        if (Next?.Kind == TokenType.OpenParen)
        {
            return ParseFunctionCall();
        }

        var result = new Node(NodeType.FunctionCall, Current);
        AdvanceToken();
        AddArgument(result, ParseExpression(GetPrecedence(NodeType.Power)));
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
