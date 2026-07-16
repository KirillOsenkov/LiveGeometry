using System;
using System.Collections.Generic;

namespace DynamicGeometry
{
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
                var unaryOperand = ParseExpression(GetPrecedence(NodeType.Negation));
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

                if (!IsOperatorToken(Current.Kind))
                {
                    break;
                }

                var operationType = GetExpressionType(Current.Kind);
                var newPrecedence = GetPrecedence(operationType);

                if (newPrecedence < precedence)
                {
                    break;
                }

                if (newPrecedence == precedence && !IsRightAssociative(operationType))
                {
                    break;
                }

                AdvanceToken();
                var rightOperand = ParseExpression(newPrecedence);
                leftOperand = BinaryOperator(operationType, leftOperand, rightOperand);
            }

            return leftOperand;
        }

        private bool EOF
        {
            get
            {
                return current >= length;
            }
        }

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
            switch (tokenType)
            {
                case TokenType.Plus:
                case TokenType.Minus:
                case TokenType.Multiply:
                case TokenType.Divide:
                case TokenType.Power:
                    return true;
                default:
                    return false;
            }
        }

        private Node ParseTerm(int precedence)
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

                    if (next.Kind == TokenType.OpenParen)
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
                ReportError("Expected " + tokenType.ToString());
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
                default:
                    return NodeType.Unknown;
            }
        }

        public static int GetPrecedence(NodeType nodeType)
        {
            switch (nodeType)
            {
                case NodeType.Constant:
                    return 0;
                case NodeType.Addition:
                case NodeType.Subtraction:
                    return 1;
                case NodeType.Multiplication:
                case NodeType.Division:
                    return 2;
                case NodeType.Power:
                    return 3;
                case NodeType.Negation:
                    return 4;
                case NodeType.PropertyAccess:
                    return 5;
                case NodeType.FunctionCall:
                    return 6;
            }

            return 0;
        }
    }
}
