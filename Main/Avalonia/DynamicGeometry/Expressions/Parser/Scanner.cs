using System;
using System.Collections.Generic;

namespace DynamicGeometry
{
    public class Scanner
    {
        private string text;
        int current = 0;
        int length = 0;
        private ScanResult result;

        public static ScanResult Scan(string expression)
        {
            var scanner = new Scanner();
            scanner.text = expression;
            scanner.Scan();
            return scanner.result;
        }

        private Scanner()
        {
            result = new ScanResult();
        }

        private void Scan()
        {
            current = 0;
            length = text.Length;

            while (current < length)
            {
                char currentChar = text[current];
                switch (currentChar)
                {
                    case '+':
                        AddToken(TokenType.Plus);
                        break;
                    case '-':
                        AddToken(TokenType.Minus);
                        break;
                    case '*':
                        AddToken(TokenType.Multiply);
                        break;
                    case '/':
                        AddToken(TokenType.Divide);
                        break;
                    case '^':
                        AddToken(TokenType.Power);
                        break;
                    case '(':
                        AddToken(TokenType.OpenParen);
                        break;
                    case ')':
                        AddToken(TokenType.CloseParen);
                        break;
                    case '.':
                        if (IsDigit(NextChar()))
                        {
                            current++;
                            ScanNumericLiteralAfterPeriod("0.");
                        }
                        else
                        {
                            AddToken(TokenType.Dot);
                        }
                        break;
                    case ',':
                        AddToken(TokenType.Comma);
                        break;
                    case ' ':
                        current++;
                        break;
                    default:
                        if (IsDigit(currentChar))
                        {
                            ScanNumericLiteral();
                        }
                        else if (IsLetter(currentChar))
                        {
                            ScanIdentifier();
                        }
                        else
                        {
                            ReportError("Invalid character");
                            return;
                        }
                        break;
                }
            }
        }

        private void ReportError(string error)
        {
            result.Errors.Add(new CompileError() { Text = error, Position = current });
        }

        private static bool IsDigit(char ch)
        {
            return ch >= '0' && ch <= '9';
        }

        private static bool IsLetter(char ch)
        {
            return (ch >= 'a' && ch <= 'z')
                || (ch >= 'A' && ch <= 'Z')
                || ch == '_'
                || char.IsLetter(ch);
        }

        private void ScanNumericLiteral()
        {
            int start = current;

            while (true)
            {
                current++;

                if (current == length)
                {
                    AddNumericLiteral(start);
                    return;
                }

                char currentChar = text[current];
                if (currentChar == '.')
                {
                    if (IsDigit(NextChar()))
                    {
                        current++;
                        ScanNumericLiteralAfterPeriod(text.Substring(start, current - start));
                        return;
                    }
                    else
                    {
                        ReportError("Decimal period must be followed by fractional part (a digit)");
                        current = length;
                        return;
                    }
                }

                if (!IsDigit(currentChar))
                {
                    AddNumericLiteral(start);
                    return;
                }
            }
        }

        private void ScanNumericLiteralAfterPeriod(string beforePeriod)
        {
            int start = current;

            while (true)
            {
                current++;

                if (current == length)
                {
                    AddNumericLiteral(start, beforePeriod);
                    return;
                }

                char currentChar = text[current];
                if (!IsDigit(currentChar))
                {
                    AddNumericLiteral(start, beforePeriod);
                    return;
                }
            }
        }

        private void AddNumericLiteral(int start)
        {
            var token = new Token(text.Substring(start, current - start), start, TokenType.NumericLiteral);
            AddToken(token);
        }

        private void AddNumericLiteral(int start, string prefix)
        {
            var token = new Token(prefix + text.Substring(start, current - start), start - prefix.Length, TokenType.NumericLiteral);
            AddToken(token);
        }

        private char NextChar()
        {
            int index = current + 1;
            if (index < length)
            {
                return text[index];
            }

            return '\0';
        }

        private void ScanIdentifier()
        {
            int start = current;

            while (true)
            {
                current++;
                if (current == length)
                {
                    AddIdentifier(start);
                    return;
                }

                char currentChar = text[current];
                if (!IsDigit(currentChar) && !IsLetter(currentChar))
                {
                    AddIdentifier(start);
                    return;
                }
            }
        }

        private void AddIdentifier(int start)
        {
            var token = new Token(text.Substring(start, current - start), start, TokenType.Identifier);
            AddToken(token);
        }

        private void AddToken(TokenType tokenType)
        {
            var token = new Token(text[current].ToString(), current, tokenType);
            AddToken(token);
            current++;
        }

        private void AddToken(Token token)
        {
            result.Tokens.Add(token);
        }
    }
}
