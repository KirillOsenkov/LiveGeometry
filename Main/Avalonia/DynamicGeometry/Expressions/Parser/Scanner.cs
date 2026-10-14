using System;
using System.Text;

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
                    case '`':
                        if (!ScanQuotedName())
                        {
                            return;
                        }

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

        /// <summary>
        /// A prime after a name: A', A'' - the image of A, as in class. Part of the name, so
        /// that an expression can say it (renamed to A', a point named in [A.X] left the label
        /// with an error, the text being [A'.X]).
        /// </summary>
        private static bool IsPrime(char ch)
        {
            return ch == '\'' || ch == '′' || ch == '″';
        }

        /// <summary>
        /// Whether an expression can say the name as it is: a letter (or _), then letters,
        /// digits and primes. Any other name it says in backticks (<see cref="NameInExpression"/>).
        /// </summary>
        public static bool IsName(string name)
        {
            if (string.IsNullOrEmpty(name) || !IsLetter(name[0]))
            {
                return false;
            }

            for (int i = 1; i < name.Length; i++)
            {
                if (!IsLetter(name[i]) && !IsDigit(name[i]) && !IsPrime(name[i]))
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// The name as an expression says it: as it is when it is a name (<see cref="IsName"/>)
        /// that no expression reads as something else - pi and e in any case are the
        /// constants, x is a function's variable - and in backticks otherwise, a backtick in it
        /// doubled: `Drag me!`, `e`. What a rename writes, and whatever writes a figure's name
        /// into an expression, so that every name can be said.
        /// </summary>
        public static string NameInExpression(string name)
        {
            // (a part of a figure has no name: `` says nothing, and doesn't compile)
            name ??= "";
            bool reserved = name == "x"
                || name.Equals("pi", StringComparison.OrdinalIgnoreCase)
                || name.Equals("e", StringComparison.OrdinalIgnoreCase);
            if (IsName(name) && !reserved)
            {
                return name;
            }

            return "`" + name.Replace("`", "``") + "`";
        }

        /// <summary>
        /// A name in backticks, from the opening one: what is between them, a doubled backtick
        /// standing for one. False, with the error said, when it is empty or not closed.
        /// </summary>
        private bool ScanQuotedName()
        {
            int start = current;
            var name = new StringBuilder();
            current++;
            while (current < length)
            {
                if (text[current] == '`')
                {
                    if (NextChar() != '`')
                    {
                        current++;
                        if (name.Length == 0)
                        {
                            current = start;
                            ReportError("A name in backticks can't be empty");
                            return false;
                        }

                        AddToken(new Token(name.ToString(), start, current - start, TokenType.Identifier, isQuoted: true));
                        return true;
                    }

                    current++;
                }

                name.Append(text[current]);
                current++;
            }

            current = start;
            ReportError("A name in backticks needs a closing backtick");
            return false;
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
                if (!IsDigit(currentChar) && !IsLetter(currentChar) && !IsPrime(currentChar))
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
