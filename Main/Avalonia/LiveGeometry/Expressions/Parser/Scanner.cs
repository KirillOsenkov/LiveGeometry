using System;
using System.Text;

namespace LiveGeometry;

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
                // (the signs a caption writes, also when one is pasted in)
                case '-':
                case '−':
                // (what a page or a word processor has for a minus: the hyphen, the
                // non-breaking hyphen, the figure dash, the en dash)
                case '\x2010':
                case '\x2011':
                case '\x2012':
                case '\x2013':
                    AddToken(TokenType.Minus);
                    break;
                case '*':
                case '×':
                case '·':
                // (the dot operator, the bullet operator, the asterisk operator, the bullet)
                case '\x22C5':
                case '\x2219':
                case '\x2217':
                case '\x2022':
                    AddToken(TokenType.Multiply);
                    break;
                case '/':
                case '÷':
                // (the division slash, the fraction slash)
                case '\x2215':
                case '\x2044':
                    AddToken(TokenType.Divide);
                    break;
                // √x, ∛x, as in class
                case '√':
                case '∛':
                    AddToken(TokenType.Root);
                    break;
                case '°':
                    AddToken(TokenType.Degree);
                    break;
                case '|':
                    AddToken(TokenType.Pipe);
                    break;
                // ∠ABC, the angle at B, as in class
                case '∠':
                    AddToken(TokenType.Angle);
                    break;
                case '<':
                    if (NextChar() == '=')
                    {
                        AddToken(TokenType.LessOrEqual, length: 2);
                    }
                    else
                    {
                        AddToken(TokenType.Less);
                    }

                    break;
                case '>':
                    if (NextChar() == '=')
                    {
                        AddToken(TokenType.GreaterOrEqual, length: 2);
                    }
                    else
                    {
                        AddToken(TokenType.Greater);
                    }

                    break;
                case '≤':
                case '⩽':
                    AddToken(TokenType.LessOrEqual);
                    break;
                case '≥':
                case '⩾':
                    AddToken(TokenType.GreaterOrEqual);
                    break;
                // = as in class, == as programmers and GeoGebra write it
                case '=':
                    AddToken(TokenType.Equal, length: NextChar() == '=' ? 2 : 1);
                    break;
                case '≠':
                    AddToken(TokenType.NotEqual);
                    break;
                case '!':
                    if (NextChar() != '=')
                    {
                        ReportError("Invalid character");
                        return;
                    }

                    AddToken(TokenType.NotEqual, length: 2);
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
                        int literalStart = current;
                        current++;
                        ScanNumericLiteralAfterPeriod(literalStart, "0.");
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
                    else if (IsSpace(currentChar))
                    {
                        current++;
                    }
                    else if (FractionOf(currentChar) is string fraction)
                    {
                        AddToken(new Token(fraction, current, length: 1, TokenType.Fraction, isQuoted: false));
                        current++;
                    }
                    else if (!ScanSuperscript())
                    {
                        ReportError("Invalid character");
                        return;
                    }
                    break;
            }
        }
    }

    /// <summary>
    /// What a pasted text has between its words besides the space: a tab, a line break, the
    /// non-breaking and the thin spaces of a page, the zero-width ones nobody sees
    /// </summary>
    private static bool IsSpace(char ch)
    {
        return ch == ' '
            || (ch >= '\t' && ch <= '\r')
            || ch == '\x0085'
            || ch == '\x00A0'
            || ch == '\x1680'
            || (ch >= '\x2000' && ch <= '\x200B')
            || ch == '\x2028'
            || ch == '\x2029'
            || ch == '\x202F'
            || ch == '\x205F'
            || ch == '\x2060'
            || ch == '\x3000'
            || ch == '\xFEFF';
    }

    /// <summary>The digit a raised one writes (² is 2), or -1</summary>
    private static int SuperscriptDigit(char ch)
    {
        return ch switch
        {
            '⁰' => 0,
            '¹' => 1,
            '²' => 2,
            '³' => 3,
            >= '⁴' and <= '⁹' => ch - '⁴' + 4,
            _ => -1
        };
    }

    /// <summary>
    /// A raised whole number, x² or x⁻¹: one token that reads as the digits it writes (2, -1)
    /// and takes the raised ones' place in the text. False when there is none here.
    /// </summary>
    private bool ScanSuperscript()
    {
        int start = current;
        var digits = new StringBuilder();
        int index = current;
        if (text[index] == '⁻')
        {
            digits.Append('-');
            index++;
        }

        while (index < length && SuperscriptDigit(text[index]) >= 0)
        {
            digits.Append((char)('0' + SuperscriptDigit(text[index])));
            index++;
        }

        if (digits.Length == 0 || digits[^1] == '-')
        {
            return false;
        }

        current = index;
        AddToken(new Token(digits.ToString(), start, current - start, TokenType.Superscript, isQuoted: false));
        return true;
    }

    /// <summary>The division a fraction's character writes (½ is 1/2), or null for any other character</summary>
    private static string FractionOf(char ch)
    {
        return ch switch
        {
            '½' => "1/2",
            '⅓' => "1/3",
            '⅔' => "2/3",
            '¼' => "1/4",
            '¾' => "3/4",
            '⅕' => "1/5",
            '⅖' => "2/5",
            '⅗' => "3/5",
            '⅘' => "4/5",
            '⅙' => "1/6",
            '⅚' => "5/6",
            '⅐' => "1/7",
            '⅛' => "1/8",
            '⅜' => "3/8",
            '⅝' => "5/8",
            '⅞' => "7/8",
            '⅑' => "1/9",
            '⅒' => "1/10",
            _ => null
        };
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

    /// <summary>Whether a name can begin with the character: a letter or _</summary>
    public static bool IsNameStart(char ch)
    {
        return IsLetter(ch);
    }

    /// <summary>Whether a name can go on with the character: a letter, _, a digit or a prime</summary>
    public static bool IsNamePart(char ch)
    {
        return IsLetter(ch) || IsDigit(ch) || IsPrime(ch);
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
    /// that no expression reads as something else - pi, π and e in any case are the
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
            || name == "π"
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
                    ScanNumericLiteralAfterPeriod(start, text.Substring(start, current - start));
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

    /// <param name="literalStart">Where the number begins in the text: at its first digit, or at the period of .5</param>
    /// <param name="beforePeriod">The number up to its digits after the period, as the parser reads it: "12.", or "0." for .5</param>
    private void ScanNumericLiteralAfterPeriod(int literalStart, string beforePeriod)
    {
        int start = current;

        while (true)
        {
            current++;

            if (current == length || !IsDigit(text[current]))
            {
                AddNumericLiteral(literalStart, beforePeriod + text.Substring(start, current - start));
                return;
            }
        }
    }

    private void AddNumericLiteral(int start)
    {
        AddNumericLiteral(start, text.Substring(start, current - start));
    }

    /// <summary>
    /// The number from <paramref name="start"/> up to here, read as <paramref name="literal"/>:
    /// its token takes its place in the text, also where it reads otherwise (.5 is 0.5).
    /// Worked out from what it reads as, the start of .5 was -1, a character too far left.
    /// </summary>
    private void AddNumericLiteral(int start, string literal)
    {
        AddToken(new Token(literal, start, current - start, TokenType.NumericLiteral, isQuoted: false));
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

    private void AddToken(TokenType tokenType, int length = 1)
    {
        var token = new Token(text.Substring(current, length), current, tokenType);
        AddToken(token);
        current += length;
    }

    private void AddToken(Token token)
    {
        result.Tokens.Add(token);
    }
}
