namespace DynamicGeometry
{
    public class Token
    {
        public readonly string Text;
        public readonly TokenType Kind;
        public int Start;

        /// <summary>How much of the expression's text the token takes: more than its text for a name in backticks</summary>
        public readonly int Length;

        /// <summary>A name in backticks (`Drag me!`): the figure called exactly that, never anything else the name could read as</summary>
        public readonly bool IsQuoted;

        public Token(string text, int start, TokenType kind)
            : this(text, start, text.Length, kind, isQuoted: false)
        {
        }

        public Token(string text, int start, int length, TokenType kind, bool isQuoted)
        {
            Text = text;
            Start = start;
            Length = length;
            Kind = kind;
            IsQuoted = isQuoted;
        }
    }
}
