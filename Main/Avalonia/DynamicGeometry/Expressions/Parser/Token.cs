namespace DynamicGeometry
{
    public class Token
    {
        public readonly string Text;
        public readonly TokenType Kind;
        public int Start;

        public Token(string text, int start, TokenType kind)
        {
            Text = text;
            Start = start;
            Kind = kind;
        }
    }
}
