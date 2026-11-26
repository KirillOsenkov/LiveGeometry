namespace LiveGeometry;

public enum NodeType
{
    Unknown,
    Negation,
    Addition,
    Subtraction,
    Multiplication,
    Division,
    Power,
    PropertyAccess,
    FunctionCall,
    PoweredCall,
    Constant,
    Fraction,
    Variable,
    Degrees,
    AbsoluteValue,
    Less,
    LessOrEqual,
    Greater,
    GreaterOrEqual,
    Equal,
    NotEqual
}
