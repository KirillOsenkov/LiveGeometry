using System;

namespace DynamicGeometry
{
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
        Constant,
        Variable
    }
}
