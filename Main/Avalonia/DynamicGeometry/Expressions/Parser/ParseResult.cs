using System.Collections.Generic;

namespace DynamicGeometry
{
    public class ParseResult
    {
        public ParseResult()
        {
            Errors = new List<CompileError>();
        }

        public ParseResult(List<CompileError> list)
        {
            Errors = list;
        }

        public Node Root { get; set; }
        public List<CompileError> Errors { get; private set; }
    }
}
