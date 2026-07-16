using System.Collections.Generic;

namespace DynamicGeometry
{
    public class ScanResult
    {
        public List<Token> Tokens { get; private set; }
        public List<CompileError> Errors { get; private set; }

        public ScanResult()
        {
            Tokens = new List<Token>();
            Errors = new List<CompileError>();
        }
    }
}
