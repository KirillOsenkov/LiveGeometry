using System;
using System.Collections.Generic;

namespace DynamicGeometry
{
    public class Node
    {
        public readonly IList<Node> Children;
        public readonly NodeType Kind;
        public readonly Token Token;

        public Node(NodeType nodeType, params Node[] children)
        {
            this.Kind = nodeType;
            this.Children = new List<Node>(children);
        }

        public Node(NodeType nodeType, Token token)
        {
            this.Kind = nodeType;
            this.Token = token;
            this.Children = new List<Node>();
        }
    }
}
