#region License
/* **********************************************************************************
 * Copyright (c) Roman Ivantsov
 * This source code is subject to terms and conditions of the MIT License
 * for Irony. A copy of the license can be found in the License.txt file
 * at the root of this distribution. 
 * By using this source code in any fashion, you are agreeing to be bound by the terms of the 
 * MIT License.
 * You must not remove this notice from this software.
 * **********************************************************************************/
#endregion

// Irony binary is taken from http://irony.codeplex.com
// This grammar is based on the ExpressionEvaluatorGrammar from Irony.Samples

using Irony.Parsing;

namespace DynamicGeometry
{
    [Language("Expression", "1.0", "Dynamic geometry expression evaluator")]
    public class ExpressionGrammar : Irony.Parsing.Grammar
    {
        public ExpressionGrammar()
        {
            this.GrammarComments = @"Arithmetical expressions for dynamic geometry.";

            // 1. Terminals
            var number = new NumberLiteral("number");
            var identifier = new IdentifierTerminal("identifier");

            // 2. Non-terminals
            var Expr = new NonTerminal("Expr");
            var Term = new NonTerminal("Term");
            var BinExpr = new NonTerminal("BinExpr");
            var ParExpr = new NonTerminal("ParExpr");
            var UnExpr = new NonTerminal("UnExpr");
            var UnOp = new NonTerminal("UnOp");
            var BinOp = new NonTerminal("BinOp", "operator");
            var PropertyAccess = new NonTerminal("PropertyAccess");
            var FunctionCall = new NonTerminal("FunctionCall");
            var CommaSeparatedIdentifierList = new NonTerminal("PointArgumentList");
            var ArgumentList = new NonTerminal("ArgumentList");

            // 3. BNF rules
            Expr.Rule = Term | UnExpr | BinExpr;
            Term.Rule = number | identifier | ParExpr | FunctionCall | PropertyAccess;
            UnExpr.Rule = UnOp + Term;
            UnOp.Rule = ToTerm("-");
            BinExpr.Rule = Expr + BinOp + Expr;
            BinOp.Rule = ToTerm("+") | "-" | "*" | "/" | "^";
            PropertyAccess.Rule = identifier + "." + identifier;
            FunctionCall.Rule = identifier + "(" + ArgumentList + ")";
            ArgumentList.Rule = Expr | CommaSeparatedIdentifierList;
            ParExpr.Rule = "(" + Expr + ")";
            CommaSeparatedIdentifierList.Rule = MakePlusRule(CommaSeparatedIdentifierList, ToTerm(","), identifier);
            
            this.Root = Expr;

            // 4. Operators precedence
            RegisterOperators(1, "+", "-");
            RegisterOperators(2, "*", "/");
            RegisterOperators(3, Associativity.Right, "^");

            MarkPunctuation("(", ")", ".", ",");
            MarkTransient(Term, Expr, BinOp, UnOp, ParExpr, ArgumentList, CommaSeparatedIdentifierList);
        }

        public static ExpressionGrammar Instance = new ExpressionGrammar();
    }
}
