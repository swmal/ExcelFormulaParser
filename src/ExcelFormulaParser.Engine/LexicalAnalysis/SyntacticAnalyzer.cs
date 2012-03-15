using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Exceptions;

namespace ExcelFormulaParser.Engine.LexicalAnalysis
{
    public class SyntacticAnalyzer : ISyntacticAnalyzer
    {
        private class AnalyzingContext
        {
            public int NumberOfOpenedParentheses { get; set; }
            public int NumberOfClosedParentheses { get; set; }
            public int OpenedStrings { get; set; }
            public int ClosedStrings { get; set; }
            public bool IsInString { get; set; }
        }
        public void Analyze(IEnumerable<Token> tokens)
        {
            var context = new AnalyzingContext();
            foreach (var token in tokens)
            {
                if (token.TokenType == TokenType.Unrecognized)
                {
                    throw new UnrecognizedTokenException(token);
                }
                EnsureParenthesesAreWellFormed(token, context);
                EnsureStringsAreWellFormed(token, context);
            }
            Validate(context);
        }

        private static void Validate(AnalyzingContext context)
        {
            if (context.NumberOfOpenedParentheses != context.NumberOfClosedParentheses)
            {
                throw new FormatException("Number of opened and closed parentheses does not match");
            }
            if (context.OpenedStrings != context.ClosedStrings)
            {
                throw new FormatException("Unterminated string");
            }
        }

        private void EnsureParenthesesAreWellFormed(Token token, AnalyzingContext context)
        {
            if (token.TokenType == TokenType.OpeningBracket)
            {
                context.NumberOfOpenedParentheses++;
            }
            else if (token.TokenType == TokenType.ClosingBracket)
            {
                context.NumberOfClosedParentheses++;
            }
        }

        private void EnsureStringsAreWellFormed(Token token, AnalyzingContext context)
        {
            if (!context.IsInString && token.TokenType == TokenType.String)
            {
                context.IsInString = true;
                context.OpenedStrings++;
            }
            else if (context.IsInString && token.TokenType == TokenType.String)
            {
                context.IsInString = false;
                context.ClosedStrings++;
            }
        }
    }
}
