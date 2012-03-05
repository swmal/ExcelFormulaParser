using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.LexicalAnalysis
{
    public class SyntacticAnalyzer : ISyntacticAnalyzer
    {
        public void Analyze(IEnumerable<Token> tokens)
        {
            EnsureParenthesesAreWellFormed(tokens);
            EnsureStringsAreWellFormed(tokens);
        }

        private void EnsureParenthesesAreWellFormed(IEnumerable<Token> tokens)
        {
            int nOpening = 0, nClosing = 0;
            foreach (var token in tokens)
            {
                if (token.TokenType == TokenType.OpeningBracket)
                {
                    nOpening++;
                }
                else if (token.TokenType == TokenType.ClosingBracket)
                {
                    nClosing++;
                }
            }
            if (nOpening != nClosing)
            {
                throw new FormatException("Number of opened and closed parentheses does not match");
            }
        }

        private void EnsureStringsAreWellFormed(IEnumerable<Token> tokens)
        {
            int nOpened = 0, nClosed = 0;
            bool isInString = false;
            foreach (var token in tokens)
            {
                if (!isInString && token.TokenType == TokenType.String)
                {
                    isInString = true;
                    nOpened++;
                }
                else if (isInString && token.TokenType == TokenType.String)
                {
                    isInString = false;
                    nClosed++;
                }
            }
            if (nOpened != nClosed)
            {
                throw new FormatException("Unterminated string");
            }
        }
    }
}
