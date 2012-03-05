using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.LexicalAnalysis
{
    public class Token
    {
        public Token(string token, TokenType tokenType)
        {
            Value = token;
            TokenType = tokenType;
        }

        public string Value { get; private set; }

        public TokenType TokenType { get; private set; }
    }
}
