using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.LexicalAnalysis;

namespace ExcelFormulaParser.Engine.Exceptions
{
    public class UnrecognizedTokenException : Exception
    {
        public UnrecognizedTokenException(Token token)
            : base( "Unrecognized token: " + token.Value)
        {

        }
    }
}
