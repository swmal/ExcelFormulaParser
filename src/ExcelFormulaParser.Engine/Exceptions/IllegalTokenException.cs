using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.LexicalAnalysis;

namespace ExcelFormulaParser.Engine.Exceptions
{
    public class IllegalTokenException : Exception
    {
        public IllegalTokenException(Token token)
            : base( "Illegal token: " + token.Value)
        {

        }
    }
}
