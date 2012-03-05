using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.LexicalAnalysis
{
    public interface ITokenFactory
    {
        Token Create(IEnumerable<Token> tokens, string token);
    }
}
