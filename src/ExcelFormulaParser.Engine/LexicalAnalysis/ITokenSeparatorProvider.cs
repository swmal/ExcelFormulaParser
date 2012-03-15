using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.LexicalAnalysis
{
    public interface ITokenSeparatorProvider
    {
        IDictionary<string, Token> Tokens { get; }

        bool IsOperator(string item);

    }
}
