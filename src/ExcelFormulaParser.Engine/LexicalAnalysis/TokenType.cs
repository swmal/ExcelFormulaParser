using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.LexicalAnalysis
{
    public enum TokenType
    {
        Operator,
        OpeningBracket,
        ClosingBracket,
        String,
        StringContent,
        Integer,
        Decimal,
        Function,
        Undefined
    }
}
