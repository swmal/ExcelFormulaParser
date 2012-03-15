using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.LexicalAnalysis
{
    public enum TokenType
    {
        Operator,
        Negator,
        OpeningBracket,
        ClosingBracket,
        OpeningEnumerable,
        ClosingEnumerable,
        Enumerable,
        Comma,
        String,
        StringContent,
        Integer,
        Boolean,
        Decimal,
        Function,
        ExcelAddress,
        Unrecognized
    }
}
