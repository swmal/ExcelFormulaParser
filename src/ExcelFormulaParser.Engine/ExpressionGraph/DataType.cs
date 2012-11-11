using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public enum DataType
    {
        Integer,
        Decimal,
        String,
        Boolean,
        Date,
        Time,
        Enumerable,
        LookupArray,
        ExcelAddress
    }
}
